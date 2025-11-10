import os, re
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename
from datetime import datetime, date
import pandas as pd

from models import db, Employee, Attendance
import config

UPLOAD_EXTENSIONS = ['.xlsx', '.xls', '.csv']

def create_app():
    app = Flask(__name__)
    app.config.from_object('config')
    db.init_app(app)

    with app.app_context():
        # ensure tables exist if not using migrations
        db.create_all()

    @app.route('/', methods=['GET', 'POST'])
    def index():
        if request.method == 'POST':
            # handle file upload
            f = request.files.get('file')
            if not f:
                flash('No file uploaded', 'warning')
                return redirect(url_for('index'))

            filename = secure_filename(f.filename)
            if not filename:
                flash('Invalid filename', 'warning')
                return redirect(url_for('index'))

            ext = os.path.splitext(filename)[1].lower()
            if ext not in UPLOAD_EXTENSIONS:
                flash('Unsupported file type', 'danger')
                return redirect(url_for('index'))

            save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            f.save(save_path)
            rows_added = process_upload(save_path, filename)
            flash(f'Processed file. Rows added/updated: {rows_added}', 'success')
            return redirect(url_for('index'))

        # GET: show recap optionally filtered by dates
        start = request.args.get('start')
        end = request.args.get('end')
        try:
            start_date = datetime.strptime(start, '%Y-%m-%d').date() if start else None
            end_date = datetime.strptime(end, '%Y-%m-%d').date() if end else None
        except Exception:
            start_date = end_date = None

        recap_df = build_recap(start_date, end_date)
        table_html = recap_df.to_html(classes='table table-sm table-striped', index=False) if not recap_df.empty else '<p>Aucune donnée</p>'
        return render_template('index.html', table_html=table_html, start=start, end=end)

    @app.route('/export', methods=['GET'])
    def export():
        start = request.args.get('start')
        end = request.args.get('end')
        try:
            start_date = datetime.strptime(start, '%Y-%m-%d').date() if start else None
            end_date = datetime.strptime(end, '%Y-%m-%d').date() if end else None
        except Exception:
            start_date = end_date = None
        recap_df = build_recap(start_date, end_date)
        if recap_df.empty:
            flash('Aucune donnée à exporter', 'warning')
            return redirect(url_for('index'))

        out_path = os.path.join(app.config['EXPORT_FOLDER'], f'recap_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx')
        recap_df.to_excel(out_path, index=False)
        return send_file(out_path, as_attachment=True, download_name=os.path.basename(out_path))

    return app

def normalize_col(name):
    if not isinstance(name, str):
        return ''
    return re.sub(r'[^A-Za-z0-9]', '', name).strip().lower()

def possible_status_cols(cols):
    # return mapping from normalized name to canonical
    mapping = {}
    for c in cols:
        n = normalize_col(c)
        if n in ['present', 'présent','présent'.lower(), 'prsent']:
            mapping[c] = 'present'
        elif n in ['absent']:
            mapping[c] = 'absent'
        elif 'cong' in n:
            mapping[c] = 'cong'
        elif 'tour' in n and 'rep' in n:
            mapping[c] = 'tour_rep'
        elif 'repos' in n or 'med' in n:
            mapping[c] = 'repos_med'
        elif 'sans' in n and 'ph' in n:
            mapping[c] = 'sans_ph'
    return mapping

def parse_date_from_filename(filename):
    # try YYYYMMDD or DDMMYYYY
    m = re.search(r'(20\d{2}[01]\d[0-3]\d)', filename)
    if m:
        try:
            return datetime.strptime(m.group(1), '%Y%m%d').date()
        except:
            pass
    m = re.search(r'(\b\d{2}[\-_/]?\d{2}[\-_/]?\d{4}\b)', filename)
    if m:
        s = re.sub(r'[^0-9]', '', m.group(1))
        try:
            return datetime.strptime(s, '%d%m%Y').date()
        except:
            pass
    return None

def process_upload(path, filename):
    # read file with pandas (try excel, then csv)
    try:
        df = pd.read_excel(path, engine='openpyxl')
    except Exception:
        try:
            df = pd.read_excel(path)
        except Exception:
            df = pd.read_csv(path)

    # normalize columns: lower, strip
    df.columns = [str(c).strip() for c in df.columns]

    # map expected columns (case-insensitive)
    colmap = {normalize_col(c): c for c in df.columns}

    # required fields
    if not any(k in colmap for k in ['matricule', 'id']):
        raise ValueError('Fichier sans colonne Matricule / id')

    # find date for rows
    date_col = None
    for k in colmap:
        if k in ['date']:
            date_col = colmap[k]
            break

    file_date = None
    if date_col is None:
        file_date = parse_date_from_filename(filename) or date.today()

    rows = 0
    with db.session.begin():
        for _, row in df.iterrows():
            try:
                matricule = None
                # find matricule key
                for key in ['matricule', 'id']:
                    if key in colmap:
                        matricule = str(row[colmap[key]]).strip()
                        break
                if not matricule or matricule.lower().startswith('nan'):
                    continue

                # lookup or create employee
                emp = Employee.query.filter_by(matricule=matricule).first()
                if not emp:
                    emp = Employee(matricule=matricule)
                    db.session.add(emp)

                # optional fields
                for fld in [('nom','nom'), ('prenom','prenom'), ('poste','poste'), ('site','site'), ('affaire','affaire')]:
                    nkey = fld[0]
                    # find corresponding column in df by normalized name
                    for k, orig in colmap.items():
                        if k == nkey:
                            val = row[orig]
                            if pd.notna(val):
                                setattr(emp, fld[1], str(val).strip())

                # determine the attendance date for this row
                row_date = None
                if date_col:
                    raw = row[date_col]
                    if pd.isna(raw):
                        row_date = file_date or date.today()
                    elif isinstance(raw, (datetime, pd.Timestamp)):
                        row_date = raw.date()
                    else:
                        # attempt parse
                        try:
                            row_date = pd.to_datetime(raw).date()
                        except:
                            row_date = file_date or date.today()
                else:
                    row_date = file_date or date.today()

                # map status columns
                status_map = possible_status_cols(df.columns)
                # default all zero
                vals = dict(present=0, absent=0, cong=0, tour_rep=0, repos_med=0, sans_ph=0)
                for orig_col, canon in status_map.items():
                    try:
                        v = row[orig_col]
                        if pd.isna(v):
                            continue
                        # treat as boolean or numeric
                        if isinstance(v, str):
                            if v.strip().lower() in ['x','1','y','yes','présent','present','p']:
                                vals[canon] = 1
                        else:
                            try:
                                if int(v) != 0:
                                    vals[canon] = 1
                            except:
                                # fallback: any non-null counts as 1
                                vals[canon] = 1
                    except Exception:
                        continue

                # ensure at least one status nonzero; if none, skip
                if sum(vals.values()) == 0:
                    # try to infer from a special column named 'présé' or 'présent' variations
                    pass

                # upsert attendance row (unique per employee/date)
                att = Attendance.query.filter_by(employee=emp, date=row_date).first()
                if not att:
                    att = Attendance(employee=emp, date=row_date, **vals)
                    db.session.add(att)
                else:
                    # increment if new file indicates presence (we consider storing raw - but here we overwrite with OR)
                    att.present = max(att.present, vals['present'])
                    att.absent = max(att.absent, vals['absent'])
                    att.cong = max(att.cong, vals['cong'])
                    att.tour_rep = max(att.tour_rep, vals['tour_rep'])
                    att.repos_med = max(att.repos_med, vals['repos_med'])
                    att.sans_ph = max(att.sans_ph, vals['sans_ph'])

                rows += 1
            except Exception as e:
                print('row error', e)
                continue
    return rows

def build_recap(start_date=None, end_date=None):
    # query and aggregate
    q = Attendance.query.join(Employee).with_entities(
        Employee.matricule, Employee.nom, Employee.prenom, Employee.poste, Employee.site, Employee.affaire,
        db.func.sum(Attendance.present).label('Présent'),
        db.func.sum(Attendance.absent).label('Absent'),
        db.func.sum(Attendance.cong).label('CONG'),
        db.func.sum(Attendance.tour_rep).label('Tour_rep'),
        db.func.sum(Attendance.repos_med).label('Repos_med'),
        db.func.sum(Attendance.sans_ph).label('Sans_ph')
    ).group_by(Employee.id)

    if start_date:
        q = q.filter(Attendance.date >= start_date)
    if end_date:
        q = q.filter(Attendance.date <= end_date)

    rows = q.all()
    import pandas as pd
    if not rows:
        return pd.DataFrame()
    data = []
    for r in rows:
        data.append({
            'Matricule': r[0],
            'Nom': r[1] or '',
            'Prénom': r[2] or '',
            'Poste': r[3] or '',
            'Site': r[4] or '',
            'Affaire': r[5] or '',
            'Présent': int(r[6] or 0),
            'Absent': int(r[7] or 0),
            'CONG': int(r[8] or 0),
            'Tour_rep': int(r[9] or 0),
            'Repos_med': int(r[10] or 0),
            'Sans_ph': int(r[11] or 0)
        })
    df = pd.DataFrame(data)
    return df

if __name__ == '__main__':
    app = create_app()
    app.run(debug=True)
