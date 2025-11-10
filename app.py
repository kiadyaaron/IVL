import os
import re
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename
from datetime import datetime, date
import pandas as pd
from models import db, Employee, Attendance
import config

# Excel styling
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

UPLOAD_EXTENSIONS = ['.xlsx', '.xls', '.csv']


def create_app():
    app = Flask(__name__)
    app.config.from_object('config')
    db.init_app(app)

    os.makedirs(app.config.get('EXPORT_FOLDER', os.path.join(os.path.dirname(__file__), 'exports')), exist_ok=True)
    os.makedirs(app.config.get('UPLOAD_FOLDER', os.path.join(os.path.dirname(__file__), 'uploads')), exist_ok=True)

    with app.app_context():
        db.create_all()

    @app.route('/', methods=['GET', 'POST'])
    def index():
        if request.method == 'POST':
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
            try:
                rows_added = process_upload(save_path, filename)
                flash(f'Processed file. Rows added/updated: {rows_added}', 'success')
            except Exception as e:
                flash(f'Erreur lors du traitement : {e}', 'danger')
            return redirect(url_for('index'))

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

        q = Attendance.query.join(Employee)
        if start_date:
            q = q.filter(Attendance.date >= start_date)
        if end_date:
            q = q.filter(Attendance.date <= end_date)

        rows = q.with_entities(
            Attendance.date,
            Employee.matricule, Employee.nom, Employee.prenom, Employee.poste, Employee.site, Employee.affaire,
            Attendance.present, Attendance.absent, Attendance.cong,
            Attendance.tour_rep, Attendance.repos_med, Attendance.sans_ph
        ).order_by(Attendance.date.asc(), Employee.matricule.asc()).all()

        detail_data = []
        for r in rows:
            detail_data.append({
                'Date': r[0].strftime('%d/%m/%Y'),
                'Matricule': r[1],
                'Nom': r[2] or '',
                'Prénom': r[3] or '',
                'Poste': r[4] or '',
                'Site': r[5] or '',
                'Affaire': r[6] or '',
                'Présent': int(r[7] or 0),
                'Absent': int(r[8] or 0),
                'CONG': int(r[9] or 0),
                'Tour_rep': int(r[10] or 0),
                'Repos_med': int(r[11] or 0),
                'Sans_ph': int(r[12] or 0)
            })

        detail_df = pd.DataFrame(detail_data)
        out_path = os.path.join(app.config['EXPORT_FOLDER'], f'recap_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx')

        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            # --- FEUILLE TABLEAU CROISÉ ---
            if not detail_df.empty:
                statuses = ['Présent', 'Absent', 'CONG', 'Tour_rep', 'Repos_med', 'Sans_ph']
                long_rows = []
                for _, r in detail_df.iterrows():
                    for s in statuses:
                        val = 'X' if int(r[s]) != 0 else ''
                        long_rows.append({
                            'Date': r['Date'],
                            'Matricule': r['Matricule'],
                            'Nom': r['Nom'],
                            'Prénom': r['Prénom'],
                            'Poste': r['Poste'],
                            'Site': r['Site'],
                            'Affaire': r['Affaire'],
                            'Status': s,
                            'Value': val
                        })
                long_df = pd.DataFrame(long_rows)

                pivot = long_df.pivot_table(
                    index=['Matricule', 'Nom', 'Prénom', 'Poste', 'Site', 'Affaire'],
                    columns=['Date', 'Status'],
                    values='Value',
                    aggfunc=lambda x: next((v for v in x if v and str(v).strip() != ''), ''),
                    fill_value=''
                )

                try:
                    cols_sorted = sorted(pivot.columns, key=lambda t: datetime.strptime(t[0], '%d/%m/%Y'))
                    pivot = pivot[cols_sorted]
                except Exception:
                    pass

                pivot.to_excel(writer, sheet_name='Tableau croisé')
            else:
                pd.DataFrame(columns=['Matricule', 'Nom', 'Prénom', 'Poste', 'Site', 'Affaire']).to_excel(writer, sheet_name='Tableau croisé')

            # --- FEUILLE RÉCAP ---
            numeric_cols = ['Présent', 'Absent', 'CONG', 'Tour_rep', 'Repos_med', 'Sans_ph']
            for c in numeric_cols:
                if c in recap_df.columns:
                    recap_df[c] = recap_df[c].astype(int)
            totals = {col: recap_df[col].sum() if col in recap_df.columns else 0 for col in numeric_cols}
            totals_row = {'Matricule': 'TOTAL', 'Nom': '', 'Prénom': '', 'Poste': '', 'Site': '', 'Affaire': ''}
            totals_row.update(totals)
            recap_df_with_total = pd.concat([recap_df, pd.DataFrame([totals_row])], ignore_index=True)
            recap_df_with_total.to_excel(writer, sheet_name='Récap', index=False)

        # --- STYLING ET COULEURS ---
        wb = openpyxl.load_workbook(out_path)
        ws = wb['Tableau croisé']

        header_fill = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')
        font_bold = Font(bold=True)
        center = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )

        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin_border

        for row in ws.iter_rows(min_row=1, max_row=2):
            for cell in row:
                cell.fill = header_fill
                cell.font = font_bold
                cell.alignment = center

        prev_val = None
        start_col = None
        color_index = 0
        colors = ["FFF2CC", "D9EAD3", "CFE2F3", "F4CCCC", "EAD1DC", "C9DAF8"]

        fixed_cols = 6  # Matricule, Nom, Prénom, Poste, Site, Affaire
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=1, column=col_idx).value
            if col_idx <= fixed_cols:
                ws.merge_cells(start_row=1, start_column=col_idx, end_row=2, end_column=col_idx)
                ws.cell(row=1, column=col_idx).alignment = center
                continue
            if val != prev_val:
                color_index = (color_index + 1) % len(colors)
                if prev_val is not None and start_col is not None and (col_idx - start_col) > 1:
                    ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=col_idx - 1)
                start_col = col_idx
                prev_val = val
            col_fill = PatternFill(start_color=colors[color_index], end_color=colors[color_index], fill_type='solid')
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                ws.cell(row=row[0].row, column=col_idx).fill = col_fill

        if prev_val and start_col and (ws.max_column - start_col + 1) > 1:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=ws.max_column)

        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 14

        # FEUILLE RECAP STYLÉE
        if 'Récap' in wb.sheetnames:
            ws2 = wb['Récap']
            fill_header = PatternFill(start_color='9BC2E6', end_color='9BC2E6', fill_type='solid')
            for cell in ws2[1]:
                cell.fill = fill_header
                cell.font = font_bold
                cell.alignment = center
                cell.border = thin_border
            total_fill = PatternFill(start_color='FFE699', end_color='FFE699', fill_type='solid')
            for cell in ws2[ws2.max_row]:
                cell.fill = total_fill
                cell.font = font_bold
                cell.alignment = center
                cell.border = thin_border
            for row in ws2.iter_rows():
                for cell in row:
                    cell.border = thin_border
            for col in range(1, ws2.max_column + 1):
                ws2.column_dimensions[get_column_letter(col)].width = 14

        wb.save(out_path)
        return send_file(out_path, as_attachment=True, download_name=os.path.basename(out_path))

    return app


# --- FONCTIONS UTILITAIRES ---

def normalize_col(name):
    if not isinstance(name, str):
        return ''
    return re.sub(r'[^A-Za-z0-9]', '', name).strip().lower()


def possible_status_cols(cols):
    mapping = {}
    for c in cols:
        n = normalize_col(str(c))
        if n in ['present', 'présent', 'prsent']:
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
    m = re.search(r'(20\d{2}[01]\d[0-3]\d)', filename)
    if m:
        try:
            return datetime.strptime(m.group(1), '%Y%m%d').date()
        except Exception:
            pass
    m = re.search(r'(\b\d{2}[\-_/]?\d{2}[\-_/]?\d{4}\b)', filename)
    if m:
        s = re.sub(r'[^0-9]', '', m.group(1))
        try:
            return datetime.strptime(s, '%d%m%Y').date()
        except Exception:
            pass
    return None


def process_upload(path, filename):
    try:
        df = pd.read_excel(path, engine='openpyxl')
    except Exception:
        try:
            df = pd.read_excel(path)
        except Exception:
            df = pd.read_csv(path)

    df.columns = [str(c).strip() for c in df.columns]
    colmap = {normalize_col(c): c for c in df.columns}

    if not any(k in colmap for k in ['matricule', 'id']):
        raise ValueError('Fichier sans colonne Matricule / id')

    date_col = None
    for k in colmap:
        if k == 'date':
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
                for key in ['matricule', 'id']:
                    if key in colmap:
                        matricule = str(row[colmap[key]]).strip()
                        break
                if not matricule or matricule.lower().startswith('nan'):
                    continue

                emp = Employee.query.filter_by(matricule=matricule).first()
                if not emp:
                    emp = Employee(matricule=matricule)
                    db.session.add(emp)
                    db.session.flush()

                for fld in [('nom', 'nom'), ('prenom', 'prenom'), ('poste', 'poste'), ('site', 'site'),
                            ('affaire', 'affaire')]:
                    nkey = fld[0]
                    if nkey in colmap:
                        val = row[colmap[nkey]]
                        if pd.notna(val):
                            setattr(emp, fld[1], str(val).strip())

                row_date = None
                if date_col:
                    raw = row[date_col]
                    if pd.isna(raw):
                        row_date = file_date or date.today()
                    elif isinstance(raw, (datetime, pd.Timestamp)):
                        row_date = raw.date()
                    else:
                        try:
                            row_date = pd.to_datetime(raw).date()
                        except Exception:
                            row_date = file_date or date.today()
                else:
                    row_date = file_date or date.today()

                status_map = possible_status_cols(df.columns)
                vals = dict(present=0, absent=0, cong=0, tour_rep=0, repos_med=0, sans_ph=0)
                for orig_col, canon in status_map.items():
                    try:
                        v = row[orig_col]
                        if pd.isna(v):
                            continue
                        if isinstance(v, str):
                            if v.strip().lower() in ['x', '1', 'y', 'yes', 'présent', 'present', 'p']:
                                vals[canon] = 1
                        else:
                            if int(v) != 0:
                                vals[canon] = 1
                    except Exception:
                        continue

                if sum(vals.values()) == 0:
                    continue

                att = Attendance.query.filter_by(employee_id=emp.id, date=row_date).first()
                if not att:
                    att = Attendance(employee_id=emp.id, date=row_date, **vals)
                    db.session.add(att)
                else:
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
    return pd.DataFrame(data)


if __name__ == '__main__':
    app = create_app()
    app.run(debug=True)
