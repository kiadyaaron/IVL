import os
import re
from datetime import datetime, date
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from flask_migrate import Migrate
from werkzeug.utils import secure_filename
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from models import db, Employee, Attendance
import config

# Extensions de fichiers acceptées pour le téléchargement
UPLOAD_EXTENSIONS = ['.xlsx', '.xls', '.csv']


def create_app():
    app = Flask(__name__)
    app.config.from_object('config')
    db.init_app(app)
    Migrate(app, db)

    # Créer les dossiers d'export et d'upload s'ils n'existent pas
    os.makedirs(app.config.get('EXPORT_FOLDER', os.path.join(os.path.dirname(__file__), 'exports')), exist_ok=True)
    os.makedirs(app.config.get('UPLOAD_FOLDER', os.path.join(os.path.dirname(__file__), 'uploads')), exist_ok=True)

    with app.app_context():
        # S'assurer que les tables de la base de données existent
        db.create_all()

    @app.route('/', methods=['GET', 'POST'])
    def index():
        if request.method == 'POST':
            f = request.files.get('file')
            if not f:
                flash('Aucun fichier téléchargé', 'warning')
                return redirect(url_for('index'))

            filename = secure_filename(f.filename)
            if not filename:
                flash('Nom de fichier invalide', 'warning')
                return redirect(url_for('index'))

            ext = os.path.splitext(filename)[1].lower()
            if ext not in UPLOAD_EXTENSIONS:
                flash('Type de fichier non supporté', 'danger')
                return redirect(url_for('index'))

            save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            f.save(save_path)

            try:
                rows_added = process_upload(save_path)
                flash(f'Fichier importé avec succès. Lignes ajoutées/mises à jour : {rows_added}', 'success')
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
        table_html = recap_df.to_html(classes='table table-striped table-sm', index=False) if not recap_df.empty else '<p>Aucune donnée</p>'
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

        # Export simple (Détails + Récap)
        q = Attendance.query.join(Employee)
        if start_date:
            q = q.filter(Attendance.date >= start_date)
        if end_date:
            q = q.filter(Attendance.date <= end_date)

        rows = q.with_entities(
            Attendance.date,
            Employee.matricule, Employee.nom, Employee.prenom, Employee.poste,
            Employee.site, Employee.affaire,
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
            if not detail_df.empty:
                detail_df.to_excel(writer, sheet_name='Détails', index=False)
            else:
                pd.DataFrame(columns=['Date', 'Matricule']).to_excel(writer, sheet_name='Détails', index=False)

            recap_df.to_excel(writer, sheet_name='Récap', index=False)

        return send_file(out_path, as_attachment=True, download_name=os.path.basename(out_path))

    return app


# -------------------------
# Fonctions utilitaires
# -------------------------

def _normalize(s: str) -> str:
    """Normalise une chaîne de caractères pour la comparaison des en-têtes de colonnes."""
    if s is None:
        return ''
    # Supprime les caractères non alphanumériques (y compris les espaces)
    return re.sub(r'[^A-Za-z0-9éàèêôùïçÉÀÈ]', '', str(s)).strip().lower()


def _is_date_string(s: str) -> bool:
    """Vérifie si la chaîne correspond au format de date jj/mm/aaaa."""
    if not s:
        return False
    s = str(s).strip()
    # Accepte jj/mm/aaaa ou j/m/aaaa
    return bool(re.match(r'^\d{1,2}\/\d{1,2}\/\d{4}$', s))


def process_upload(path: str) -> int:
    """
    Lit un fichier Excel à double en-tête (header=[0,1]) où les dates sont en première ligne
    et les sous-colonnes (Présent, Absent, ...) en seconde ligne.
    Insère / met à jour les attendances.
    Retourne le nombre d'enregistrements d'attendance traités.
    """
    # Lecture en multi-index header (lignes 0 et 1)
    df = pd.read_excel(path, header=[0, 1], engine='openpyxl')
    # S'assurer que les noms de colonnes sont strings pour la cohérence
    df.columns = [(str(a).strip(), str(b).strip()) for a, b in df.columns]

    # Définition des variantes canoniques pour les colonnes fixes
    canonical_fixed = {
        'matricule': ['matricule', 'id'],
        'nom': ['nom', 'name'],
        'prenom': ['prenom', 'prénom', 'prenom'],
        'poste': ['poste', 'position'],
        'site': ['site'],
        'affaire': ['affaire'],
        'taux_logement': ['tauxlogement', 'taux_logement', 'taux_lgt', 'taux logement', 'tauxlgt'],
        'taux_repas': ['tauxrepas', 'taux_repas', 'taux repas']
    }

    # Identification des colonnes fixes dans le DataFrame
    fixed_map = {}
    for col in df.columns:
        a, b = col
        a_n = _normalize(a)
        b_n = _normalize(b)
        for canon, variants in canonical_fixed.items():
            # Une correspondance est trouvée si le nom normalisé de l'un des deux niveaux
            # correspond à une variante canonique
            if a_n in variants or b_n in variants:
                fixed_map[canon] = col

    # Identification des colonnes de date / statut
    # La structure de l'image (Date en niveau 0, Statut en niveau 1) est gérée ici.
    date_columns = {}
    for col in df.columns:
        a, b = col # a = niveau 0 (Date), b = niveau 1 (Statut)
        if _is_date_string(a):
            date_str = a.strip()
            status = b.strip() or ''
            date_columns.setdefault(date_str, {})[status] = col

    # Si aucune date n'est trouvée au niveau 0, on essaie au niveau 1 (cas moins fréquent)
    if not date_columns:
        for col in df.columns:
            a, b = col
            if _is_date_string(b):
                date_str = b.strip()
                status = a.strip() or ''
                date_columns.setdefault(date_str, {})[status] = col

    rows_processed = 0
    with db.session.begin():
        for idx, row in df.iterrows():
            # 1. Identification de l'employé par Matricule
            matricule = None
            if 'matricule' in fixed_map:
                matricule = row[fixed_map['matricule']]
            else:
                # Heuristique de secours: prendre la première colonne
                matricule = row[df.columns[0]]

            if pd.isna(matricule):
                continue
            matricule = str(matricule).strip()
            if not matricule:
                continue

            emp = Employee.query.filter_by(matricule=matricule).first()
            if not emp:
                emp = Employee(matricule=matricule)
                db.session.add(emp)
                db.session.flush()

            # 2. Mise à jour des informations fixes de l'employé
            for key in ('nom', 'prenom', 'poste', 'site', 'affaire'):
                if key in fixed_map:
                    val = row[fixed_map[key]]
                    if pd.notna(val):
                        try:
                            setattr(emp, key, str(val).strip())
                        except Exception:
                            pass

            # Mise à jour des taux (logement / repas)
            if 'taux_logement' in fixed_map:
                try:
                    val = row[fixed_map['taux_logement']]
                    if pd.notna(val):
                        v = float(val)
                        # Le code tente de mettre à jour taux_lgt ou taux_logement
                        if hasattr(emp, 'taux_lgt'):
                            setattr(emp, 'taux_lgt', v)
                        elif hasattr(emp, 'taux_logement'):
                            setattr(emp, 'taux_logement', v)
                except Exception:
                    pass

            if 'taux_repas' in fixed_map:
                try:
                    val = row[fixed_map['taux_repas']]
                    if pd.notna(val):
                        v = float(val)
                        if hasattr(emp, 'taux_repas'):
                            setattr(emp, 'taux_repas', v)
                except Exception:
                    pass

            # 3. Parcours des dates et statuts (le cœur de l'import)
            for date_str, status_map in date_columns.items():
                try:
                    date_obj = datetime.strptime(date_str, '%d/%m/%Y').date()
                except Exception:
                    continue

                present_flag = 0
                absent_flag = 0
                cong_flag = 0
                tourrep_flag = 0
                repos_flag = 0
                sansph_flag = 0

                # HELPER CORRIGÉ: Teste si une cellule indique une présence/un statut
                def cell_true(val):
                    """Vérifie si la valeur de la cellule indique un statut positif (Présent, Absent, X, 1, 1.0, etc.)."""
                    if pd.isna(val):
                        return False
                    
                    # CORRECTION: Si c'est un nombre (int ou float, comme 1.0), vérifiez s'il est > 0
                    if isinstance(val, (int, float)):
                        return val > 0
                    
                    # Sinon, traitez-le comme une chaîne
                    s = str(val).strip().lower()
                    if not s:
                        return False
                    
                    # Correspondance avec les indicateurs courants ('x', '1', 'yes', 'p', etc.)
                    return s in ('x', '1', 'yes', 'y', 'présent', 'present', 'p') or re.match(r'^\d+$', s) and int(s) > 0


                # Vérification des statuts en utilisant la clé du niveau 1 de l'en-tête (ex: 'Présent', 'Absent', 'Repos médica')
                for status_label, colkey in status_map.items():
                    sval = row[colkey]
                    
                    if pd.notna(sval) and cell_true(sval):
                        sn = status_label.strip().lower()
                        
                        # Correspondance souple avec les labels
                        if 'prés' in sn or 'present' in sn:
                            present_flag = 1
                        elif 'abs' in sn:
                            absent_flag = 1
                        elif 'cong' in sn:
                            cong_flag = 1
                        elif 'tour' in sn and 'rep' in sn:
                            tourrep_flag = 1
                        elif 'repos' in sn or 'méd' in sn or 'med' in sn:
                            repos_flag = 1
                        elif 'sans' in sn and 'ph' in sn:
                            sansph_flag = 1

                # Si aucun statut n'est marqué pour cette date, on passe
                if (present_flag + absent_flag + cong_flag + tourrep_flag + repos_flag + sansph_flag) == 0:
                    continue

                # 4. Insertion / Mise à jour de l'attendance (upsert)
                att = Attendance.query.filter_by(employee_id=emp.id, date=date_obj).first()
                if not att:
                    att = Attendance(employee_id=emp.id, date=date_obj,
                                     present=present_flag, absent=absent_flag, cong=cong_flag,
                                     tour_rep=tourrep_flag, repos_med=repos_flag, sans_ph=sansph_flag)
                    db.session.add(att)
                else:
                    # OR-combine les marques existantes et nouvelles (priorité à 1 si déjà 1)
                    att.present = max(att.present, present_flag)
                    att.absent = max(att.absent, absent_flag)
                    att.cong = max(att.cong, cong_flag)
                    att.tour_rep = max(att.tour_rep, tourrep_flag)
                    att.repos_med = max(att.repos_med, repos_flag)
                    att.sans_ph = max(att.sans_ph, sansph_flag)

                rows_processed += 1

    return rows_processed


def build_recap(start_date=None, end_date=None):
    """Construit le DataFrame de récapitulatif des présences par employé."""
    q = Attendance.query.join(Employee).with_entities(
        Employee.matricule, Employee.nom, Employee.prenom, Employee.poste,
        Employee.site, Employee.affaire,
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