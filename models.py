from flask_sqlalchemy import SQLAlchemy
from datetime import date

db = SQLAlchemy()

class Employee(db.Model):
    __tablename__ = 'employees'
    id = db.Column(db.Integer, primary_key=True)
    matricule = db.Column(db.String(64), unique=True, nullable=False, index=True)
    nom = db.Column(db.String(128))
    prenom = db.Column(db.String(128))
    poste = db.Column(db.String(128))
    site = db.Column(db.String(128))
    affaire = db.Column(db.String(128))

    # nouvelles colonnes demandées
    classe = db.Column(db.String(50))
    affectation = db.Column(db.String(100))
    ville = db.Column(db.String(100))

    # taux statiques (float)
    taux_lgt = db.Column(db.Float, default=0.0)   # "taux logement"
    taux_repas = db.Column(db.Float, default=0.0) # "taux repas"

    attendances = db.relationship('Attendance', back_populates='employee', cascade='all, delete-orphan')

    def __repr__(self):
        return f"<Employee {self.matricule} {self.nom}>"

class Attendance(db.Model):
    __tablename__ = 'attendances'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    date = db.Column(db.Date, nullable=False, index=True)

    present = db.Column(db.Integer, default=0)
    absent = db.Column(db.Integer, default=0)
    cong = db.Column(db.Integer, default=0)
    tour_rep = db.Column(db.Integer, default=0)
    repos_med = db.Column(db.Integer, default=0)
    sans_ph = db.Column(db.Integer, default=0)

    employee = db.relationship('Employee', back_populates='attendances')

    __table_args__ = (db.UniqueConstraint('employee_id', 'date', name='_emp_date_uc'),)

    def __repr__(self):
        return f"<Attendance emp_id={self.employee_id} date={self.date}>"
