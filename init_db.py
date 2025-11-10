import pymysql
pymysql.install_as_MySQLdb()

from app import create_app
from models import db

app = create_app()

with app.app_context():
    db.create_all()
    print("Database initialized at", app.config["SQLALCHEMY_DATABASE_URI"])
