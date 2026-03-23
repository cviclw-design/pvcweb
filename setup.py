from app import app, db
from models import User
import pandas as pd
import os

def setup_database():
    with app.app_context():
        db.create_all()
        
        # Create admin user
        if not User.query.filter_by(username='admin').first():
            admin = User(username='admin', password_hash='pbkdf2:sha256:...')
            db.session.add(admin)
            db.session.commit()
        
        # Import IEEMA data (modify paths)
        import_ieema('IEEMA.xlsx', 'ieema_transformer')

def import_ieema(excel_file, table_name):
    df = pd.read_excel(excel_file)
    # Convert to table_name and insert
    pass

if __name__ == '__main__':
    setup_database()
    print("✅ Database setup complete! Run: python app.py")