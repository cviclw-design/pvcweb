from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from sqlalchemy import text
from datetime import datetime
import math

db = SQLAlchemy()

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(120), nullable=False)

class PVCResult(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    item = db.Column(db.String(50), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    # Add all your PVC fields here
    basicrate = db.Column(db.Float)
    quantity = db.Column(db.Float)
    pvcactual = db.Column(db.Float)
    pvccontractual = db.Column(db.Float)
    ldamtactual = db.Column(db.Float)
    fairprice = db.Column(db.Float)
    selectedscenario = db.Column(db.String(10))
    result_json = db.Column(db.Text)  # Full detailed result

def get_ieema_data(table_name, base_date, current_date):
    """Replace your Excel loading with DB query"""
    base_query = text(f"""
        SELECT * FROM {table_name} 
        WHERE date = :base_date
    """)
    curr_query = text(f"""
        SELECT * FROM {table_name} 
        WHERE date = :current_date
    """)
    
    base_row = db.engine.execute(base_query, {'base_date': base_date}).fetchone()
    curr_row = db.engine.execute(curr_query, {'current_date': current_date}).fetchone()
    
    return dict(base_row), dict(curr_row) if curr_row else None

def calculate_pvc(data, ieema_table, weights):
    """Your existing pvc1.py logic adapted for DB"""
    base_date = data['pvcbasedate']
    current_date = data['calldate']
    
    base_row, curr_row = get_ieema_data(ieema_table, base_date, current_date)
    
    # Your WPICOEFF calculation logic here (unchanged)
    pvc_percent = 0
    for index, weight in weights.items():
        base_val = base_row.get(index, 0) if base_row else 0
        curr_val = curr_row.get(index, 0) if curr_row else 0
        pvc_percent += weight * (curr_val - base_val) / base_val / 100
    
    # LD calculation (your existing logic)
    delay_days = max((datetime.strptime(data['supdate'], '%Y-%m-%d') - 
                     datetime.strptime(data['scheduleddate'], '%Y-%m-%d')).days, 0)
    ld_weeks = math.ceil(delay_days / 7)
    ld_rate = min(ld_weeks * 0.5, 10)
    
    base_amt = data['basicrate'] * data['quantity']
    pvc_actual = base_amt * (1 + pvc_percent/100)
    
    return {
        'data': {
            'user_id': data['user_id'],
            'item': data['item'],
            'basicrate': data['basicrate'],
            'quantity': data['quantity'],
            'pvcactual': pvc_actual,
            'pvc_percent': pvc_percent,
            'ldamtactual': pvc_actual * ld_rate / 100,
            'fairprice': pvc_actual * (1 - ld_rate/100),
            'selectedscenario': 'A2'
        },
        'details': {'weights': weights, 'indices': {'base': base_row, 'current': curr_row}}
    }