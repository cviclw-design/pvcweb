from pvc1 import get_ieema_df, calculate_single_record_from_dict
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import os

# -------------------------
# 1. CREATE APP
# -------------------------
app = Flask(__name__)
app.config['SECRET_KEY'] = 'pvc-webapp-2026'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///pvc.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

ieema_df = None

with app.app_context():
    ieema_df = get_ieema_df()

@app.route('/calculate', methods=['POST'])
@login_required
def calculate():

    # 1. GET VALUES FROM FORM
    origdp_str     = request.form.get('origdp') or ''
    refixeddp_str  = request.form.get('refixeddp') or ''
    extendeddp_str = request.form.get('extendeddp') or ''
    supply_str     = request.form.get('supdate') or ''
    rateapplied = request.form.get('rateapplied') or ''

    # --------- IF VALID, CONTINUE AS BEFORE ---------
    data = {
        'user_id': current_user.id,
        'username': current_user.username,
        'item': request.form['item'],
        'basicrate': float(request.form.get('basicrate', 0) or 0),
        'quantity': float(request.form.get('quantity', 0) or 0),
        'freight': float(request.form.get('freight', 0) or 0),

        'pvcbasedate': request.form.get('pvcbasedate') or '',
        'origdp': origdp_str,
        'refixeddp': refixeddp_str,
        'extendeddp': extendeddp_str,
        'caldate': request.form.get('caldate') or '',
        'supdate': supply_str,
        'rateapplied': rateapplied,

        'lowerrate': float(request.form.get('lowerrate', 0) or 0),
        'lowerfreight': float(request.form.get('lowerfreight', 0) or 0),
        'lowerbasicdate': request.form.get('lowerbasicdate') or '',
    }

    
    # Map to pvc1.py expected keys
    one = {
        "acc_qty": data['quantity'],
        "basic_rate": data['basicrate'],
        "freight_rate_per_unit": data['freight'],

        "pvc_base_date": data['pvcbasedate'],
        "call_date": data['caldate'],

        # DP dates used inside pvc1.py to derive:
        # - scheduled_date for PVC contractual
        # - LD base date for delay
        "orig_dp": data['origdp'],
        "refixeddp": data['refixeddp'],
        "extendeddp": data['extendeddp'],

        "sup_date": data['supdate'],

        "lower_rate": data['lowerrate'],
        "lower_freight": data['lowerfreight'],
        "lower_basic_date": data['lowerbasicdate'],
        "rateapplied": data['rateapplied'], 
    }

    global ieema_df
    result_row = calculate_single_record_from_dict(one, ieema_df)

    # Scenario amounts (A2/B2/C1/D1) after LD where applicable
    scenario_amounts = {
        "A2": result_row.get("pvc_actual_less_ld_new"),
        "B2": result_row.get("pvc_contractual_less_ld_new"),
        "C1": result_row.get("lower_actual"),
        "D1": result_row.get("lower_contractual"),
    }
    selected = result_row["selected_scenario_new"]

    # per-set PVC based on selected scenario
    pvc_per_set = None
    if selected == "A2":
        pvc_per_set = result_row.get("pvc_per_set_a2")
    elif selected == "B2":
        pvc_per_set = result_row.get("pvc_per_set_b2")
    elif selected == "C1":
        pvc_per_set = result_row.get("pvc_per_set_c1")
    elif selected == "D1":
        pvc_per_set = result_row.get("pvc_per_set_d1")

    result = {
        "data": {
            "pvcactual": result_row["pvc_actual"],
            "pvccontractual": result_row["pvc_contractual"],
            "lower_actual": result_row["lower_actual"],
            "lower_contractual": result_row["lower_contractual"],
            "ldamtactual": result_row["ld_amt_actual"],
            "ld_amt_contractual": result_row["ld_amt_contractual"],

            "fairprice": result_row["fair_price_new"],
            
            "delay_days": result_row.get("delay_days"),
            "ld_weeks": result_row.get("ld_weeks_new"),
            "ld_rate_pct": result_row.get("ld_rate_pct_new"),
            "ld_applicable": result_row.get("ld_applicable", True),
            "selectedscenario": selected,
            "pvc_per_set_a2": result_row.get("pvc_per_set_a2"),
            "pvc_per_set_b2": result_row.get("pvc_per_set_b2"),
            "pvc_per_set_c1": result_row.get("pvc_per_set_c1"),
            "pvc_per_set_d1": result_row.get("pvc_per_set_d1"),
        },
        "scenario_details": result_row.get("scenario_details", []),
        "scenario_amounts": scenario_amounts,
    }

    calc = PVCResult(
        user_id=data['user_id'],
        username=data['username'],
        item=data['item'],
        basicrate=data['basicrate'],
        quantity=data['quantity'],
        freight=data['freight'],
        pvcbasedate=data['pvcbasedate'],
        origdp=data['origdp'],
        refixeddp=data['refixeddp'],
        extendeddp=data['extendeddp'],
        caldate=data['caldate'],
        supdate=data['supdate'],
        rateapplied=data['rateapplied'],
        pvcactual=result["data"]["pvcactual"],
        ldamtactual=result["data"]["ldamtactual"],
        fairprice=result["data"]["fairprice"],
        selectedscenario=result["data"]["selectedscenario"],
    )
    db.session.add(calc)
    db.session.commit()

    return render_template(
        'result.html',
        item=data['item'],
        data=data,
        result=result,
        calc_id=calc.id
    )

# -------------------------
# 5. INIT DB & MAIN
# -------------------------
def init_db():
    db.create_all()
    if not User.query.filter_by(username='admin').first():
        admin = User(
            username='admin',
            password_hash=generate_password_hash('admin123')
        )
        db.session.add(admin)
        db.session.commit()
        print("Admin user created: admin / admin123")

import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
