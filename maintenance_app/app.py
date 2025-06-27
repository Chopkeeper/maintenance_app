from flask import Flask, render_template, request, redirect, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import pandas as pd

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
db = SQLAlchemy(app)

class Maintenance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    computer_name = db.Column(db.String(100))
    problem = db.Column(db.String(200))
    repair_date = db.Column(db.DateTime, default=datetime.utcnow)
    details = db.Column(db.Text)
    cleaned = db.Column(db.Boolean)

@app.route('/')
def index():
    records = Maintenance.query.order_by(Maintenance.repair_date.desc()).all()
    return render_template('index.html', records=records)

@app.route('/add', methods=['POST'])
def add():
    record = Maintenance(
        computer_name=request.form['computer_name'],
        problem=request.form['problem'],
        details=request.form['details'],
        cleaned='cleaned' in request.form
    )
    db.session.add(record)
    db.session.commit()
    return redirect('/')

@app.route('/export')
def export():
    records = Maintenance.query.all()
    df = pd.DataFrame([{
        'ชื่อเครื่อง': r.computer_name,
        'ปัญหา': r.problem,
        'วันที่ซ่อม': r.repair_date.strftime('%Y-%m-%d'),
        'รายละเอียด': r.details,
        'ทำความสะอาด': '✓' if r.cleaned else ''
    } for r in records])
    filepath = 'static/export.xlsx'
    df.to_excel(filepath, index=False)
    return send_file(filepath, as_attachment=True)

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
