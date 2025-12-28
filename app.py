from flask import Flask, render_template, send_file, request, redirect, url_for
from sqlalchemy import create_engine, text
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io
import urllib.parse
import os

app = Flask(__name__)

# --- DATABASE CONFIGURATION ---
# NOTE: You must change this to a Cloud Database connection string for Render!
# For now, it reads from Environment Variables or defaults to your local (which won't work on cloud)
server = os.environ.get('DB_SERVER', 'ADHAM')
database = os.environ.get('DB_NAME', 'InstallmentDB')
username = os.environ.get('DB_USER', 'sa')
password = os.environ.get('DB_PASS', 'your_password')

# Connection String for Linux (Render)
params = urllib.parse.quote_plus(
    f'DRIVER={{ODBC Driver 17 for SQL Server}};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password};'
)

app.config['SQLALCHEMY_DATABASE_URI'] = f"mssql+pyodbc:///?odbc_connect={params}"
engine = create_engine(app.config['SQLALCHEMY_DATABASE_URI'])

@app.route('/')
def index():
    try:
        with engine.connect() as conn:
            query = text("SELECT * FROM v_InstallmentStatus ORDER BY DueDate ASC")
            installments = conn.execute(query).fetchall()

            summary_query = text("""
                SELECT 
                    ISNULL(SUM(PaidAmount), 0) as Collected,
                    ISNULL(SUM(Amount - PaidAmount), 0) as Pending,
                    ISNULL(SUM(CASE WHEN DueDate < GETDATE() AND IsPaid = 0 THEN Amount - PaidAmount ELSE 0 END), 0) as Late
                FROM Installments
            """)
            summary = conn.execute(summary_query).fetchone()
        return render_template('index.html', installments=installments, summary=summary)
    except Exception as e:
        return f"<h3>Database Connection Error</h3><p>{e}</p><p>Ensure your database is hosted on the cloud (Azure/AWS) and credentials are correct.</p>"

@app.route('/export')
def export_excel():
    query = "SELECT * FROM v_InstallmentStatus"
    with engine.connect() as conn:
        df = pd.read_sql(query, conn)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Installments')
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Installments_Report.xlsx'
    )

@app.route('/add', methods=['GET', 'POST'])
def add_new():
    if request.method == 'POST':
        name = request.form['fullname']
        phone = request.form['phone']
        item = request.form['item']
        total_amount = float(request.form['total_amount'])
        months = int(request.form['months'])
        start_date_str = request.form['start_date']
        
        monthly_amount = total_amount / months
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')

        with engine.connect() as conn:
            trans = conn.begin()
            try:
                res_cust = conn.execute(
                    text("INSERT INTO Customers (FullName, PhoneNumber) OUTPUT INSERTED.CustomerID VALUES (:name, :phone)"),
                    {"name": name, "phone": phone}
                )
                customer_id = res_cust.fetchone()[0]

                res_cont = conn.execute(
                    text("INSERT INTO Contracts (CustomerID, ItemDescription, TotalAmount, StartDate) OUTPUT INSERTED.ContractID VALUES (:cid, :item, :total, :date)"),
                    {"cid": customer_id, "item": item, "total": total_amount, "date": start_date}
                )
                contract_id = res_cont.fetchone()[0]

                for i in range(months):
                    due_date = start_date + relativedelta(months=i+1)
                    conn.execute(
                        text("INSERT INTO Installments (ContractID, DueDate, Amount) VALUES (:cont_id, :due, :amt)"),
                        {"cont_id": contract_id, "due": due_date, "amt": monthly_amount}
                    )
                
                trans.commit()
                return redirect(url_for('index'))
            except Exception as e:
                trans.rollback()
                return f"Error: {e}"

    return render_template('add_new.html')

@app.route('/pay/<int:id>', methods=['POST'])
def pay_installment(id):
    update_query = text("""
        UPDATE Installments 
        SET IsPaid = 1, 
            PaidAmount = Amount, 
            PaymentDate = GETDATE() 
        WHERE InstallmentID = :id
    """)
    try:
        with engine.connect() as conn:
            conn.execute(update_query, {'id': id})
            conn.commit()
    except Exception as e:
        print(f"Error: {e}")
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)