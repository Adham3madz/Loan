from flask import Flask, render_template, send_file, request, redirect, url_for
from sqlalchemy import create_engine, text
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta # For adding months easily
import io
from flask import request, redirect, url_for
import urllib.parse

app = Flask(__name__)

# --- CONFIGURATION FOR SERVER 'ADHAM' ---
# Using Windows Authentication (Trusted_Connection=Yes)
params = urllib.parse.quote_plus(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    'SERVER=ADHAM;'
    'DATABASE=InstallmentDB;'
    'Trusted_Connection=yes;'
)
app.config['SQLALCHEMY_DATABASE_URI'] = f"mssql+pyodbc:///?odbc_connect={params}"
engine = create_engine(app.config['SQLALCHEMY_DATABASE_URI'])

@app.route('/')
def index():
    with engine.connect() as conn:
        # 1. Get the List of Installments (Existing Code)
        query = text("SELECT * FROM v_InstallmentStatus ORDER BY DueDate ASC")
        installments = conn.execute(query).fetchall()

        # 2. Get the Dashboard Totals (NEW CODE)
        # We use ISNULL to return 0 instead of None if the table is empty
        summary_query = text("""
            SELECT 
                ISNULL(SUM(PaidAmount), 0) as Collected,
                ISNULL(SUM(Amount - PaidAmount), 0) as Pending,
                ISNULL(SUM(CASE WHEN DueDate < GETDATE() AND IsPaid = 0 THEN Amount - PaidAmount ELSE 0 END), 0) as Late
            FROM Installments
        """)
        summary = conn.execute(summary_query).fetchone()

    # Pass 'summary' to the HTML
    return render_template('index.html', installments=installments, summary=summary)

@app.route('/export')
def export_excel():
    # Export data using Pandas
    query = "SELECT * FROM v_InstallmentStatus"
    
    # Read into DataFrame
    with engine.connect() as conn:
        df = pd.read_sql(query, conn)
    
    # Create Excel file in memory
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
        # 1. Get Form Data
        name = request.form['fullname']
        phone = request.form['phone']
        item = request.form['item']
        total_amount = float(request.form['total_amount'])
        months = int(request.form['months'])
        start_date_str = request.form['start_date']
        
        # Calculate monthly installment
        monthly_amount = total_amount / months
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')

        with engine.connect() as conn:
            trans = conn.begin() # Start Transaction (Safety)
            try:
                # A. Create Customer
                # Note: In a real app, you'd check if customer exists first. 
                # Here we simplify: always create new for now.
                res_cust = conn.execute(
                    text("INSERT INTO Customers (FullName, PhoneNumber) OUTPUT INSERTED.CustomerID VALUES (:name, :phone)"),
                    {"name": name, "phone": phone}
                )
                customer_id = res_cust.fetchone()[0]

                # B. Create Contract
                res_cont = conn.execute(
                    text("INSERT INTO Contracts (CustomerID, ItemDescription, TotalAmount, StartDate) OUTPUT INSERTED.ContractID VALUES (:cid, :item, :total, :date)"),
                    {"cid": customer_id, "item": item, "total": total_amount, "date": start_date}
                )
                contract_id = res_cont.fetchone()[0]

                # C. Generate Installments (The Loop)
                for i in range(months):
                    # Add 'i' months to the start date
                    due_date = start_date + relativedelta(months=i+1)
                    
                    conn.execute(
                        text("INSERT INTO Installments (ContractID, DueDate, Amount) VALUES (:cont_id, :due, :amt)"),
                        {"cont_id": contract_id, "due": due_date, "amt": monthly_amount}
                    )
                
                trans.commit() # Save everything
                return redirect(url_for('index'))
                
            except Exception as e:
                trans.rollback() # Undo if error
                return f"Error occurred: {e}"

    return render_template('add_new.html')

# --- Add this new route to app.py ---

@app.route('/pay/<int:id>', methods=['POST'])
def pay_installment(id):
    # This SQL query marks the installment as fully paid
    update_query = text("""
        UPDATE Installments 
        SET IsPaid = 1, 
            PaidAmount = Amount, 
            PaymentDate = GETDATE() 
        WHERE InstallmentID = :id
    """)
    
    try:
        with engine.connect() as conn:
            # We need to commit the transaction explicitly
            conn.execute(update_query, {'id': id})
            conn.commit() 
    except Exception as e:
        print(f"Error: {e}")
        
    return redirect(url_for('index'))

if __name__ == '__main__':
    # Make accessible on local network (for mobile access)
    app.run(host='0.0.0.0', port=5000, debug=True)