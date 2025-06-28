from flask import Flask, request, render_template, send_file, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename
from loan_sheet_calculator import get_loan_sheet
from default_calc_v2 import export_to_excel, manual
import tempfile
import os
from flask import Flask, redirect, url_for, session, request, render_template, send_file
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from authlib.integrations.flask_client import OAuth
import os

app = Flask(__name__)





#####################################

@app.route('/upload', methods=['GET','POST'])
def upload():
    if request.method=='POST':
        file = request.files.get('file')
        months_raw = request.form.get('months')

        if not file or file.filename.split('.')[-1].lower() not in ['csv','xls','xlsx']:
            return "Invalid file", 400
        
        months=[int(x.strip()) for x in months_raw.split(',') if x.strip().isdigit()]

        filename = secure_filename(file.filename)
        base_name=filename.split('_')[0]
        input_path = os.path.join('input',f'{base_name}_loandata.xlsx')

        os.makedirs('input', exist_ok=True)
        os.makedirs('output', exist_ok=True)
        file.save(input_path)

        # Generate zip
        zip_filename = get_loan_sheet(input_path, months)
        return redirect(url_for('download_file', filename=zip_filename))  # Auto-download

    return render_template('upload.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory('output', filename, as_attachment=True)





@app.route('/upload_new',methods=['POST'])
def upload_new():
    use_manual=request.form.get('use_manual')=='true'
    transactions_file=request.files['transactions']
    days_before=int(request.form['days_before'])
    days_after=int(request.form['days_after'])


    os.makedirs("uploads",exist_ok=True)
    txn_filename=secure_filename(transactions_file.filename)
    base_name=txn_filename.split('_')[0]
    txn_path = os.path.join("uploads", txn_filename)
    transactions_file.save(txn_path)

    os.makedirs("output", exist_ok=True)

    if use_manual:
        start = request.form['start']
        tenor = int(request.form['tenor'])
        repayment = float(request.form['repayment'])
        freq = request.form['frequency']        

        output_filename = f"{base_name}_repayment_data.xlsx"
        output_path = os.path.join("output", output_filename)

        output_path, schedule = manual(start, tenor, repayment, freq, txn_path, days_before, days_after, output_path)

        # Package all needed data for preview
        loan_info = {
            "start": start,
            "tenor": tenor,
            "repayment": repayment,
            "freq": freq,
        }
        filename = os.path.basename(output_path)

        return render_template('preview.html', schedule=schedule, loan_info=loan_info, filename=filename)

            

    else:
        loan_data_file = request.files['loan_data']
        loan_filename = secure_filename(loan_data_file.filename)
        loan_path = os.path.join("uploads", loan_filename)
        loan_data_file.save(loan_path)

        output_filename = f"{base_name}_repayment_data.xlsx"
        output_path = os.path.join("output", output_filename)

        # Call your existing file-based export function
        export_to_excel(loan_path, txn_path, days_before, days_after, output_path)
        return send_file(output_path, as_attachment=True)



    








if __name__=='__main__':
    app.run(debug=True, port=5052)