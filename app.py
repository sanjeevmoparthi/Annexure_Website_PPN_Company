from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import pandas as pd
from io import BytesIO
import annexure_functions as annex
import os

ALLOWED_EXT = {'xls', 'xlsx'}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 MB limit

# app = Flask(__name__)
app = Flask(__name__, static_folder='static', template_folder='templates')

app.secret_key = "change_this_in_prod"
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


@app.route('/')
def index():
    return render_template('index.html')


# -------------------------------------------------------
#                     ANNEXURE – 1
# -------------------------------------------------------
@app.route('/download/1', methods=['POST'])
def download_annexure1():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please select a file to upload", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        
        df = pd.read_excel(file, engine="openpyxl", read_only=True)

        out_io = annex.annexure1_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure1_Vendor_Wise_Margin.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-1: {str(e)}", "danger")
        return redirect(url_for('index'))


# -------------------------------------------------------
#                     ANNEXURE – 2
# -------------------------------------------------------
@app.route('/download/2', methods=['POST'])
def download_annexure2():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please select a file to upload", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        
        df = pd.read_excel(file, engine="openpyxl", read_only=True)

        out_io = annex.annexure2_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure2_Brand_Wise_Margin.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-2: {str(e)}", "danger")
        return redirect(url_for('index'))


# -------------------------------------------------------
#                     ANNEXURE – 3
# -------------------------------------------------------
@app.route('/download/3', methods=['POST'])
def download_annexure3():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please select a file to upload", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        
        df = pd.read_excel(file, engine="openpyxl", read_only=True)

        out_io = annex.annexure3_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure3_Brand_Wise_Sales.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-3: {str(e)}", "danger")
        return redirect(url_for('index'))
    
# -------------------------------------------------------
#                     ANNEXURE – 4
# -------------------------------------------------------
@app.route('/download/4', methods=['POST'])
def download_annexure4():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please select a file to upload", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        
        df = pd.read_excel(file, engine="openpyxl", read_only=True)

        out_io = annex.annexure4_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure4_Product_Wise_Sales_Summary.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-4: {str(e)}", "danger")
        return redirect(url_for('index'))
    

# -------------------------------------------------------
#                     ANNEXURE – 5
# -------------------------------------------------------
@app.route('/download/5', methods=['POST'])
def download_annexure5():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please select a file to upload", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        
        df = pd.read_excel(file, engine="openpyxl", read_only=True)

        out_io = annex.annexure5_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure5_Product_Category_Contribution.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-5: {str(e)}", "danger")
        return redirect(url_for('index'))
    
@app.route('/download/6', methods=['POST'])
def download_annexure6():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index'))

    try:
        
        df = pd.read_excel(file, engine="openpyxl", read_only=True)

        out_io = annex.annexure6_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure6_Negative_Margin.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"Error generating Annexure-6: {str(e)}", "danger")
        return redirect(url_for('index'))

# -------------------------------------------------------
#                     ANNEXURE – 7
# -------------------------------------------------------
@app.route('/download/7', methods=['POST'])
def download_annexure7():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        
        df = pd.read_excel(file, engine="openpyxl", read_only=True)

        out_io = annex.annexure7_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure7_Profit_Below_10.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-7: {str(e)}", "danger")
        return redirect(url_for('index'))

# -------------------------------------------------------
#                     ANNEXURE – 8
# -------------------------------------------------------
@app.route('/download/8', methods=['POST'])
def download_annexure8():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        
        df = pd.read_excel(file, engine="openpyxl", read_only=True)

        out_io = annex.annexure8_generate_excel_bytes(df)

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure8_Neither_Profit_Nor_Loss.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-8: {str(e)}", "danger")
        return redirect(url_for('index'))

# -------------------------------------------------------
#                     ANNEXURE – 9
# -------------------------------------------------------
@app.route('/download/9', methods=['POST'])
def download_annexure9():
    file = request.files.get('file')
    if not file or file.filename == '':
        flash("Please upload a file", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Only .xls/.xlsx allowed", "danger")
        return redirect(url_for('index'))

    try:
        
        df = pd.read_excel(file, engine="openpyxl", read_only=True)

        out_io = annex.annexure9_generate_excel_bytes(df)  # Make sure this function exists in annexure_functions

        return send_file(
            out_io,
            as_attachment=True,
            download_name="Annexure9_HighVendorMargin_LessProfitMargin.xlsx",  # Change name as per your requirement
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        flash(f"Error generating Annexure-9: {str(e)}", "danger")
        return redirect(url_for('index'))



# -------------------------------------------------------
#                     RUN APP
# -------------------------------------------------------
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
