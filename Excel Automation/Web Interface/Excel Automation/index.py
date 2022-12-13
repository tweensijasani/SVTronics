import Add_manex_pn
# from werkzeug.utils import secure_filename
from flask_session import Session
from flask import Flask, render_template, request, send_from_directory, session


app = Flask(__name__)
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
Session(app)
# flag = 0
CustomerBom = ''
ManexBom = ''


@app.route('/', methods=["GET", "POST"])
def index():
    # global flag
    return render_template('index.html', error="no")


@app.route('/user_manual', methods=["GET", "POST"])
def user_manual():
    return render_template('user_manual.html')


@app.route('/output', methods=["GET", "POST"])
def output():
    # global flag, CustomerBom, ManexBom
    # if flag == 0:
    #     flag = 1
    customer_bom = request.files['customer_bom']
    work_order = request.form.get("work_order")
    if work_order:
        customer_bom.filename = f"{work_order}_{customer_bom.filename}"
    cust_file_path = f"upload_folder/{customer_bom.filename}"
    customer_bom.save(cust_file_path)
    manex_bom = request.files['manex_bom']
    man_file_path = f"upload_folder/{manex_bom.filename}"
    manex_bom.save(man_file_path)
    designator = request.form.get("designator")
    quantity = request.form.get("quantity")
    start_row = int(request.form.get("start_row"))
    end_row = int(request.form.get("end_row"))
    des = {'comma': ',', 'hyphen': '-', 'space': ' ', 'other': request.form.get('del')}
    sep = {'none': None, 'comma': ',', 'hyphen': '-', 'space': ' ', 'colon': ':', 'other': request.form.get('sep')}
    delimiter = des.get(request.form.get("delimiter"))
    separator = sep.get(request.form.get("separator"))
    safe = Add_manex_pn.ReadCustBom(cust_file_path, man_file_path, designator, quantity, int(start_row), int(end_row), delimiter, separator)
    session["CustomerBom"] = customer_bom.filename
    session["ManexBom"] = manex_bom.filename
    # flag = 0
    if safe:
        return render_template('output.html', customer_bom=customer_bom.filename, manex_bom=manex_bom.filename)
    else:
        return render_template('index.html', error="notsafe")
    # else:
    #     return render_template('index.html', flag=flag, error="yes")


@app.route('/download/<path:filename>', methods=['GET'])
def download(filename):
    # global CustomerBom, ManexBom
    path = "upload_folder"
    if filename == "customer":
        filename = session["CustomerBom"]
    else:
        filename = session["ManexBom"]
    return send_from_directory(path, filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
