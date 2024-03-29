import Add_manex_pn
import MMD_EditPN
import PCB_EditPN
import csv_to_mmd_converter
import report_generator
import re
import os
import shutil
import logging
from flask_session import Session
from flask import Flask, render_template, request, send_from_directory, session

logging.basicConfig(level=logging.DEBUG, filename="Excel_logfile.txt", filemode="a+",
                    format="%(asctime)-15s %(levelname)-8s %(message)s")

app = Flask(__name__)
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
Session(app)


@app.route('/', methods=["GET", "POST"])
def index():
    return render_template('index.html')


@app.route('/add_manexpn', methods=["GET", "POST"])
def add_manexpn():
    return render_template('add_manexpn.html', error="no")


@app.route('/error_tab', methods=["GET", "POST"])
def error_tab():
    return render_template('error_tab.html', error=session["error"], possibilities=session["possibilities"])


@app.route('/user_manual', methods=["GET", "POST"])
def user_manual():
    return render_template('user_manual.html')


@app.route('/processing', methods=["GET", "POST"])
def processing():
    possible_errors = ["Make sure that the BOM is the first sheet of the workbook",
                       "Check if the file extension is either .xls or .xlsx",
                       "The RefDes is recommended to be comma separated",
                       "The column names in Web Manex BOM should consists RefDesg, QtyEach, item_no and PART_NO",
                       "Start row in Manex BOM should be 2",
                       "Make sure you enter correct Customer BOM details"]
    session["possibilities"] = possible_errors
    if request.form.get("submit"):
        session["task"] = 1
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
        session['CustomerBom'] = customer_bom.filename
        session['ManexBom'] = manex_bom.filename
        result, error = Add_manex_pn.CheckBom(cust_file_path, man_file_path, designator, quantity, int(start_row), int(end_row), delimiter, separator)
        if result is True:
            return render_template('output.html', customer_bom=customer_bom.filename, manex_bom=manex_bom.filename)
        elif bool(result):
            session["customer_bom"] = cust_file_path
            session["manex_bom"] = man_file_path
            session['bom_data'] = result[1]
            session['file_extension'] = result[2]
            session['bom_col_des'] = result[3]
            session['start_row'] = int(start_row)
            session['end_row'] = int(end_row)
            session['sep_count'] = 0
            session['sep_len'] = len(result[0])
            session['sep_dict'] = {}
            session['sep_detail'] = result[0]
            session['sep_position'] = result[4]
            session['manex_dict'] = result[5]
            if session['sep_count'] < session['sep_len']:
                return render_template('separator_check.html', item=session['sep_detail'][session['sep_count']][0])
            else:
                result, error = Add_manex_pn.CustBomInfo(session['sep_detail'], session['bom_data'], session["customer_bom"],
                                                  session["manex_bom"], session['file_extension'], session['start_row'],
                                                  session['end_row'], session['bom_col_des'], session["sep_dict"], session['sep_position'], session['manex_dict'])
                if result is True:
                    return render_template('output.html', customer_bom=session['CustomerBom'],
                                           manex_bom=session['ManexBom'])
                else:
                    session["error"] = error
                    return render_template('add_manexpn.html', error="notsafe")
        else:
            if type(error) == list:
                session["error"] = error[0]
                session["possibilities"] = error[1]
            else:
                session["error"] = error
            return render_template('add_manexpn.html', error="notsafe")
    else:
        return index()


@app.route('/separator_check', methods=["GET", "POST"])
def separator_check():
    return render_template('separator_check.html')


@app.route('/separator_detail', methods=["GET", "POST"])
def separator_detail():
    return render_template('separator_detail.html')


@app.route('/check', methods=["GET", "POST"])
def check():
    positive = request.form.get("yes")
    if positive is not None:
        return render_template("separator_info.html", item=session['sep_detail'][session['sep_count']][0])
    elif session['sep_count'] + 1 < session['sep_len']:
        session['sep_count'] += 1
        return render_template('separator_check.html', item=session['sep_detail'][session['sep_count']][0])
    else:
        if session["task"] == 1:
            result, error = Add_manex_pn.CustBomInfo(session['sep_detail'], session['bom_data'], session["customer_bom"], session["manex_bom"], session['file_extension'], session['start_row'], session['end_row'], session['bom_col_des'], session["sep_dict"], session['sep_position'], session['manex_dict'])
            if result is True:
                return render_template('output.html', customer_bom=session['CustomerBom'], manex_bom=session['ManexBom'])
            else:
                session["error"] = error
                return render_template('add_manexpn.html', error="notsafe")
        elif session["task"] == 2:
            result, error = MMD_EditPN.CustBomInfo(session['sep_detail'], session['bom_data'], session["customer_bom"], session["bot_mmd"], session["top_mmd"])
            if result is True:
                return render_template('mmd_output.html', customer_bom=session['CustomerBom'],
                                       mmd_file1=session["botmmd"], mmd_file2=session["topmmd"])
            else:
                session["error"] = error
                return render_template('add_manexpn_to_mmd.html', error="notsafe")
        elif session["task"] == 3:
            result, error = PCB_EditPN.CustBomInfo(session['sep_detail'], session['bom_data'], session["customer_bom"],
                                                session["bot_pcb"], session["top_pcb"])
            if result is True:
                return render_template('pcb_output.html', customer_bom=session['CustomerBom'],
                                       pcb_file1=session["botpcb"], pcb_file2=session["toppcb"])
            else:
                session["error"] = error
                return render_template('edit_PN_in_pcb.html', error="notsafe")


@app.route('/sep_verify', methods=["GET", "POST"])
def sep_verify():
    positive = request.form.get("yes")
    if positive is not None:
        session["sep_dict"].update({session['sep_detail'][session['sep_count']][0]: session['result']})
        session['sep_detail'][session['sep_count']].append(session['result'])
        session['sep_detail'][session['sep_count']][2] = 1
        if session['sep_count'] + 1 < session['sep_len']:
            session['sep_count'] += 1
            return render_template('separator_check.html', item=session['sep_detail'][session['sep_count']][0])
        else:
            if session["task"] == 1:
                result, error = Add_manex_pn.CustBomInfo(session['sep_detail'], session['bom_data'], session["customer_bom"],
                                                  session["manex_bom"], session['file_extension'], session['start_row'],
                                                  session['end_row'], session['bom_col_des'], session["sep_dict"], session['sep_position'], session['manex_dict'])
                if result is True:
                    return render_template('output.html', customer_bom=session['CustomerBom'],
                                           manex_bom=session['ManexBom'])
                else:
                    session["error"] = error
                    return render_template('add_manexpn.html', error="notsafe")
            elif session["task"] == 2:
                result, error = MMD_EditPN.CustBomInfo(session['sep_detail'], session['bom_data'], session["customer_bom"],
                                                session["bot_mmd"], session["top_mmd"])
                if result is True:
                    return render_template('mmd_output.html', customer_bom=session['CustomerBom'],
                                           mmd_file1=session["botmmd"], mmd_file2=session["topmmd"])
                else:
                    session["error"] = error
                    return render_template('add_manexpn_to_mmd.html', error="notsafe")
            elif session["task"] == 3:
                result, error = PCB_EditPN.CustBomInfo(session['sep_detail'], session['bom_data'],
                                                       session["customer_bom"],
                                                       session["bot_pcb"], session["top_pcb"])
                if result is True:
                    return render_template('pcb_output.html', customer_bom=session['CustomerBom'],
                                           pcb_file1=session["botpcb"], pcb_file2=session["toppcb"])
                else:
                    session["error"] = error
                    return render_template('edit_PN_in_pcb.html', error="notsafe")
    else:
        return render_template("separator_info.html", item=session['sep_detail'][session['sep_count']][0])


@app.route('/separator_info', methods=["GET", "POST"])
def separator_info():
    return render_template('separator_info.html')


@app.route('/sep_detail', methods=["GET", "POST"])
def sep_detail():
    base = request.form.get("base")
    range_from = request.form.get("rfrom")
    range_to = request.form.get("rto")
    res = []
    my_list1 = list(filter(None, re.split(r'(\d+)', range_from)))
    my_list2 = list(filter(None, re.split(r'(\d+)', range_to)))
    if base.isalpha():
        base = base.upper()
    elif not base.isnumeric():
        base_list = list(base)
        for val in base_list:
            if val.isalpha():
                base_list[base_list.index(val)] = val.upper()
        base = "".join(base_list)
    if len(my_list1) > 1:
        if my_list1[0].isalpha():
            for i in range(ord(my_list1[0]), ord(my_list2[0]) + 1):
                for j in range(int(my_list1[1]), int(my_list2[1]) + 1):
                    res.append(f"{base}{chr(i)}{j}")
        else:
            for i in range(int(my_list1[1]), int(my_list2[1]) + 1):
                for j in range(ord(my_list1[0]), ord(my_list2[0]) + 1):
                    res.append(f"{base}{i}{chr(j)}")
    else:
        if my_list1[0].isalpha():
            for i in range(ord(my_list1[0]), ord(my_list2[0]) + 1):
                res.append(f"{base}{chr(i)}")
        else:
            for j in range(int(my_list1[0]), int(my_list2[0]) + 1):
                res.append(f"{base}{j}")
    session['result'] = res
    return render_template("separator_detail.html", item=session['sep_detail'][session['sep_count']][0], values=res)


@app.route('/download/<path:filename>', methods=['GET'])
def download(filename):
    path = "upload_folder"
    # file = {
    #     "customer": session["CustomerBom"],
    #     "manex": session["ManexBom"],
    #     "mmd_file": session["MmdFile"],
    #     "mmdfile1": session["botmmd"],
    #     "mmdfile2": session["topmmd"],
    #     "cad_file": session["cadfile"],
    #     "asc_file": session["ascfile"],
    #     "pcbfile1": session["botpcb"],
    #     "pcbfile2": session["toppcb"],
    # }
    if filename == "customer":
        filename = session["CustomerBom"]
    elif filename == "manex":
        filename = session["ManexBom"]
    elif filename == "mmd_file":
        filename = session["MmdFile"]
    elif filename == "mmdfile1":
        filename = session["botmmd"]
    elif filename == "mmdfile2":
        filename = session["topmmd"]
    elif filename == "cad_file":
        filename = session["cadfile"]
    elif filename == "asc_file":
        filename = session["ascfile"]
    elif filename == "pcbfile1":
        filename = session["botpcb"]
    elif filename == "pcbfile2":
        filename = session["toppcb"]
    elif filename == "model":
        path = f"{path}/{session['model']}"
        if os.path.exists(f'{path}.zip'):
            os.remove(f'{path}.zip')
        shutil.make_archive(f'upload_folder/{session["model"]}', 'zip', path)
        path = 'upload_folder'
        filename = f'{session["model"]}.zip'
    # else:
    #     filename = file[filename]
    return send_from_directory(path, filename, as_attachment=True)


@app.route('/csv_to_mmd', methods=["GET", "POST"])
def csv_to_mmd():
    return render_template('csv_to_mmd.html')


@app.route('/mmd_creator', methods=["GET", "POST"])
def mmd_creator():
    possible_errors = ["Make sure you are using the right csv format",
                       "Actaul data should start from row 2",
                       "Column B should have internal part numbers whereas Column C should have RefDes",
                       "X-axis, Y-axis and Rotation Angle must on column H, I and J, respectively"]
    session["possibilities"] = possible_errors
    csv_file = request.files['csv_file']
    csv_file_path = f"upload_folder/{csv_file.filename}"
    csv_file.save(csv_file_path)
    result, mmd_file, error = csv_to_mmd_converter.convert(csv_file_path)
    session['MmdFile'] = mmd_file
    if result is True:
        return render_template('csv_to_mmd_output.html', mmd_file=mmd_file)
    else:
        session["error"] = error
        return render_template('csv_to_mmd.html', error="notsafe")


@app.route('/add_manexpn_to_mmd', methods=["GET", "POST"])
def add_manexpn_to_mmd():
    return render_template('add_manexpn_to_mmd.html')


@app.route('/mmd_editor', methods=["GET", "POST"])
def mmd_editor():
    possible_errors = ["Make sure that the format of mmd file has not changed",
                       "Make sure the Customer BOM is the first sheet of the workbook",
                       "The RefDes is recommended to be comma separated in Customer BOM",
                       "Verify the correctness of the inputs about the BOM"]
    session["possibilities"] = possible_errors
    if request.form.get("submit"):
        session["task"] = 2
        customer_bom = request.files['customer_bom']
        mmd_file = request.files['mmd_file']
        top_mmd = request.files['top_mmd']
        bot_mmd = request.files['bot_mmd']
        cust_file_path = f"upload_folder/{customer_bom.filename}"
        customer_bom.save(cust_file_path)
        designator = request.form.get("designator")
        pn = request.form.get("pn")
        start_row = int(request.form.get("start_row"))
        end_row = int(request.form.get("end_row"))
        des = {'comma': ',', 'hyphen': '-', 'space': ' ', 'other': request.form.get('del')}
        sep = {'none': None, 'comma': ',', 'hyphen': '-', 'space': ' ', 'colon': ':', 'other': request.form.get('sep')}
        delimiter = des.get(request.form.get("delimiter"))
        separator = sep.get(request.form.get("separator"))
        session['CustomerBom'] = customer_bom.filename
        if bool(mmd_file):
            mmd_file_path = f"upload_folder/{mmd_file.filename}"
            mmd_file.save(mmd_file_path)
            session['botmmd'] = mmd_file.filename
            session['botmmd'] = session['botmmd'].replace(".mmd", "_Svt_PartNo.mmd")
            session['botmmd'] = session['botmmd'].replace(".MMD", "_Svt_PartNo.mmd")
            result, error = MMD_EditPN.ReadBom(cust_file_path, designator, pn, int(start_row), int(end_row), delimiter, separator, mmd_file_path, False)
            if result is True:
                return render_template('mmd_output.html', customer_bom=customer_bom.filename, mmd_file1=session['botmmd'],
                                       mmd_file2=False)
            if error is not None:
                session["error"] = error
                return render_template('add_manexpn_to_mmd.html', error="notsafe")
        else:
            top_mmd_path = f"upload_folder/{top_mmd.filename}"
            top_mmd.save(top_mmd_path)
            session['topmmd'] = top_mmd.filename
            bot_mmd_path = f"upload_folder/{bot_mmd.filename}"
            bot_mmd.save(bot_mmd_path)
            session['botmmd'] = bot_mmd.filename
            session['botmmd'] = session['botmmd'].replace(".mmd", "_Svt_PartNo.mmd")
            session['botmmd'] = session['botmmd'].replace(".MMD", "_Svt_PartNo.mmd")
            session['topmmd'] = session['topmmd'].replace(".mmd", "_Svt_PartNo.mmd")
            session['topmmd'] = session['topmmd'].replace(".MMD", "_Svt_PartNo.mmd")
            result, error = MMD_EditPN.ReadBom(cust_file_path, designator, pn, int(start_row), int(end_row), delimiter, separator, bot_mmd_path, top_mmd_path)
            if result is True:
                return render_template('mmd_output.html', customer_bom=customer_bom.filename, mmd_file1=session['botmmd'], mmd_file2=session['topmmd'])
            if error is not None:
                session["error"] = error
                return render_template('add_manexpn_to_mmd.html', error="notsafe")
        if bool(result):
            session["customer_bom"] = cust_file_path
            if bool(mmd_file):
                session["bot_mmd"] = mmd_file_path
                session["top_mmd"] = False
            else:
                session["bot_mmd"] = bot_mmd_path
                session["top_mmd"] = top_mmd_path
            session['bom_data'] = result[1]
            session['bom_col_des'] = result[3]
            session['start_row'] = int(start_row)
            session['end_row'] = int(end_row)
            session['sep_count'] = 0
            session['sep_len'] = len(result[0])
            session['sep_detail'] = result[0]
            session['sep_dict'] = {}
            if session['sep_count'] < session['sep_len']:
                return render_template('separator_check.html', item=session['sep_detail'][session['sep_count']][0])
            else:
                result, error = MMD_EditPN.CustBomInfo(session['sep_detail'], session['bom_data'], session["customer_bom"],
                                                session["bot_mmd"], session["top_mmd"])
                if result is True:
                    return render_template('mmd_output.html', customer_bom=session['CustomerBom'],
                                           mmd_file1=session["botmmd"], mmd_file2=session["topmmd"])
                else:
                    session["error"] = error
                    return render_template('add_manexpn_to_mmd.html', error="notsafe")
    else:
        return index()


@app.route('/mmd_output', methods=["GET", "POST"])
def mmd_output():
    return render_template('mmd_output.html')


@app.route('/gencad_editor', methods=["GET", "POST"])
def gencad_editor():
    return render_template('gencad_editor.html')


@app.route('/gencad_output', methods=["GET", "POST"])
def gencad_output():
    possible_errors = ["Check if there are any format changes"]
    session["possibilities"] = possible_errors
    cad_file = request.files['cad_file']
    cad_file_path = f"upload_folder/{cad_file.filename}"
    cad_file.save(cad_file_path)
    change_count = 0
    try:
        gencad = open(cad_file_path, 'r')
        file_data = gencad.readlines()
        start = 0
        stop = 0
        count = 0
        for line in file_data:
            if line == "$COMPONENTS\n":
                start = count
            if line == "$ENDCOMPONENTS\n":
                stop = count
            count += 1

        change_count = 0
        for line in range(start + 1, stop):
            match = re.match("^COMPONENT\s", file_data[line])
            if match:
                if "-" in file_data[line]:
                    item = file_data[line].strip().split(" ")
                    value = item.pop()
                    edit = value.replace("-", "_")
                    item.append(edit)
                    item = " ".join(item)
                    file_data[line] = item + "\n"
                    change_count += 1
        gencad.close()
        var = gencad.name.split("/")
        new_file_name = var.pop().replace(".cad", "_modified.cad")
        new_file_name = new_file_name.replace(".CAD", "_modified.CAD")
        var.append(new_file_name)
        new_file = "/".join(var)

        with open(new_file, "w") as f:
            for item in file_data:
                f.write("%s" % item)
        f.close()
        result = True

    except Exception as e:
        logging.error(f"{e.__class__}")
        logging.error("Error while editing .cad file!")
        logging.error(f"{e}")
        result = False

    if result is True:
        session['cadfile'] = new_file_name
        return render_template('gencad_output.html', gencad=session['cadfile'], count=change_count)
    else:
        return render_template('gencad_editor.html', error="notsafe")


@app.route('/ascii_editor', methods=["GET", "POST"])
def ascii_editor():
    return render_template('ascii_editor.html')


@app.route('/ascii_output', methods=["GET", "POST"])
def ascii_output():
    possible_errors = ["Check if there are any format changes"]
    session["possibilities"] = possible_errors
    ascii_file = request.files['ascii_file']
    ascii_file_path = f"upload_folder/{ascii_file.filename}"
    ascii_file.save(ascii_file_path)
    change_count = 0
    try:
        ascfile = open(ascii_file_path, 'r')
        file_data = ascfile.readlines()
        start = 0
        stop = 0
        count = 0
        for line in file_data:
            start_match = re.match(r"^\*PART\*\s", line)
            if start_match:
                print(line)
                start = count
                print(start)
            stop_match = re.match(r"^\*ROUTE\*\s", line)
            if stop_match:
                print(line)
                stop = count
                print(stop)
            count += 1

        change_count = 0
        for line in range(start + 1, stop):
            item = file_data[line].strip().split(" ")
            value = item[0]
            if "-" in value:
                value = value.replace("-", "_")
                item[0] = value
            item = " ".join(item)
            file_data[line] = item + "\n"
            change_count += 1

        ascfile.close()

        var = ascfile.name.split("/")
        temp = var.pop()
        new_file_name = temp.replace(".ASC", "_modified.ASC")
        new_file_name = temp.replace(".asc", "_modified.asc")
        var.append(new_file_name)
        new_file = "/".join(var)

        with open(new_file, "w") as f:
            for item in file_data:
                f.write("%s" % item)
        f.close()
        result = True

    except Exception as e:
        logging.error(f"{e.__class__}")
        logging.error("Error while editing .ASC file!")
        logging.error(f"{e}")
        result = False
        session["error"] = f"Error while editing .ASC file\n{e.__class__}\n{e}"

    if result is True:
        session['ascfile'] = new_file_name
        return render_template('ascii_output.html', ascii=session['ascfile'], count=change_count)
    else:
        return render_template('ascii_editor.html', error="notsafe")


@app.route('/edit_PN_in_pcb', methods=["GET", "POST"])
def edit_PN_in_pcb():
    return render_template('edit_PN_in_pcb.html')


@app.route('/pcb_editor', methods=["GET", "POST"])
def pcb_editor():
    possible_errors = ["Make sure that the format of pcb file has not changed",
                       "Make sure the Customer BOM is the first sheet of the workbook",
                       "The RefDes is recommended to be comma separated in Customer BOM",
                       "Verify the correctness of the inputs about the BOM"]
    session["possibilities"] = possible_errors
    if request.form.get("submit"):
        session["task"] = 3
        customer_bom = request.files['customer_bom']
        pcb_file = request.files['pcb_file']
        top_pcb = request.files['top_pcb']
        bot_pcb = request.files['bot_pcb']
        cust_file_path = f"upload_folder/{customer_bom.filename}"
        customer_bom.save(cust_file_path)
        designator = request.form.get("designator")
        pn = request.form.get("pn")
        start_row = int(request.form.get("start_row"))
        end_row = int(request.form.get("end_row"))
        des = {'comma': ',', 'hyphen': '-', 'space': ' ', 'other': request.form.get('del')}
        sep = {'none': None, 'comma': ',', 'hyphen': '-', 'space': ' ', 'colon': ':', 'other': request.form.get('sep')}
        delimiter = des.get(request.form.get("delimiter"))
        separator = sep.get(request.form.get("separator"))
        session['CustomerBom'] = customer_bom.filename
        if bool(pcb_file):
            pcb_file_path = f"upload_folder/{pcb_file.filename}"
            pcb_file.save(pcb_file_path)
            session['botpcb'] = pcb_file.filename
            session['botpcb'] = session['botpcb'].replace(".pcb", "_modified.pcb")
            session['botpcb'] = session['botpcb'].replace(".PCB", "_modified.pcb")
            result, error = PCB_EditPN.ReadBom(cust_file_path, designator, pn, int(start_row), int(end_row), delimiter, separator, pcb_file_path, False)
            if result is True:
                return render_template('pcb_output.html', customer_bom=customer_bom.filename, pcb_file1=session['botpcb'],
                                       pcb_file2=False)
            if error is not None:
                session["error"] = error
                return render_template('edit_PN_in_pcb.html', error="notsafe")
        else:
            top_pcb_path = f"upload_folder/{top_pcb.filename}"
            top_pcb.save(top_pcb_path)
            session['topmmd'] = top_pcb.filename
            bot_pcb_path = f"upload_folder/{bot_pcb.filename}"
            bot_pcb.save(bot_pcb_path)
            session['botpcb'] = bot_pcb.filename
            session['botpcb'] = session['botpcb'].replace(".pcb", "_modified.pcb")
            session['botpcb'] = session['botpcb'].replace(".PCB", "_modified.pcb")
            session['toppcb'] = session['toppcb'].replace(".pcb", "_modified.pcb")
            session['toppcb'] = session['toppcb'].replace(".PCB", "_modified.pcb")
            result, error = PCB_EditPN.ReadBom(cust_file_path, designator, pn, int(start_row), int(end_row), delimiter, separator, bot_pcb_path, top_pcb_path)
            if result is True:
                return render_template('pcb_output.html', customer_bom=customer_bom.filename, pcb_file1=session['botpcb'], pcb_file2=session['toppcb'])
            if error is not None:
                session['error'] = error
                return render_template('edit_PN_in_pcb.html', error="notsafe")
        if bool(result):
            session["customer_bom"] = cust_file_path
            if bool(pcb_file):
                session["bot_pcb"] = pcb_file_path
                session["top_pcb"] = False
            else:
                session["bot_pcb"] = bot_pcb_path
                session["top_pcb"] = top_pcb_path
            session['bom_data'] = result[1]
            session['bom_col_des'] = result[3]
            session['start_row'] = int(start_row)
            session['end_row'] = int(end_row)
            session['sep_count'] = 0
            session['sep_len'] = len(result[0])
            session['sep_detail'] = result[0]
            session['sep_position'] = {}
            if session['sep_count'] < session['sep_len']:
                return render_template('separator_check.html', item=session['sep_detail'][session['sep_count']][0])
            else:
                result, error = PCB_EditPN.CustBomInfo(session['sep_detail'], session['bom_data'], session["customer_bom"],
                                                session["bot_pcb"], session["top_pcb"])
                if result is True:
                    return render_template('pcb_output.html', customer_bom=session['CustomerBom'],
                                           pcb_file1=session["botpcb"], pcb_file2=session["toppcb"])
                else:
                    session["error"] = error
                    return render_template('edit_PN_in_pcb.html', error="notsafe")
    else:
        return index()


@app.route('/pcb_output', methods=["GET", "POST"])
def pcb_output():
    return render_template('pcb_output.html')


@app.route('/aoi_report_generator', methods=["GET", "POST"])
def aoi_report_generator():
    return render_template('aoi_report_generator.html')


@app.route('/reports_output', methods=["GET", "POST"])
def reports_output():
    possible_errors = ["Check all your inputs for correctness", "Verify that the inputs are as per the required formats", "Time interval format is HH:MM:SS", "Date format is MM/DD/YYYY"]
    session["possibilities"] = possible_errors
    if request.form.get("submit"):
        batch = request.form.get("batch")
        time = request.form.get("time")
        serial = int(request.form.get("serial"))
        copies = int(request.form.get("copies"))
        model = request.form.get("model")
        boardsn = int(request.form.get("boardsn"))
        date = request.form.get("date")
        time_int = request.form.get("time_int")
        mount = int(request.form.get("mount"))
        solder = int(request.form.get("solder"))
        ic_count = int(request.form.get("ic_count"))
        total = int(request.form.get("total"))

        date = date.split("T")
        ddate = date[0].split("-")
        ddate = f"{ddate[1]}/{ddate[0]}/{ddate[2]}"
        dtime = date[1]

        parent_dir = "upload_folder/"
        directory = f"{model}"
        path = os.path.join(parent_dir, directory)
        if os.path.exists(path):
            shutil.rmtree(path)
        os.mkdir(path)
        session["model"] = model
        session['copies'] = copies

        if batch == "single":
            result, error = report_generator.single_batch(serial, copies, model, boardsn, ddate, dtime, time_int, mount, solder, ic_count, total)
        else:
            result, error = report_generator.batchwise(serial, copies, time, model, boardsn, ddate, dtime, time_int, mount, solder, ic_count, total)

        if result is True:
            return render_template('reports_output.html', model_id=model)
        else:
            session["error"] = error
            return render_template('aoi_report_generator.html', error="notsafe")
    else:
        return index()


if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
