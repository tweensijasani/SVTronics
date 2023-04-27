import sys
import xlrd
import copy
import logging
import pathlib
import tkinter as tk
from win32com import client
from tkinter import filedialog, messagebox

logging.basicConfig(level=logging.DEBUG, filename="HF_logfile.txt", filemode="a+",
                    format="%(asctime)-15s %(levelname)-8s %(message)s")

root = tk.Tk()
root.withdraw()


def getfile():
    try:
        counter = 0
        Bomfile = filedialog.askopenfilename(title="Select HF CONTROL BOM", filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")))
        while Bomfile == '' and counter < 1:
            messagebox.showerror(title="File Error", message="Customer BOM Not Selected")
            logging.warning("Manex BOM Not Selected")
            Bomfile = filedialog.askopenfilename(title="Select HF CONTROL BOM",
                                                    filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xls")))
            counter += 1
        if Bomfile == '':
            messagebox.showerror(title="Invalid Input", message="Something went wrong!! Please try again....")
            logging.error("Terminated: Customer BOM not selected!")
            sys.exit(1)
        else:
            logging.info("Customer BOM Excel Selected")
            return Bomfile

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__}")
        logging.error("Error while fetching file!")
        logging.error(f"{e}")
        print(e, "\n Error while fetching file!")
        sys.exit(1)


def read_bom(Bomfile):
    try:
        file_extension = pathlib.Path(Bomfile).suffix
        logging.info("Reading Customer BOM Excel...")
        if file_extension == ".xls" or file_extension == ".XLS":
            wb_bom = xlrd.open_workbook(Bomfile)
            ws_bom = wb_bom.sheet_by_index(0)
            end = ws_bom.nrows
            start = 0

            for row in range(0, end):
                var = ws_bom.row_values(row)
                if var == ['Seq No', 'Part No', 'Item Desc', '', 'QTY Per', 'B/D Rev', 'Ref Designators', 'notes']:
                    start = row + 1
                    break

            if start == 0:
                messagebox.showerror(title="Not Compatible", message="The BOM format does not match Standard HF CONTROL BOM!!")
                sys.exit(1)

            current_row = []
            pointer = 0
            count = 0
            final_data = [["Seq No", "Part No", "Item Desc", "QTY Per", "B/D Rev", "Ref Designators", "Mfr", "Mpn", "Notes"]]
            for row in range(start, end):
                var = ws_bom.row_values(row)
                if bool(var[0]):
                    RefDes = (var[6].strip()).split(" ")
                    current_row = [var[0], var[1].strip(), var[2].strip(), var[4], var[5].strip(), RefDes]
                    pointer = copy.copy(count)

                elif bool(var[6]):
                    RefDes = (var[6].strip()).split(" ")
                    for x in range(pointer + 1, len(final_data)):
                        final_data[x][5].extend(RefDes)
                    current_row = final_data[count]

                if bool(var[7]):
                    note, mnf, mno = note_formatting(var[7])
                    if bool(mnf) and bool(mno):
                        for item in range(0, min(len(mnf), len(mno))):
                            row_data = copy.copy(current_row)
                            row_data.append(mnf[item])
                            row_data.append(mno[item])
                            final_data.append(row_data)
                            count += 1
                    else:
                        row_data = copy.copy(current_row)
                        row_data.append("")
                        row_data.append("")
                        final_data.append(row_data)
                        count += 1
                    if bool(note):
                        final_data[pointer + 1].append(" \n".join(note))

            return final_data

        else:
            messagebox.showerror(title="Not Compatible",
                                 message="Please use .xls BOM!!")
            sys.exit(1)

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__} from line 45")
        logging.error("Error while reading Customer BOM!")
        logging.error(f"{e}")
        print(e, "\n Error while reading Customer BOM!")
        sys.exit(1)


def write_bom(final_data):
    try:
        xlApp = client.Dispatch("Excel.Application")
        wkbk = xlApp.Workbooks.open(Bomfile)
        wksht = wkbk.Sheets.Add(Before=None, After=wkbk.Sheets(wkbk.Sheets.count))

        for i in range(0, len(final_data)):
            for j in range(1, len(final_data[i]) + 1):
                if type(final_data[i][j - 1]) == list:
                    wksht.Cells(i + 1, j).Value = ", ".join(final_data[i][j - 1])
                else:
                    wksht.Cells(i + 1, j).Value = final_data[i][j - 1]

        wkbk.Save()
        wkbk.Close(True)
        xlApp.Quit()

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__} from line 105")
        logging.error("Error while writing Customer BOM!")
        logging.error(f"{e}")
        print(e, "\n Error while writing Customer BOM!")
        if bool(wkbk):
            wkbk.Close()
            xlApp.Quit()
        sys.exit(1)


def note_formatting(raw_note):
    try:
        notes = raw_note.split("\n")
        note = []
        mnf = []
        mno = []
        for item in notes:
            item = item.split(" ")
            temp1 = []
            temp2 = []
            if item[0].replace(")", "").isnumeric():
                flag = 0
                for x in range(1, len(item)):
                    if flag == 0:
                        for y in item[x]:
                            if y.isnumeric():
                                flag = 1
                                break
                    if flag == 0:
                        temp1.append(item[x])
                    else:
                        temp2.append(item[x])
                mnf.append(" ".join(temp1))
                mno.append(" ".join(temp2))
            else:
                note.append(" ".join(item))

        return note, mnf, mno

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__} from line 242")
        logging.error("Error while interpreting data!")
        logging.error(f"{e}")
        print(e, "\n Error while interpreting data!")
        sys.exit(1)


if __name__ == "__main__":
    logging.info("Execution Started...")
    print("Execution Started!!!")

    Bomfile = getfile()
    final_data = read_bom(Bomfile)
    write_bom(final_data)

    logging.info("Successfully Executed!!!\n\n")
    print("Successfully Executed!!!")
    messagebox.showinfo(title="Status", message="Completed!!!")