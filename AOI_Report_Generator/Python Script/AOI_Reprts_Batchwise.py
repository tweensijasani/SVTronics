import sys
import logging
import easygui.boxes
import tkinter as tk
from easygui import *
from datetime import date, datetime, timedelta
from tkinter import messagebox

logging.basicConfig(level=logging.DEBUG, filename="reports_log.txt", filemode="a+",
                    format="%(asctime)-15s %(levelname)-8s %(message)s")

root = tk.Tk()
root.withdraw()

easygui.boxes.global_state.PROPORTIONAL_FONT_FAMILY = ("MS", "Arial")
easygui.boxes.global_state.MONOSPACE_FONT_FAMILY = "Arial"
easygui.boxes.global_state.PROPORTIONAL_FONT_SIZE = 14
easygui.boxes.global_state.MONOSPACE_FONT_SIZE = 14
easygui.boxes.global_state.TEXT_ENTRY_FONT_SIZE = 14
easygui.boxes.global_state.prop_font_line_length = 30
easygui.boxes.global_state.fixw_font_line_length = 40


def reports():
    try:
        logging.info("Getting File Naming Info...")
        text = "Report Name Information"
        title = "Enter Details"
        input_list = ["Serial No.", "No. of copies"]
        output = multenterbox(text, title, input_list)

        while output[0] is None or output[1] is None or not output[1].isnumeric():
            messagebox.showerror(title="Invalid Input", message="Please enter valid text formats")
            text = "Report Name Information"
            title = "Enter Details"
            input_list = ["Serial No.", "No. of copies"]
            output = multenterbox(text, title, input_list)

        serialno = int(output[0].strip())
        copies = int(output[1].strip())
        logging.info("Naming info populated")

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__} from line 67")
        logging.error("Error while getting File Naming Inputs!")
        logging.error(f"{e}")
        print(e, "\n Error while getting File Naming Inputs!")
        sys.exit(1)

    try:
        logging.info("Getting Report Info...")
        text = "Report Details"
        title = "Enter Details"
        input_list = ["Model ID", "Board S/N", "Date(MM/DD/YYYY)", "Time(HH:MM:SS)", "Time Interval in Hours", "Time Interval in Minutes", "Time Interval in Seconds"]
        output = multenterbox(text, title, input_list)

        while output[0] is None or output[1] is None or output[2] is None or output[3] is None:
            messagebox.showerror(title="Invalid Input", message="Please enter valid text formats")
            text = "Report Details"
            title = "Enter Details"
            input_list = ["Model ID", "Board S/N", "Date(MM/DD/YYYY)", "Time(HH:MM:SS)", "Time Interval in Hours", "Time Interval in Minutes", "Time Interval in Seconds"]
            output = multenterbox(text, title, input_list)

        modelid = output[0].strip()
        boardsn = int(output[1].strip())
        rdate = output[2].strip()
        rtime = output[3].strip()
        dhours = int(output[4].strip()) if bool(output[4]) else 0
        dminutes = int(output[5].strip()) if bool(output[5]) else 0
        dseconds = int(output[6].strip()) if bool(output[6]) else 0
        logging.info("Report info populated")

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__} from line 67")
        logging.error("Error while getting report details!")
        logging.error(f"{e}")
        print(e, "\n Error while getting report details!")
        sys.exit(1)

    try:
        logging.info("Getting Report Output...")
        text = "Inspection Counts"
        title = "Enter Details"
        input_list = ["Mounting", "Solder", "IC", "Total"]
        output = multenterbox(text, title, input_list)

        while not output[0].isnumeric() or not output[1].isnumeric() or not output[2].isnumeric() or not output[3].isnumeric():
            messagebox.showerror(title="Invalid Input", message="Please enter valid text formats")
            text = "Inspection Counts"
            title = "Enter Details"
            input_list = ["Mounting", "Solder", "IC", "Total"]
            output = multenterbox(text, title, input_list)

        mounting = int(output[0].strip()) if bool(output[0]) else 0
        solder = int(output[1].strip()) if bool(output[1]) else 0
        ic_count = int(output[2].strip()) if bool(output[2]) else 0
        total = int(output[3].strip()) if bool(output[3]) else 0
        logging.info("Inspection counts noted")

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__} from line 67")
        logging.error("Error while getting inspection counts!")
        logging.error(f"{e}")
        print(e, "\n Error while getting inspection counts!")
        sys.exit(1)

    try:
        logging.info("Creating report files...")
        tdate = date.today().strftime("%Y%m%d")
        ttime = datetime.now().strftime("%H%M%S")
        endsn = serialno + copies
        count = 0
        data_list = [["Model ID : ", modelid, "\n"],
                     ["Board S/N :", 0, ", DateTime : ", f"{rdate} ", 0, ", User : master\n"],
                     ["Machine : Local, Side : Top, Work Order :\n"],
                     ["  Summary Statistics : Mounting            Solder              IC                  Total\n"],
                     ["   Inspection Counts : ", 0, 0, 0, 0, "\n"],
                     ["         Fail Counts : 0                   0                   0                   0\n"],
                     ["              Yields : ", 0, 0, 0, 0, "\n"],
                     ["Inspection Result    : T\n"]]

        space_dict = {1: "                   ", 2: "                  ", 3: "                 ", 4: "                "}
        data_list[4][1] = f"{mounting}{space_dict[len(str(mounting))]}"
        data_list[4][2] = f"{solder}{space_dict[len(str(solder))]}"
        data_list[4][3] = f"{ic_count}{space_dict[len(str(ic_count))]}"
        data_list[4][4] = f"{total}"
        data_list[6][1] = "0.00 %              " if mounting == 0 else "100.00 %            "
        data_list[6][2] = "0.00 %              " if solder == 0 else "100.00 %            "
        data_list[6][3] = "0.00 %              " if ic_count == 0 else "100.00 %            "
        data_list[6][4] = "0.00 %" if total == 0 else "100.00 %"

        for x in range(serialno, endsn):
            filename = f"SimpleCoverageExport_{x}_{tdate}_{ttime}.txt"
            temp = f"{boardsn}"
            while len(temp) < 9:
                temp = f"0{temp}"
            data_list[1][1] = f"{temp}"
            data_list[1][4] = rtime
            data = data_list[0] + data_list[1] + data_list[2] + data_list[3] + data_list[4] + data_list[5] + data_list[6] + data_list[7]
            with open(filename, "w") as f:
                f.write("".join(data))
            f.close()
            rtime = datetime.strptime(rtime, "%H:%M:%S") + timedelta(hours=dhours, minutes=dminutes, seconds=dseconds)
            rtime = rtime.strftime("%H:%M:%S")
            boardsn += 1
            count += 1
        logging.info("Finished writing reports!!")

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__} from line 67")
        logging.error("Error while creating reports!")
        logging.error(f"{e}")
        print(e, "\n Error while creating reports!")
        sys.exit(1)


if __name__ == "__main__":
    logging.info("Execution Started...")
    print("Execution Started!!!")
    reports()
    logging.info("Successfully Executed!!!\n\n")
    print("Successfully Executed!!!")
    messagebox.showinfo(title="Status", message="Completed!!!")
