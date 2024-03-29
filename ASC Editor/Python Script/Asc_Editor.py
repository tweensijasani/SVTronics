import re
import sys
import logging
import tkinter as tk
from tkinter import filedialog, messagebox


logging.basicConfig(level=logging.DEBUG, filename="asc_logfile.txt", filemode="a+",
                    format="%(asctime)-15s %(levelname)-8s %(message)s")

root = tk.Tk()
root.withdraw()


def get_file():
    try:
        counter = 0
        file = filedialog.askopenfilename(title="Select ASC File")
        while file == '' and counter < 1:
            messagebox.showerror(title="File Error", message="ASC File Not Selected")
            logging.warning("ASC File Not Selected")
            file = filedialog.askopenfilename(title="Select ASC File")
            counter += 1
        if file == '':
            messagebox.showerror(title="Invalid Input", message="Something went wrong!! Please try again....")
            logging.error("Terminated: ASC File not selected!")
            sys.exit(1)
        else:
            logging.info(f"{file} Selected")
        return file

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__}")
        logging.error("Error while fetching file!")
        logging.error(f"{e}")
        print(e, "\n Error while fetching file!")
        sys.exit(1)


def write_file(file):
    try:
        logging.info("Reading asc file...")
        ascfile = open(file, 'r')
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
        for line in range(start+1, stop):
            item = file_data[line].strip().split(" ")
            value = item[0]
            if "-" in value:
                value = value.replace("-", "_")
                item[0] = value
            item = " ".join(item)
            file_data[line] = item + "\n"
            change_count += 1

        ascfile.close()
        logging.info("Finished reading!")

        logging.info("Writing modified ASC file...")
        var = ascfile.name.split("/")
        temp = var.pop()
        new_file_name = temp.replace(".ASC", "_modified.ASC")
        new_file_name = new_file_name.replace(".asc", "_modified.asc")
        var.append(new_file_name)
        new_file = "/".join(var)

        with open(new_file, "w") as f:
            for item in file_data:
                f.write("%s" % item)
        f.close()
        logging.info("Finished writing!")

        messagebox.showinfo(title="Completed", message=f"Dash replaced to Underscore at {change_count} places!")
        logging.info(f"Dash replaced to Underscore at {change_count} places!")

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__}")
        logging.error("Error while editing .ASC file!")
        logging.error(f"{e}")
        print(e, "\n Error while editing .ASC file!")
        sys.exit(1)


if __name__ == "__main__":
    logging.info("Execution Started...")
    print("Execution Started!!!")

    file = get_file()
    write_file(file)

    logging.info("Successfully Executed!!!\n\n")
    print("Successfully Executed!!!")
    messagebox.showinfo(title="Status", message="Completed!!!")