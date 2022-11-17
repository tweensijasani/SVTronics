import re
import sys
import logging
import tkinter as tk
from tkinter import filedialog, messagebox


logging.basicConfig(level=logging.DEBUG, filename="Gencad_logfile.txt", filemode="a+",
                    format="%(asctime)-15s %(levelname)-8s %(message)s")

root = tk.Tk()
root.withdraw()


try:
    counter = 0
    file = filedialog.askopenfilename(title="Select Gencad File")
    while file == '' and counter < 1:
        messagebox.showerror(title="File Error", message="Gencad File Not Selected")
        logging.warning("Gencad File Not Selected")
        file = filedialog.askopenfilename(title="Select Gencad File")
        counter += 1
    if file == '':
        messagebox.showerror(title="Invalid Input", message="Something went wrong!! Please try again....")
        logging.error("Terminated: Gencad File not selected!")
        sys.exit(1)
    else:
        logging.info(f"{file} Selected")

except Exception as e:
    messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
    logging.error(f"{e.__class__}")
    logging.error("Error while fetching file!")
    logging.error(f"{e}")
    print(e, "\n Error while fetching file!")
    sys.exit(1)

try:
    logging.info("Reading .cad file...")
    gencad = open(file, 'r')
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
    for line in range(start+1, stop):
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
    logging.info("Finished reading!")

    logging.info("Writing modified .cad file...")
    var = gencad.name.split("/")
    new_file_name = var.pop().replace(".cad", "_modified.cad")
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
    logging.error("Error while editing .cad file!")
    logging.error(f"{e}")
    print(e, "\n Error while editing .cad file!")
    sys.exit(1)
