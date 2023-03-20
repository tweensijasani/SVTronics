import sys
import csv
import random
import logging
import tkinter as tk
from easygui import *
from tkinter import filedialog, messagebox

logging.basicConfig(level=logging.DEBUG, filename="random_select_logs.txt", filemode="a+",
                    format="%(asctime)-15s %(levelname)-8s %(message)s")

root = tk.Tk()
root.withdraw()


def getfile():
    try:
        counter = 0
        csv_file = filedialog.askopenfilename(title="Select CSV File", filetypes=(("CSV Files", "*.csv"),))
        while csv_file == '' and counter < 1:
            messagebox.showerror(title="File Error", message="CSV File Not Selected!")
            logging.warning("CSV File Not Selected!")
            csv_file = filedialog.askopenfilename(title="Select CSV File")
            counter += 1
        if csv_file == '':
            messagebox.showerror(title="Invalid Input", message="Something went wrong!! Please try again....")
            logging.error("Terminated: CSV file not selected!")
            print("Terminated: CSV file not selected!")
            sys.exit(1)
        else:
            logging.info("CSV File Selected!")
        return csv_file

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__}")
        logging.error("Error while fetching file!")
        logging.error(f"{e}")
        print(e, "\n Error while fetching file!")
        sys.exit(1)


def select(csv_file):
    try:
        logging.info("Reading CSV file...")
        item_list = []
        with open(csv_file) as csv_file:
            csv_reader = csv.reader(csv_file)
            for row in csv_reader:
                item_list.append(row[0].strip())
        logging.info("Finished reading!!")

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__}")
        logging.error("Error while reading file!")
        logging.error(f"{e}")
        print(e, "\n Error while reading file!")
        sys.exit(1)

    try:
        logging.info("Getting Count Info...")
        text = "Random Select"
        title = "Enter Count"
        input_list = ["Count"]
        output = multenterbox(text, title, input_list)

        while output[0] is None or not output[0].isnumeric():
            messagebox.showerror(title="Invalid Input", message="Please enter only Integer Value")
            text = "Random Select"
            title = "Enter Count"
            input_list = ["Count"]
            output = multenterbox(text, title, input_list)

        count = int(output[0].strip())
        logging.info("Count info populated")

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__}")
        logging.error("Error while getting Count Input!")
        logging.error(f"{e}")
        print(e, "\n Error while getting Count Input!")
        sys.exit(1)

    try:
        logging.info("Selecting random items...")
        if count > len(item_list):
            messagebox.showerror(title=f"IndexOutOfBound", message="Count is greater than the list of items!!")
            sys.exit(1)
        else:
            output_list = random.sample(item_list, count)
            messagebox.showinfo(title="Selected Items", message=f"{output_list}")
            with open("Selected_List.txt", "w") as f:
                for item in output_list:
                    f.write("%s\n" % item)
            f.close()

    except Exception as e:
        messagebox.showerror(title=f"{e.__class__}", message="Something went wrong!! Please try again....")
        logging.error(f"{e.__class__} from line 67")
        logging.error("Error while selecting random items!")
        logging.error(f"{e}")
        print(e, "\n Error while selecting random items!")
        sys.exit(1)


if __name__ == "__main__":
    logging.info("Execution Started...")
    print("Execution Started!!!")
    csv_file = getfile()
    select(csv_file)
    logging.info("Successfully Executed!!!\n\n")
    print("Successfully Executed!!!")
    messagebox.showinfo(title="Status", message="Completed!!!")