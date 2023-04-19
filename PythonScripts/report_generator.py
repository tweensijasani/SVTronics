import logging
from datetime import date, datetime, timedelta


logging.basicConfig(level=logging.DEBUG, filename="reports_logfile.txt", filemode="a+",
                    format="%(asctime)-15s %(levelname)-8s %(message)s")


def batchwise(serialno, copies, time, modelid, boardsn, rdate, rtime, time_int, mounting, solder, ic_count, total):
    try:
        logging.info("Creating report files...")
        time = time.split(":")
        nhours = int(time[0])
        nminutes = int(time[1])
        nseconds = int(time[2])
        time_int = time_int.split(":")
        dhours = int(time_int[0])
        dminutes = int(time_int[1])
        dseconds = int(time_int[2])
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
            if count != 0 and count % 10 == 0:
                ttime = datetime.strptime(ttime, "%H%M%S") + timedelta(hours=nhours, minutes=nminutes, seconds=nseconds)
                ttime = ttime.strftime("%H%M%S")
            filename = f"upload_folder/{modelid}/SimpleCoverageExport_{x}_{tdate}_{ttime}.txt"
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
        return True, None

    except Exception as e:
        logging.error(f"{e.__class__}")
        logging.error("Error while creating reports!")
        logging.error(f"{e}")
        print(e, "\n Error while creating reports!")
        return False, f"Error while creating reports!\n{e.__class__}\n{e}"


def single_batch(serialno, copies, modelid, boardsn, rdate, rtime, time_int, mounting, solder, ic_count, total):
    try:
        logging.info("Creating report files...")
        time_int = time_int.split(":")
        dhours = int(time_int[0])
        dminutes = int(time_int[1])
        dseconds = int(time_int[2])
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
            filename = f"upload_folder/{modelid}/SimpleCoverageExport_{x}_{tdate}_{ttime}.txt"
            temp = f"{boardsn}"
            while len(temp) < 9:
                temp = f"0{temp}"
            data_list[1][1] = f"{temp}"
            data_list[1][4] = rtime
            data = data_list[0] + data_list[1] + data_list[2] + data_list[3] + data_list[4] + data_list[5] + data_list[
                6] + data_list[7]
            with open(filename, "w") as f:
                f.write("".join(data))
            f.close()
            rtime = datetime.strptime(rtime, "%H:%M:%S") + timedelta(hours=dhours, minutes=dminutes, seconds=dseconds)
            rtime = rtime.strftime("%H:%M:%S")
            boardsn += 1
            count += 1
        logging.info("Finished writing reports!!")
        return True, None

    except Exception as e:
        logging.error(f"{e.__class__}")
        logging.error("Error while creating reports!")
        logging.error(f"{e}")
        print(e, "\n Error while creating reports!")
        return False, f"Error while creating reports!\n{e.__class__}\n{e}"


if __name__ == "__main__":
    logging.info("Execution Started...")
    print("Execution Started!!!")
    logging.info("Successfully Executed!!!\n\n")
    print("Successfully Executed!!!")
