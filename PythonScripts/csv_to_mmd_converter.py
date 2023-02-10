import csv
import logging


logging.basicConfig(level=logging.DEBUG, filename="csv_logfile.txt", filemode="a+",
                    format="%(asctime)-15s %(levelname)-8s %(message)s")


def convert(csv_file):

    try:
        logging.info("Creating mmd file...")
        var = csv_file.split("/")
        last_item = var.pop()
        if ".csv" in last_item:
            new_file_name = last_item.replace(".csv", ".mmd")
        else:
            new_file_name = last_item.replace(".CSV", ".MMD")
        var.append(new_file_name)
        new_file = "/".join(var)

        mmd_data = []
        fid_data = ["[Fiducial]\n", "Fid1_X=0\n", "Fid1_Y=0\n", "Fid2_X=0\n", "Fid2_Y=0\n", "[Part Info]\n", "Coordinate Transform=NO\n"]
        mmd_data.extend(fid_data)
        line_count = 0
        csv_data = []
        switcher = {1: "0000000", 2: "000000", 3: "00000", 4: "0000", 5: "000", 6: "00", 7: "0", 8: ""}

        logging.info("Creating mmd data...")
        with open(csv_file) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            for row in csv_reader:
                if line_count == 0:
                    line_count += 1
                else:
                    csv_data.append(f"#{switcher[len(str(line_count))]}{line_count}={row[7]}\t{row[8]}\t{row[9]}\t{row[1]}\t{row[2]}\n")
                    line_count += 1

        mmd_data.append(f"Part Count={line_count-1}\n")
        mmd_data.extend(csv_data)

    except Exception as e:
        logging.error(f"{e.__class__}")
        logging.error("Error while creating mmd file data!")
        logging.error(f"{e}")
        print(e, "\n Error while creating mmd file data!")
        return False, False

    try:
        logging.info("Writing mmd file...")
        with open(new_file, "w") as f:
            for item in mmd_data:
                f.write("%s" % item)
        f.close()

        logging.info("Finished writing!")
        return True, new_file_name

    except Exception as e:
        logging.error(f"{e.__class__}")
        logging.error("Error while writing mmd file!")
        logging.error(f"{e}")
        print(e, "\n Error while writing mmd file!")
        return False, False

