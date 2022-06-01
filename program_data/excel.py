# -*- coding: utf-8 -*-
#
#  Created by: Jonathan O'Connell | joconnell@usgs.gov
#
#  Source repo: https://code.usgs.gov/koocanusa/servo-asr-generator
#
#  Script for generating servo ASRs based on an ASR template.
#  If wb = xw.Book.caller() is enabled, the servo excel template macro can be used to run the script.
#  Alternativley a workbook can be defined manually and ran from python from within this script.
#  The README.md file has further instructions.


import pandas as pd
import xlwings as xw
from datetime import datetime
from pathlib import PurePath


def clean_comments(type_col, com_col):
    """This function will format the FA and RA asr_comments to send to the NWQL comments section in the ASR.

    Parameters
    ----------
    type_col : Pandas Series
        Pandas column containing sample type i.e. fa/ra
    com_col : Pandas Series
        Pandas column containing the asr_comments

    Returns
    -------
    String of formatted FA/RA comments
    """
    comments = zip(type_col.str.upper(), com_col)
    return str(dict(comments)).replace("{", "").replace("}", "").replace("'", "")


def main():
    site_dict = {
        "12301933": "Kootenai River bl Libby Dam nr Libby MT",
        "12301919": "Lake Koocanusa at forebay, nr Libby, MT",
        "12300110": "Lake Koocanusa at international boundary",
        "12301830": "Lake Koocanusa at Tenmile Cr nr Libby, MT",
    }

    # sheet_list = ["ASR1", "ASR2", "ASR3", "ASR4", "blank"]
    # local_user_dir = os.path.dirname(os.path.realpath(__file__))
    # root_dir = local_user_dir[:51]
    # sys.path.append(local_user_dir)
    #  Ensure no files in temp folder.
    # for file in glob.glob(f"{local_user_dir}\\temp\\*.docx"):
    #     os.remove(file)

    try:
        # This line will use the calling excel workbook as the source data.
        wb = xw.Book.caller()
    except Exception:
        print("Debugging triggered, generating ASRs with test excel file!")
        wb = xw.Book(".\\testing.xlsm")
    sheet = wb.sheets["servo"]

    # header_loc = com.find_header(sheet)
    station_info = (
        sheet.range("station_info")
        .options(pd.DataFrame, header=1, index=False, expand="table")
        .value
    )
    station_info = station_info.iloc[0]
    #  Collect and define variables, if no data available, make assumptions.
    stationID = str(int(station_info["stationID"]))
    station_name = site_dict[stationID]
    station_depth = station_info["depth (m)"]

    try:
        # Station depth imports as float, int makes 10.0 as 10 then make it a string.
        depth = str(int(station_depth))
    except (TypeError, ValueError):
        # If depth has a letter in it, remove it
        if station_depth is None:
            print("No depth found, assume river site.")
            depth = "-699999999999999999999999999999999999999999999"
        else:
            depth = str(
                int(
                    float(
                        "".join(ch for ch in str(station_depth) if ch.isdigit() or ch == ".")
                    )
                )
            )
    try:
        ship_date = str(station_info["ship_date"].strftime("%m/%d/%Y"))
    except AttributeError:
        ship_date = datetime.now().strftime("%m/%d/%Y")
    try:
        sCond = str(int(station_info["SC"]))
    except ValueError:
        sCond = "250e"
    try:
        if station_info["blank_datetime"] is not None:
            blank_datetime = pd.to_datetime(station_info["blank_datetime"]).strftime(
                "%Y%m%d %H%M"
            )
        else:
            blank_datetime = None
    except AttributeError:
        print("Invalid blank datetime, skipping Blank ASR")
        blank_datetime = None

    #  Read in sample data to Pandas, rename headers, drop first row if NaN, keep only first 8 rows.
    # data = xw.load(index=0, header=13)
    # data = pd.read_excel(wb.fullname, header=header_loc, usecols="A:J")
    data = sheet.range("data_range").options(pd.DataFrame, header=1, index=False).value
    new_header = [
        "SAM-ID",
        "Date-Time",
        "Temp",
        "Li-Batt",
        "Pic-Batt",
        "volume",
        "comment",
        "asr_comment",
        "invalid",
        "type",
    ]
    data.rename(columns=dict(zip(data.columns, new_header)), inplace=True)
    if pd.isna(data["SAM-ID"][0]):
        data.drop([0], inplace=True)
    # data = data.iloc[:8]  # type: ignore

    #  Set datetime for RA samples based on FA time.
    data.loc[data["Date-Time"].isna(), "Date-Time"] = data.iloc[::2]["Date-Time"].to_numpy(
        copy=True
    )

    #  Drop any sample that does not have a volume or a type (FA/RA) or marked invalid with an "x".
    # data = data.dropna(subset=["volume"])
    data = data.dropna(subset=["type"])
    data = data[data["invalid"] != "x"]
    date_groups = data.groupby("Date-Time")

    asr_count = 0
    #  Group data by date, and generate an ASR for that date.  This usually means 1 FA and 1 RA sample per ASR.
    for date, vals in date_groups:
        asr_count += 1
        asr_sheet = wb.sheets(f"ASR{asr_count}")

        try:
            sample_datetime = pd.to_datetime(date)
            sample_datetime_fmt = sample_datetime.strftime("%Y%m%d %H%M")
        except AttributeError:
            print(f"Invalid sample datetime {date}.  Could not be parsed by pd.to_datetime()")
            continue
        # End date is 7 days later, inlcuding the first day, so add 6 to start date.
        sample_datetime_end = sample_datetime + pd.Timedelta(days=6)
        sample_datetime_end_fmt = sample_datetime_end.strftime("%Y%m%d %H%M")

        asr_comments = vals.dropna(subset=["asr_comment"])
        if asr_comments.empty is True:
            cleaned_comment = ""
        else:
            cleaned_comment = clean_comments(asr_comments["type"], asr_comments["asr_comment"])
        #  Samples are grouped by date-time: count and record number of associated FA and RA samples onto ASR.
        try:
            fa_count = date_groups.get_group(date)["type"].str.lower().value_counts()["fa"]
            if fa_count > 0:
                asr_sheet.range("H41").value = fa_count
                asr_sheet.range("D36").value = "3132"
        except KeyError:
            print(f"No FA sample found on {sample_datetime_fmt}")
        try:
            ra_count = date_groups.get_group(date)["type"].str.lower().value_counts()["ra"]
            if ra_count > 0:
                asr_sheet.range("S42").value = ra_count
                if asr_sheet.range("D36").value is None:
                    asr_sheet.range("D36").value = "3306"
                else:
                    asr_sheet.range("L36").value = "3306"
        except KeyError:
            print(f"No RA sample found on {sample_datetime_fmt}")

        asr_sheet.range("Y12:AF12").value = list(sample_datetime_fmt)
        asr_sheet.range("Y14:AF14").value = list(sample_datetime_end_fmt)
        asr_sheet.range("P12:W12").value = list(stationID)

        asr_sheet.range("I28").value = station_name
        asr_sheet.range("AJ52").value = ship_date
        asr_sheet.range("H56").value = sCond

        #  No depth associated with the river site.
        if stationID == "12301933":
            asr_sheet.range("I29").value = cleaned_comment
        else:
            asr_sheet.range("I29").value = f"{depth} m depth   {cleaned_comment}  "

        print(f"ASR for {sample_datetime_fmt} complete.")

    if blank_datetime is not None:
        # Blank servo ASR if applicable
        wb.sheets("blank").range("H41").value = "1"
        wb.sheets("blank").range("D36").value = "3132"
        wb.sheets("blank").range("Y12:AF12").value = list(blank_datetime)
        wb.sheets("blank").range("P12:W12").value = list(stationID)

        wb.sheets("blank").range("I28").value = station_name
        wb.sheets("blank").range("AJ52").value = ship_date
        wb.sheets("blank").range("H56").value = sCond

        #  No depth associated with the river site.
        if stationID == "12301933":
            wb.sheets("blank").range("I29").value = "ServoSipper Blank"
        else:
            wb.sheets("blank").range("I29").value = f"{depth} m depth   ServoSipper Blank"
        print(f"ASR for blank on {blank_datetime} complete.")

    base_file_name = pd.to_datetime(data.iloc[0]["Date-Time"]).strftime("%Y%m%d_%H%M")
    if stationID == "12301933":
        new_file_name = base_file_name
    else:
        new_file_name = f"{base_file_name}_{depth}m"

    asr_list = [f"ASR{num}" for num in range(1, asr_count + 1)]
    if blank_datetime is not None:
        asr_list += ["blank"]

    parent_path = PurePath(wb.fullname).parent
    asr_save_path = PurePath(
        parent_path, "ASRs", f"{stationID}_{new_file_name}_NWQL_ASR_Servo.pdf"
    )
    wb.to_pdf(asr_save_path, include=asr_list)

    # wb.save()
    print("Done!")


main()
