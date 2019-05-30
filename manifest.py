#!/usr/bin/env python
# -*- coding: utf-8 -*-

# manifest.py
# author: Da Zhu
# E-mail: gonglisuozcd@gmail.com
"""
Manifest Generator

Usage:
    python manifest.py [options]
Options:
    -h --help        Display this help message
Example:
    python manifest.py
"""
import os
import sys
import getopt
import time
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, colors
except ImportError as e:
    print("ImportError: %s" % e)
    print("\nThis script depends on some third-party packages to function properly!")
    print("Requirements:")
    print("OpenPyXL    ----    A Python library to read/write Excel files")
    print("Installing openpyxl using pip as below:\npip install openpyxl")
    sys.exit(-1)


def generate_manifest():
    workbook = Workbook()  # a blank Workbook
    worksheet = workbook.active  # a blank Worksheet
    worksheet.title = "superman"
    custom_font = Font(name="Times New Roman", size=12, italic=False, color=colors.BLACK, bold=False)
    custom_alignment = Alignment(horizontal="center", vertical="center")
    directories = list()
    for item in os.listdir(os.getcwd()):
        if not os.path.isfile(os.path.join(os.getcwd(), item)):
            if len(os.listdir(os.path.join(os.getcwd(), item))):
                directories.append(item)
    # order by epicenter distance
    directories.sort(key=lambda dir_name:int(dir_name[0:3] if str(dir_name)[2].isdigit() else dir_name[0:2]))
    rows = list()
    for directory in directories:
        subdirectories = os.listdir(os.path.join(os.getcwd(), directory))
        # order by sequential number
        subdirectories.sort(key=lambda rsn:int(rsn[0:2] if str(rsn)[1].isdigit() else rsn[0:1]))
        for subdirectory in subdirectories:
            file_dir = os.path.join(os.getcwd(), directory, subdirectory)
            file_name_list = os.listdir(file_dir)
            rows.append((file_dir, file_name_list))

    for idx in range(len(rows)):
        file_dir = rows[idx][0]
        worksheet.row_dimensions[1 + idx].height = 20
        worksheet.column_dimensions["A"].width = 70
        worksheet["%c%s" % (65, 1 + idx)] = file_dir
        worksheet["%c%s" % (65, 1 + idx)].font = custom_font
        worksheet["%c%s" % (65, 1 + idx)].alignment = Alignment(horizontal="left", vertical="center")

        for idy in range(len(rows[idx][1])):
            worksheet.row_dimensions[1 + idx].height = 20
            worksheet.column_dimensions["%c" % (66 + idy)].width = 30
            worksheet["%c%s" % (66 + idy, 1 + idx)] = rows[idx][1][idy]
            worksheet["%c%s" % (66 + idy, 1 + idx)].font = custom_font
            worksheet["%c%s" % (66 + idy, 1 + idx)].alignment = custom_alignment
    timestamp = time.strftime("%Y/%m/%d %H:%M:%S", time.localtime(time.time()))
    out_name = "Manifest-NR.xlsx"
    workbook.save(filename=out_name)
    print("*" * 90)
    print_one = "*  I have successfully finished my job! (%s)"  % timestamp
    print("%-88s *" % print_one)
    print_two = "*  The manifest file named '%s' is right in the current directory." % out_name
    print("%-88s *" % print_two)
    print("*" * 90)
    os.system("dir")


def usage():
    print(__doc__)


def main():
    args = sys.argv[1:]
    try:
        options, args = getopt.getopt(args, "h", ["help"])
    except:
        usage()
        sys.exit(-1)
    for option, value in options:
        if option in ("-h", "--help"):
            usage()
            sys.exit(1)
    if len(args) != 0:
        usage()
        sys.exit(-1)
    generate_manifest()


if __name__ == "__main__":
    # import pdb
    # pdb.set_trace()
    main()
