#!/usr/bin/env python
"""
Run the script Read_OCR for all files in a directory
Usage:
python batch_read_OCR.py source_dir_with_docx result_dir
"""
import sys
import os

import Read_OCR
import Csv_Excel

source_dir = str(sys.argv[1])
if len(sys.argv) >= 3:
    result_dir = str(sys.argv[2])

if not os.path.exists(result_dir):
    os.makedirs(result_dir)

for (dirname, dirs, files) in os.walk(source_dir):
    for filename in files:
        filename_main, file_extension = os.path.splitext(filename)
        print(filename_main)
        print(file_extension)
        if file_extension == '.docx':
            fullfilename = os.path.join(dirname,filename)
            csv_file_path = os.path.join(result_dir, filename_main + '.csv')
            export_file_path = os.path.join(result_dir, filename_main + '.xls')
            log_file_path = os.path.join(result_dir, filename_main + '_stats.txt')
            Read_OCR.readocr(fullfilename, csv_file_path, log_file_path)
            Csv_Excel.csv_to_xls(csv_file_path, export_file_path)
