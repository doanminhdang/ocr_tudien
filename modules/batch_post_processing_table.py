#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Thu Aug 30 2018

@author: dang

Written for Python 2.7
Test command:
python post_processing_table.py xlsFile csvFile new_xlsFile
"""
import sys
import os
import csv
import xlrd
import re
import Csv_Excel

text_to_move = [u'(các)',
                u'(thuộc)',
                u'(tính)',
                u'(chứng)',
                u'(có)',
                u'(sự)',
                u'(được)',
                u'(kỹ thuật)',
                u'(hiện tượng)',
                u'(trạng thái)',
                u'(nhóm)',
                u'(bị)',
                u'(tác nhân)',
                u'(quá trình)',
                u'(phương pháp)',
                u'(cơn)',
                u'(vật)',
                u'(một)',
                u'(thuyết)',
                u'(chất)',
                u'(độ)',
                u'(quang)',
                u'(cái)',
                u'(than)',
                u'(phụ)',
                u'(tờ)',
                u'(loại)',
                u'(mạch)',
                u'(phụ)'
                ]

def post_process_xls(xlsFile, csvFile, new_xlsFile, reportFile):
    work_book = xlrd.open_workbook(xlsFile)
    work_sheet = work_book.sheet_by_index(0)
    report = ''
    with open(csvFile, 'wb') as file_in:
        text_writer = csv.writer(file_in, delimiter='\t', quoting=csv.QUOTE_ALL)
        for rownum in range(work_sheet.nrows):
            text = work_sheet.row_values(rownum)
            for word in text_to_move:
                if re.search(re.escape(word)+r'$', text[3]):
                    text[3] = re.sub(re.escape(word)+r'$', '', text[3]).strip()
                    text[4] = word+u' '+text[4]
            text_writer.writerow(
                list(x.encode('utf-8') if type(x) == type(u'') else x
                    for x in text))
            if rownum>0 and rownum<work_sheet.nrows-2:
                this_word = work_sheet.row_values(rownum)[3]
                prev_word = work_sheet.row_values(rownum-1)[3]
                next_word = work_sheet.row_values(rownum+1)[3]
                if len(this_word)>0 and len(prev_word)>0 and len(next_word)>0:
                    if this_word[0].upper()!=prev_word[0].upper() and prev_word[0].upper()==next_word[0].upper():
                        report += 'Check 1st letter in row ' + str(rownum) + '\n'
    Csv_Excel.csv_to_xls(csvFile, new_xlsFile)
    with open(reportFile, 'wt') as text_file:
        text_file.write(report)

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
        if file_extension == '.xls':
            xls_file = os.path.join(dirname,filename)
            csv_file = os.path.join(result_dir, filename_main + '.csv')
            new_xls_file = os.path.join(result_dir, filename_main + '.xls')
            log_file_path = os.path.join(result_dir, filename_main + '_post_process_stats.txt')
            post_process_xls(xls_file, csv_file, new_xls_file, log_file_path)

