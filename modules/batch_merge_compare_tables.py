#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Thu Sep 20 2018

To merge words in the 4th column of file like 131-140_Minh.xls to the 4th column of file like 131-140_Minh_Giao.xls (this 4th column should match the 5th column of the former file)

@author: dang

Written for Python 2.7
Test command:
python batch_merge_compare_tables.py dir_Giao_xls dir_before_xls result_dir
"""
import sys
import os
import csv
import xlrd
import re
import Csv_Excel
import difflib
import numpy

offset = 6

def read_worksheet_to_table(work_sheet):
    table = []
    for rownum in range(work_sheet.nrows):
        table.append(work_sheet.row_values(rownum))
    return table
        
def compare_merge_xls(xlsFile_Giao, xlsFile_before, csvFile, new_xlsFile, reportFile):
    work_book_Giao = xlrd.open_workbook(xlsFile_Giao)
    work_sheet_Giao = work_book_Giao.sheet_by_index(0)
    work_book_before = xlrd.open_workbook(xlsFile_before)
    work_sheet_before = work_book_before.sheet_by_index(0)
    
    table_Giao = read_worksheet_to_table(work_sheet_Giao)
    table_before = read_worksheet_to_table(work_sheet_before)
    
    count_success = 0
    count_failure = 0
    
    report = xlsFile_Giao+':\n'
    
    for rownum in range(len(table_before)):
        if table_before[rownum][4] != '':
            report += '  Merging text at row '+str(rownum)+': '+table_before[rownum][4]+'...\n'
            rows_to_compare = range(max(0,rownum-offset), min(rownum+offset,len(table_Giao)))
            similarity = []
            for kk in rows_to_compare:
                similarity += [difflib.SequenceMatcher(None, table_before[rownum][5], table_Giao[kk][4]).ratio()]
            if max(similarity)>0.9:
                matching_row = rows_to_compare[numpy.argmax(similarity)]
                table_Giao[matching_row][4] = table_before[rownum][4] + table_Giao[matching_row][4]
                report += '    Successful! Best match='+str(max(similarity))+' at row_Giao: '+str(matching_row)+'\n'
                count_success += 1
            else:
                report += '    Failed! Highest match='+str(max(similarity))+'\n'
                count_failure += 1
    report += '  Stats:\n'
    report += '    Success: '+str(count_success)+'\n'
    report += '    Failure: '+str(count_failure)+'\n'
    print('  Stats:')
    print('    Success: '+str(count_success)+'')
    print('    Failure: '+str(count_failure)+'')
    
    with open(csvFile, 'wb') as file_in:
        text_writer = csv.writer(file_in, delimiter='\t', quoting=csv.QUOTE_ALL)
        for rownum in range(len(table_Giao)):
            text_writer.writerow(
                list(x.encode('utf-8') if type(x) == type(u'') else x
                    for x in table_Giao[rownum]))
    Csv_Excel.csv_to_xls(csvFile, new_xlsFile)
    with open(reportFile, 'wt') as text_file:
        text_file.write(report.encode('utf-8'))

if len(sys.argv) < 4:
    raise ValueError('This command need 3 directories: python batch_merge_compare_tables.py dir_Giao_xls dir_before_xls result_dir')

dir_Giao_xls = str(sys.argv[1])
dir_before_xls = str(sys.argv[2])
result_dir = str(sys.argv[3])

if not os.path.exists(result_dir):
    os.makedirs(result_dir)

for (dirname, dirs, files) in os.walk(dir_Giao_xls):
    for filename in files:
        filename_main, file_extension = os.path.splitext(filename)
        print(filename_main+file_extension)
        if file_extension == '.xls':
            xlsFile_Giao = os.path.join(dirname, filename)
            xlsFile_before = os.path.join(dir_before_xls, re.sub(r'_Giao','',filename))
            csvFile = os.path.join(result_dir, filename_main + '.csv')
            new_xlsFile = os.path.join(result_dir, filename_main + '.xls')
            log_file_path = os.path.join(result_dir, filename_main + '_post_process_stats.txt')
            compare_merge_xls(xlsFile_Giao, xlsFile_before, csvFile, new_xlsFile, log_file_path)

