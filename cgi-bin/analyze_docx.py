#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Tue Apr 24 13:15:35 2018

@author: dang

Written for Python 2.7
Test command:
python readocr.py
"""
import cgi, cgitb
import os, sys
import glob
import shutil

# Web deploy

cgitb.enable()

cgi_bin_dir = os.getcwd()

server_name = os.environ["SERVER_NAME"]
#BASE_URL = 'http://localhost:8000'.rstrip('/')
BASE_URL = 'http://'+server_name+'/ocr_tudien'.rstrip('/')  # for web
TOP_DIR = '../ocr_tudien'.rstrip('/')
MODULE_DIR = 'modules'.rstrip('/')
DATA_DIR = 'data'
UPLOAD_DIR = 'uploads'.rstrip('/')
TEMP_RESULT_DIR = 'results'.rstrip('/')
OUTPUT_DIR = 'downloads'.rstrip('/')
ABS_TOP_DIR = cgi_bin_dir.replace('/cgi-bin', '/ocr_tudien').rstrip('/')
#OUTPUT_FILE = 'results.zip'

# Note: running on localhost:8000, the TOP_DIR is '.', because the current directory
# is with the HTML file that calls to this Python script (upper dir)
# While running on web server of Hawkhost, TOP_DIR is '..'

from sys import path
path.append(TOP_DIR+'/'+MODULE_DIR)

#import Textparser
#import Pdfsplitter
import Read_OCR

def create_user_dir(base_dir):
    """Create directory for the request, with the structure
    base_dir/YYYY/MM/DD/hour-minute-second, and output the relative path
    """
    from datetime import datetime
    # Create directory

    def generate_parent_dir():
        today = datetime.utcnow()

        relative_path_dir_year = str(today.year)
        relative_path_dir_month = os.path.join(relative_path_dir_year, str(today.month))
        relative_path_dir_day = os.path.join(relative_path_dir_month, str(today.day))
        relative_path_dir_now = os.path.join(relative_path_dir_day, str(today.hour)+'-'+str(today.minute)+'-'+str(today.second))

        path_dir_year = os.path.join(base_dir, relative_path_dir_year)
        path_dir_month = os.path.join(base_dir, relative_path_dir_month)
        path_dir_day = os.path.join(base_dir, relative_path_dir_day)

        if not os.path.exists(path_dir_year):
            os.makedirs(path_dir_year)

        if not os.path.exists(path_dir_month):
            os.makedirs(path_dir_month)

        if not os.path.exists(path_dir_day):
            os.makedirs(path_dir_day)

        return relative_path_dir_now


    relative_path_dir_time = generate_parent_dir()

    while os.path.exists(os.path.join(base_data_dir, relative_path_dir_time)):
        relative_path_dir_time = generate_parent_dir()

    os.makedirs(os.path.join(base_data_dir, relative_path_dir_time))

    return relative_path_dir_time


base_data_dir = os.path.join(TOP_DIR, DATA_DIR)
relative_user_dir = create_user_dir(base_data_dir)

os.makedirs(os.path.join(base_data_dir, relative_user_dir, UPLOAD_DIR))
os.makedirs(os.path.join(base_data_dir, relative_user_dir, TEMP_RESULT_DIR))
os.makedirs(os.path.join(base_data_dir, relative_user_dir, OUTPUT_DIR))

# Create instance of FieldStorage, it can only be initiated once per request
form = cgi.FieldStorage()

# Get data from fields
delete_data = form.getvalue('delete_data')
if not delete_data:  # set default value
    delete_data = 'yes'

def save_uploaded_file(cgi_form, form_field, upload_dir, whitelist_ext):
    """This saves a file uploaded by an HTML form.
       The form_field is the name of the file input field from the form.
       For example, the following form_field would be "file_1":
           <input name="file_1" type="file">
       The upload_dir is the directory where the file will be written.
       The whitelist_ext is the set of allowed file extensions for uploading.
       If no file was uploaded or if the field does not exist then
       this does nothing.
    """
    if not cgi_form.has_key(form_field): return False
    file_item = cgi_form[form_field]
    if not file_item.file: return False
    # Strip leading path from file name to avoid
    # directory traversal attacks.
    # Replace \ by / to make sure compatibility with Windows path
    filename_base = os.path.basename(file_item.filename.replace("\\", "/"))
    mainname, extname = os.path.splitext(filename_base)  # mainname is '123.php.', extname is '.jpg'
    # Use white list of file type to be uploaded
    if not extname in whitelist_ext: return False
    # Replace . by _ to protect against double extension attacks which can activate PHP scripts
    filename_base = mainname.replace('.', '_') + extname
    file_path = os.path.join(upload_dir, filename_base)
    with open(file_path, 'wb') as outfile:
        shutil.copyfileobj(file_item.file, outfile)
        # outfile.write(file_item.file.read())
        # while 1:
        #     chunk = file_item.file.read(100000)
        #     if not chunk: break
        #     outfile.write (chunk)
    return filename_base


text_ext = set(['.docx', '.DOCX'])

white_list = set()
white_list = white_list.union(text_ext)

file_input_1 = save_uploaded_file(form, "upload1", os.path.join(TOP_DIR, DATA_DIR, relative_user_dir, UPLOAD_DIR), white_list)

if file_input_1:
    file_1_path = os.path.join(ABS_TOP_DIR, DATA_DIR, relative_user_dir, UPLOAD_DIR, file_input_1)
    url_file_1 = BASE_URL+'/'+DATA_DIR+'/'+relative_user_dir+'/'+UPLOAD_DIR+'/'+file_input_1
    message_file_1 = 'File: '+file_input_1+' was uploaded successfully.'
else:
    message_file_1 = 'Input file is not an accepted file. It was not uploaded.'

proceed_flag = file_input_1# and os.path.splitext(file_input_1)[1]==os.path.splitext(file_input_2)[1]

if proceed_flag:
    # Processing text files DOCX
    export_file_path = os.path.join(ABS_TOP_DIR, DATA_DIR, relative_user_dir, OUTPUT_DIR, file_input_1 + '.csv')
    url_file_export = BASE_URL+'/'+DATA_DIR+'/'+relative_user_dir+'/'+OUTPUT_DIR+'/'+file_input_1 + '.csv'
    log_file_path = os.path.join(ABS_TOP_DIR, DATA_DIR, relative_user_dir, OUTPUT_DIR, file_input_1 + '_stat.txt')
    url_file_log = BASE_URL+'/'+DATA_DIR+'/'+relative_user_dir+'/'+OUTPUT_DIR+'/'+file_input_1 + '_stat.txt'
    
    Read_OCR.readocr(file_1_path, export_file_path, log_file_path)

    # Result file to be given back
    # os.system('rm ../data/results.zip')
    # os.system('rm '+TOP_DIR+'/'+DATA_DIR+'/'+relative_user_dir+'/'+OUTPUT_DIR+'/'+OUTPUT_FILE)
    # os.system('zip -r ../data/results.zip ../data/results')
    #os.system('zip -r -j '+TOP_DIR+'/'+DATA_DIR+'/'+relative_user_dir+'/'+OUTPUT_DIR+'/'+OUTPUT_FILE+' '+TOP_DIR+'/'+DATA_DIR+'/'+relative_user_dir+'/'+TEMP_RESULT_DIR)

    #url_file_result = BASE_URL+'/'+DATA_DIR+'/'+relative_user_dir+'/'+OUTPUT_DIR+'/'+OUTPUT_FILE


# Clean up
if not delete_data != 'yes':
    ## os.system('rm -r ../data/results/*')
    #os.system('rm -r '+TOP_DIR+'/'+DATA_DIR+'/'+relative_user_dir+'/'+TEMP_RESULT_DIR+'/*')
    # os.system('rm -r ../data/uploads/*')
    os.system('rm -r '+TOP_DIR+'/'+DATA_DIR+'/'+relative_user_dir+'/'+UPLOAD_DIR+'/*')

# Output to the web
print "Content-type:text/html"
print ""
print "<html>"
print "<head>"
print "<title>Kết quả xử lý file văn bản</title>"
print "<meta charset=\"utf-8\"/>"
print "</head>"
print "<body>"
print "<h1>Kết quả xử lý file văn bản</h1>"
print "<h2>%s</h2>" % (message_file_1)
if delete_data != 'yes' and file_input_1:
    print "<h2>Link của file đã tải lên: <a href=\"%s\">%s</a></h2>" % (url_file_1, url_file_1)
if proceed_flag:
    print "<h2>Link file kết quả thu về: <a href=\"%s\">%s</a></h2>" % (url_file_export, url_file_export)
    print "<h2>Link file thống kê: <a href=\"%s\">%s</a></h2>" % (url_file_log, url_file_log)
if not delete_data != 'yes':
    print "<h2>Source data was deleted on the server.</h2>"
print "</body>"
print "</html>"
