# -*- coding: utf-8 -*-

import os
import subprocess
import pandas as pd
import time
from werkzeug.datastructures import FileStorage
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from compare_excel_files import compare, compare_sheets

UPLOAD_FOLDER = 'C:/Users/lakhotem/Temp/Python/'
ALLOWED_EXTENSIONS = set(['xls', 'xlsx'])

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

file1_name = ""
file2_name = ""
rslt_df_for_web = pd.DataFrame(None, dtype=str)
output_excel_file = pd.ExcelWriter(str(str(os.getenv('HOMEPATH')).replace("\\","/") + "/Downloads/output_"+str(time.time()).split('.')[0]+".xlsx"))
dict_df = {}

def allowed_files(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/upload", methods = ['POST'])
@app.route("/excelcompare", methods = ['GET'])
def get_files():

    if (request.method == 'GET'):
        return render_template("input.html", heading = "Compare two excel files", tables = dict_df)

    if(request.method == 'POST'):

        if (('file1' not in request.files) | ('file2' not in request.files)):
          return render_template("input.html", heading="One or more files not uploaded")

        # Get the uploaded filenames from the html
        uploaded_file1 = request.files["file1"]
        uploaded_file2 = request.files['file2']
        file1_name = uploaded_file1.filename
        file2_name = uploaded_file2.filename

        # If no file or only one file uploaded, stop the processing
        if(((file1_name == '') & (file2_name == '')) | (file1_name == '') | (file2_name == '')):
            returned_file_names = "One or more files not uploaded"
            return render_template("input.html", heading="Comparison Results are", file_names=returned_file_names)
        elif ((allowed_files(file1_name)) | (allowed_files(file2_name))):

            # If both the files are uploaded, save them in a local directory for future use (Check if other options are available for processing file without saaving them)
            uploaded_file1.save(os.path.join(app.config['UPLOAD_FOLDER'], file1_name))
            uploaded_file2.save(os.path.join(app.config['UPLOAD_FOLDER'], file2_name))

            #file1_temp = None
            #with open(uploaded_file1, 'r') as fp:
                #file = FileStorage(fp)
            #file_in_cloud = uploaded_file1.read()
            #file_in_cloud = file_in_cloud.decode('utf-8')
            #file_in_cloud.save()

            # Just a simple string for printing on the web page
            returned_file_names = str("Comparing " + file1_name + " and " + file2_name)
            # Get the output directory for saving the generated excel file after comparison and merging
            output_path = str(str(os.getenv('HOMEPATH')).replace("\\","/") + "/Downloads/")
            #result_str = ""#compare(str(UPLOAD_FOLDER+file1_name),str(UPLOAD_FOLDER+file2_name),output_path)

            # Get comparison result for each sheet
            file1 = pd.read_excel(str(UPLOAD_FOLDER+file1_name), None)
            #file1 = pd.read_excel(file_in_cloud, None)
            file2 = pd.read_excel(str(UPLOAD_FOLDER+file2_name), None)

            if (file1.keys() == file2.keys()):
                print("Both files have same sheets")
                # Get the names of all sheets
                all_tabs = list(file1.keys())
                print(all_tabs)
                
                #dict_index = all_tabs.count
                dict_index = 0
                for sheet in all_tabs:
                    print("Comparing Sheet -", sheet)
                    rslt_df_for_web = compare_sheets(str(UPLOAD_FOLDER+file1_name), str(UPLOAD_FOLDER+file2_name), sheet)
                    print(rslt_df_for_web)
                    #rslt_df_for_web.to_excel(output_excel_file, sheet_name=sheet, index=False)
                    dict_df[dict_index] = rslt_df_for_web
                    dict_index = dict_index+1

            #output_excel_file.save()
            #subprocess.call(str("cd " + output_path), shell=True)
            #return render_template("input_v1.html", heading="Comparison Results are", file_names=returned_file_names, tables=[rslt_df_for_web])
            return render_template("input.html", heading="Comparison Results are", file_names=returned_file_names, tables = dict_df)
            #return render_template("input.html", heading="Comparison Results are", file_names=returned_file_names, tables=[list_df])
        else:
            returned_file_names = "File other then .xlsx or .xls"
            return render_template("input.html", heading="Comparison Results are", file_names=returned_file_names)

@app.route("/generate", methods=['POST'])
def generate_excel():
    if (request.method == 'POST'):
        output_excel_file.save()
        return render_template("input.html", folder_loc=output_excel_file.path)

if __name__ == "__main__":
    app.run()


