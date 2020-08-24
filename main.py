from zipfile import ZipFile
import os
import time
import re

start = time.process_time()

def find_sheet_id(workbook_xml_string, sheetName):
    result = re.search('<sheets>(.*)</sheets>', workbook_xml_string)
    return re.search(f'''<sheet name="{sheetName}".*(?<=r:id=)"(.*)"/>''', result.group(1)).group(1)

def find_sheet_file(workbook_rel_string, sheet_rid):
    return re.search(f'''<Relationships xmlns=".*?">.*?<Relationship Id="{sheet_rid}".*?(?<=Target=)"(.*?)"/>.*?</Relationships>''', workbook_rel_string).group(1)

def find_table_rid_of_sheet(sheet_xml):
    return re.search('<tableParts.*?>.*?r:id="(.*?)".*?/></tableParts>', sheet_xml).group(1)

def read_file_data_into_string(filepath):
    buffer = ""
    with zip_file.open(filepath) as f:
        for line in f:
            buffer += str(line)
    return buffer

def print_all_files(files):
    for filename in files:
        if not os.path.isdir(filename):
            with zip_file.open(filename) as f:
                for line in f:
                    print(line)

zip_file = ZipFile('Sample.xlsx')
files = zip_file.namelist()

workbook_xml_string = read_file_data_into_string('xl/workbook.xml')
data_sheet_id = find_sheet_id(workbook_xml_string, "Data")

workbook_rels_xml_string = read_file_data_into_string('xl/_rels/workbook.xml.rels')
data_sheet_file = find_sheet_file(workbook_rels_xml_string, data_sheet_id)
data_sheet_file_path = f'''xl/{data_sheet_file}'''

data_sheet_xml_string = read_file_data_into_string(data_sheet_file_path)

table_rid = find_table_rid_of_sheet(data_sheet_xml_string)

print(data_sheet_file_path)
print(table_rid)