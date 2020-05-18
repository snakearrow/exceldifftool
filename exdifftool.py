import xlrd
import csv
import os
import logging
from os import sys

# convert excel files to csv and write csv files
def csv_from_excel(excel_file):
    worksheets = []
    filenames = []
    
    workbook = xlrd.open_workbook(excel_file)
    all_worksheets = workbook.sheet_names()
    for worksheet_name in all_worksheets:
        worksheet = workbook.sheet_by_name(worksheet_name)
        worksheets.append(worksheet_name)
        filename = u'{}_{}.csv'.format(excel_file, worksheet_name)
        filenames.append(filename)
        with open(filename, 'w') as your_csv_file:
            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL, lineterminator='\n')
            for rownum in range(worksheet.nrows):
                wr.writerow([entry for entry in worksheet.row_values(rownum)])
                
    written_csvs = {excel_file : worksheets}
    return written_csvs, filenames

def get_sheet_names(excel_file):
    written_csvs, filenames = csv_from_excel(excel_file)
    remove_csvs(filenames)
    return list(written_csvs.values())[0]
    
# find matching sheets
def filter_diffs(files1, files2):
    diffs = []
    for excel1, sheets1 in files1.items():
        for excel2, sheets2 in files2.items():
            for sheet1 in sheets1:
                for sheet2 in sheets2:
                    if sheet1 == sheet2:
                        filename1 = u'{}_{}.csv'.format(excel1, sheet1)
                        filename2 = u'{}_{}.csv'.format(excel2, sheet2)
                        t = (filename1, filename2)
                        diffs.append(t)
    return diffs

# returns True if two rows are equal
def lines_equal(line1, line2):
    cells1 = line1.split(",")
    cells2 = line2.split(",")
    cells1 = [i for i in cells1 if i]
    cells2 = [i for i in cells2 if i]
    if len(cells1) != len(cells2):
        return False
    
    for i in range(len(cells1)):
        if cells1[i] != cells2[i]:
            return False
            
    return True

# finds the cells which are not equal    
def neq_details(row_num, line1, line2):
    cells1 = line1.split(",")
    cells2 = line2.split(",")
    details = []
    
    for i in range(max(len(cells1), len(cells2))):
        cell1 = ""
        cell2 = ""
        if i < len(cells1):
            cell1 = cells1[i]
        if i < len(cells2):
            cell2 = cells2[i]
            
        if cell1 != cell2:
            details.append("neqd {}:{}".format(row_num, i)) # format: neqd row:col
            
    return details

def do_diff(filename1, filename2):
    logging.debug("\n*** do_diff of " + filename1 + " and " + filename2 + " ***")
    logging.debug(filename1 + "   **************   " + filename2)
    logging.debug('-'*60)
    
    diff = []
    
    lines1 = [line.rstrip() for line in open(filename1, 'r')]
    lines2 = [line.rstrip() for line in open(filename2, 'r')]
    
    for i in range(max(len(lines1), len(lines2))):
        line1 = ""
        line2 = ""
        if i < len(lines1):
            line1 = lines1[i]
        if i < len(lines2):
            line2 = lines2[i]
        
        # line2 was added in file2 and does not exist in file1
        if not line1:
            logging.debug('<empty>{}{}:{}'.format(' '*35, i+1, line2))
            diff.append("add {}:{}".format("-1", i))
        # line1 was added in file1 and does not exist in file2
        if not line2:
            logging.debug('{}:{}{}<empty>'.format(i+1, line1, ' '*35))
            diff.append("add {}:{}".format(i, "-1"))
            
        # both lines are present
        if line1 and line2:
            if lines_equal(line1, line2):
                logging.debug('{}:{} <equal>'.format(i+1, i+1))
                diff.append("eq {}:{}".format(i, i))
                continue
            
            # check if line was merely moved but still exists
            found = False
            k = 0
            for l in lines2:
                if lines_equal(l, line1):
                    logging.debug('{}:{} <moved> {}:{}'.format(i+1, line1, k, l))
                    diff.append("mv {}:{}".format(i, k))
                    found = True
                    break
                k += 1
            if found:
                continue
                
            # check if content differs
            logging.debug('{}:{} <neq> {}:{}'.format(i+1, line1, i+1, line2))
            for d in neq_details(i, line1, line2):
                diff.append(d)
            
    logging.debug('-'*60)
    logging.debug("\n")
    logging.debug(diff)
    return diff

def get_row_and_cols_of_csv(csv):
    rows = len(csv)
    cols = []
    for r in csv:
        cols.append(len(r.split(",")))
    return rows, max(cols)
    
def remove_csvs(filenames):
    for f in filenames:
        try:
            os.remove(f)
        except FileNotFoundError:
            continue
    
def run(filename1, filename2, sheet1, sheet2):
    files1, filenames1 = csv_from_excel(filename1)
    files2, filenames2 = csv_from_excel(filename2)
    diffs = filter_diffs(files1, files2)
    
    csv1 = [line.rstrip() for line in open(filenames1[sheet1], 'r')]
    csv2 = [line.rstrip() for line in open(filenames2[sheet2], 'r')]
    
    diffs = do_diff(filenames1[sheet1], filenames2[sheet2])
    
    remove_csvs(filenames1 + filenames2)
    return csv1, csv2, diffs
    
