#!/usr/bin/python3
# temporary workaround to run with ipython
# sys.argv = ['']

# Parser for command-line options, arguments and sub-commands
import argparse

# check command-line arguments
parser = argparse.ArgumentParser(
    description='Check folder for certain microscope digital slide files and exports it to a CSV file.',
    formatter_class=argparse.ArgumentDefaultsHelpFormatter)
parser.add_argument(
    '-p', '--path',
    default='/media/dfsP/DIGITALE MIKROSKOPIE',
    help='set path to the folder that contains the slide files')
parser.add_argument(
    '-x', '--extensions',
    nargs='+',
    default=['mrxs', 'ndpi', 'svs', 'vmic',
             'vsf'],#, 'xlsx', 'docx', 'txt', 'csv'],
    help='set the file extensions')
parser.add_argument(
    '-s', '--splitByExtension',
    action='store_true',
    default='True',
    help='split worksheets by file extensions')
parser.add_argument(
    '-o', '--output',
    default='/media/dfsP/DIGITALE MIKROSKOPIE/digital slides.xlsx',
    # default='digital slides.xlsx',
    help='set output filename')
parser.add_argument(
    '-v', '--verbose',
    action='store_true',
    default='True',
    help='more output while the script is running')
arguments = parser.parse_args()

print('Starting script:')

# import modules
if arguments.verbose:
    print('  importing modules…', end='')

# Creating Excel XLSX files
import xlsxwriter

# Miscellaneous operating system interfaces
import os

# Object-oriented filesystem paths
import pathlib

# Regular expression operations
import re

# System-specific parameters and functions
import sys

if arguments.verbose:
    print(' done')

# initital settings
if arguments.verbose:
    print('  initializing settings…', end='')
counter = dict()
counter['all'] = 0
counter['other'] = 0
files = dict()
maxLengths = dict()
for ext in arguments.extensions:
    counter[ext] = 0
    files[ext] = list()
    maxLengths[ext] = dict()
    maxLengths[ext]['path'] = 0
    maxLengths[ext]['name'] = 0
if arguments.verbose:
    print(' done')

# get file list
if arguments.verbose:
    print('  getting file list:')
fileList = pathlib.Path(arguments.path).glob('**/*') 
# fileList = pathlib.Path(arguments.path).glob('*')
for item in fileList:
    tmpFile = dict()
    tmpFile['suffix'] = item.suffix[1:]

    # use only 'valid' file extensions
    if tmpFile['suffix'] in arguments.extensions:
        counter['all'] += 1
        counter[tmpFile['suffix']] += 1
        tmpFile['path'] = '/'.join(item.parts[:-1])[1:]
        tmpFile['name'] = item.name

        # determine longest entry for column width
        if len(tmpFile['path']) > maxLengths[tmpFile['suffix']]['path']:
            maxLengths[tmpFile['suffix']]['path'] = len(tmpFile['path'])
        if len(tmpFile['name']) > maxLengths[tmpFile['suffix']]['name']:
            maxLengths[tmpFile['suffix']]['name'] = len(tmpFile['name'])
        files[tmpFile['suffix']].append(
            [counter[tmpFile['suffix']], tmpFile['suffix'], tmpFile['path'], tmpFile['name']])
    else:
        if item.is_file():
            counter['all'] += 1
            counter['other'] += 1

# exit if no files were found
if counter['all'] == 0:
    print('[WARNING] No files found in folder \'' + arguments.path + '\'')
    sys.exit('Exiting')
if arguments.verbose:
    print('    total files found in \'' + arguments.path + '\': ' + str(counter['all']))
    for ext in arguments.extensions:
        print('      ', ext, ': ', counter[ext], sep='')
    print('      other: ', counter['other'], sep='')

# create an XLSX workbook
if arguments.verbose:
    print('  creating XLSX workbook…', end='')
workbook = xlsxwriter.Workbook(arguments.output)
if arguments.verbose:
    print(' done')

# table headers
workbookHeader = ['#', 'extension', 'file path', 'file name']

# iterate over all extensions
if arguments.verbose:
    print('  creating worksheets:')

# start in second row
row = 1
col = 0
for ext in arguments.extensions:
    if arguments.verbose:
        print('    ', ext, '…', sep='', end='')
    if arguments.splitByExtension:
        worksheet = workbook.add_worksheet(ext)

        # reset rows
        row = 1
        col = 0
    else:
        if workbook.sheetname_count == 0:
            worksheet = workbook.add_worksheet()

    # write formatted header row
    headerBold = workbook.add_format({'bold': True})
    headerBoldRight = workbook.add_format({'bold': True, 'align': 'right'})
    for i in range(len(workbookHeader)):
        worksheet.write(0, i, workbookHeader[i], (headerBold if i >= 1 else headerBoldRight))

    # freeze first row
    worksheet.freeze_panes(1, 0)

    # set autofilter
    worksheet.autofilter(0, 0, 1, 3)

    # fill in entries
    if arguments.splitByExtension:
        for number, filetype, path, filename in tuple(files[ext]):
            worksheet.write(row, col,     number)
            worksheet.write(row, col + 1, filetype)
            worksheet.write(row, col + 2, path)
            worksheet.write(row, col + 3, filename)
            row += 1
    else:
        for number, filetype, path, filename in tuple(files[ext]):
            worksheet.write(row, col,     row)
            worksheet.write(row, col + 1, filetype)
            worksheet.write(row, col + 2, path)
            worksheet.write(row, col + 3, filename)
            row += 1
    if row == 1:
        worksheet.write(row, col, 'no files found')

    # adjust column widths
    if arguments.splitByExtension:
        worksheet.set_column(0, 0, len(str(counter[ext])) + 2)
        worksheet.set_column(1, 1, max(len(ext), len(workbookHeader[1])) + 2)
        worksheet.set_column(2, 2, max(maxLengths[ext]['path'], len(workbookHeader[2])) + 2)
        worksheet.set_column(3, 3, max(maxLengths[ext]['name'], len(workbookHeader[3])) + 2)
    if arguments.verbose:
        print(' done')

# adjust column widths
if not arguments.splitByExtension:
    worksheet.set_column(0, 0, len(str(row)) + 2)
    worksheet.set_column(1, 1, max([len(i) for i in arguments.extensions] + [len(workbookHeader[1])]) + 2)
    worksheet.set_column(
        2, 2, max([maxLengths[i]['path'] for i in maxLengths] + [len(workbookHeader[2])]) + 2)
    worksheet.set_column(
        3, 3, max([maxLengths[i]['name'] for i in maxLengths] + [len(workbookHeader[3])]) + 2)

# close workbook
workbook.close()

if arguments.verbose:
    print('  total entries:', counter['all'] - counter['other'])
print('All done and exiting.')
