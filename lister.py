#!/usr/bin/python3
# temporary workaround to run with ipython
# sys.argv = ['']

# to do:
# * add links to files and folders (✅ 2020-10-12)
# * calculate size of folder from header file (mrxs, vsf) (✅ 2020-10-13)
# * correct file links for MRXS files (✅ 2020-10-13)
# * fix ignoring new URLs if limit exceeds (65 530 maximum) (✅ 2020-10-14)
# * delete test files (Neuer Ordner/test) (✅ 2020-10-15)
# * JSON export for HTML search form
# * make hyperlinks optional (✅ 2020-10-15)
# * fix encoding issue (✅ 2020-11-02)
# * fix for VSI folder size (✅ 2021-03-04)

# Parser for command-line options, arguments and sub-commands
import argparse

# check for boolean value
def str2bool(v):
    if isinstance(v, bool):
        return v
    if v.lower() in ('yes', 'true', 't', 'y', '1'):
        return True
    elif v.lower() in ('no', 'false', 'f', 'n', '0'):
        return False
    else:
        raise argparse.ArgumentTypeError('Boolean value expected.')

# check command-line arguments
class Formatter(argparse.ArgumentDefaultsHelpFormatter, argparse.RawDescriptionHelpFormatter):
    pass
parser = argparse.ArgumentParser(
    description='Check folder for certain microscope digital slide files and exports it to a CSV file.\nE.g.:\n  lister.py -p path/to/myslides -x tiff jpg mrxs -s True -o output.xlsx -v True',
    formatter_class=Formatter)
parser.add_argument(
    '-p', '--path',
    default='/media/dfsP/DIGITALE MIKROSKOPIE',
    help='set path to the folder that contains the slide files')
parser.add_argument(
    '-x', '--extensions',
    nargs='+',
    default=['mrxs', 'ndpi', 'svs', 'vmic', 'vsf', 'vsi'],
    help='set the file extensions')
parser.add_argument(
    '-s', '--splitByExtension',
    type=str2bool,
    nargs='?',
    const=True,
    default=False,
    help='split worksheets by file extensions')
parser.add_argument(
    '-l', '--links',
    type=str2bool,
    nargs='?',
    const=True,
    default=False,
    help='insert hyperlinks to files and folders')
parser.add_argument(
    '-o', '--output',
    default='digital slides.xlsx',
    help='set output filename')
parser.add_argument(
    '-v', '--verbose',
    type=str2bool,
    nargs='?',
    const=True,
    default=False,
    help='more output while the script is running')
arguments=parser.parse_args()

# Basic date and time types
import datetime

# Babel: Date and Time Formatting
import babel.dates

currentDate = datetime.date.today()

print('Starting script:', babel.dates.format_date(
    currentDate, 'EEEE, MMMM dd yyyy'))

# import modules
if arguments.verbose:
    print('  importing modules…', end='')

# Creating Excel XLSX files
import xlsxwriter

# Object-oriented filesystem paths
import pathlib

# Regular expression operations
import re

# System-specific parameters and functions
import sys

# Mathematical functions
import math

if arguments.verbose:
    print(' done')

# initital settings
if arguments.verbose:
    print('  initializing settings…', end='')
counter = dict()
counter['all'] = 0
counter['other'] = 0
counter['URLs'] = 65530
uniqueSuffix = 'üöäüöäüß'
paths = dict()
paths['windows'] = 'L:'
paths['linux'] = '/media/dfsP'
files = dict()
folderSizes = dict()
maxLengths = dict()
for ext in arguments.extensions:
    counter[ext] = 0
    files[ext] = list()
    maxLengths[ext] = dict()
    maxLengths[ext]['path'] = 0
    maxLengths[ext]['name'] = 0
    maxLengths[ext]['date'] = 0
    maxLengths[ext]['size'] = 0
if arguments.verbose:
    print(' done')

# get file list
if arguments.verbose:
    print('  getting file list:')
fileList = pathlib.Path(arguments.path).glob('**/*') 
for item in fileList:
    tmpFile = dict()
    tmpFile['suffix'] = item.suffix[1:]

    # calculate folder size
    if tmpFile['suffix'] in ['img', 'dat', 'ets']:
        if re.match('(.+-level\d+\.img)|(Data\d+\.dat)', item.name):
            if str(item.parent).encode('utf8', 'surrogateescape').decode('ISO-8859-15') in folderSizes:
                folderSizes[str(item.parent).encode('utf8', 'surrogateescape').decode('ISO-8859-15')] += item.stat().st_size
            else:
                folderSizes[str(item.parent).encode('utf8', 'surrogateescape').decode('ISO-8859-15')] = item.stat().st_size
        if re.match('frame_t\.ets', item.name):
            if str(item.parent.parent).encode('utf8', 'surrogateescape').decode('ISO-8859-15') in folderSizes:
                folderSizes[str(item.parent.parent).encode('utf8', 'surrogateescape').decode('ISO-8859-15')] += item.stat().st_size
            else:
                folderSizes[str(item.parent.parent).encode('utf8', 'surrogateescape').decode('ISO-8859-15')] = item.stat().st_size

    # use only 'valid' file extensions
    if tmpFile['suffix'] in arguments.extensions:
        counter['all'] += 1
        counter[tmpFile['suffix']] += 1

        # remove base linux path
        tmpFile['path'] = '\\'.join(item.parts[:-1])[len(paths['linux']) + 2:]

        # correct path for MRXS files
        if tmpFile['suffix'] == 'mrxs':
            tmpFile['path'] += '\\' + item.name[:-5]
        # correct path for VSI files
        if tmpFile['suffix'] == 'vsi':
            tmpFile['path'] += '\\_' + item.name[:-4] + '_'
        tmpFile['path'] = tmpFile['path'].encode('utf8', 'surrogateescape').decode('ISO-8859-15')
        tmpFile['name'] = item.name.encode('utf8', 'surrogateescape').decode('ISO-8859-15')
        tmpFile['date'] = datetime.datetime.fromtimestamp(item.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
        tmpFile['size'] = item.stat().st_size

        # determine longest entry for column width
        if len(tmpFile['path']) > maxLengths[tmpFile['suffix']]['path']:
            maxLengths[tmpFile['suffix']]['path'] = len(tmpFile['path'])
        if len(tmpFile['name']) > maxLengths[tmpFile['suffix']]['name']:
            maxLengths[tmpFile['suffix']]['name'] = len(tmpFile['name'])
        if len(tmpFile['date']) > maxLengths[tmpFile['suffix']]['date']:
            maxLengths[tmpFile['suffix']]['date'] = len(tmpFile['date'])
        if len(str(tmpFile['size'])) > maxLengths[tmpFile['suffix']]['size']:
            maxLengths[tmpFile['suffix']]['size'] = len(str(tmpFile['size']))

        # add file to list
        files[tmpFile['suffix']].append(
            [counter[tmpFile['suffix']], tmpFile['suffix'], tmpFile['path'], tmpFile['name'], tmpFile['date'], tmpFile['size']])
    else:
        if item.is_file():
            counter['all'] += 1
            counter['other'] += 1

# insert folder sizes after full iteration
if 'vsf' in arguments.extensions:
    for item in files['vsf']:
        if paths['linux'] + '/' + item[2].replace('\\', '/') in folderSizes:
            files['vsf'][item[0] - 1][5] = folderSizes[paths['linux'] + '/' + item[2].replace('\\', '/')]
            if len(str(files['vsf'][item[0] - 1][5])) > maxLengths['vsf']['size']:
                maxLengths['vsf']['size'] = len(str(files['vsf'][item[0] - 1][5]))
if 'mrxs' in arguments.extensions:
    for item in files['mrxs']:
        if paths['linux'] + '/' + item[2].replace('\\', '/') in folderSizes:
            files['mrxs'][item[0] - 1][5] = folderSizes[paths['linux'] + '/' + item[2].replace('\\', '/')]
            if len(str(files['mrxs'][item[0] - 1][5])) > maxLengths['mrxs']['size']:
                maxLengths['mrxs']['size'] = len(str(files['mrxs'][item[0] - 1][5]))
        else:
            files['mrxs'][item[0] - 1][2] += uniqueSuffix
if 'vsi' in arguments.extensions:
    for item in files['vsi']:
        if paths['linux'] + '/' + item[2].replace('\\', '/') in folderSizes:
            files['vsi'][item[0] - 1][5] = folderSizes[paths['linux'] + '/' + item[2].replace('\\', '/')]
            if len(str(files['vsi'][item[0] - 1][5])) > maxLengths['vsi']['size']:
                maxLengths['vsi']['size'] = len(
                    str(files['vsi'][item[0] - 1][5]))
        else:
            files['vsi'][item[0] - 1][2] += uniqueSuffix

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
workbookHeader = ['#', 'extension', 'file path', 'file name', 'file date', 'file size']

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

    # declare different formats
    headerBold = workbook.add_format({'bold': True})
    headerBoldRight = workbook.add_format({'bold': True, 'align': 'right'})
    numberSpace = workbook.add_format(
        {'num_format': '### ### ### ### ### ### ### ##0'})
    for i in range(len(workbookHeader)):
        worksheet.write(0, i, workbookHeader[i], (headerBoldRight if (i == 0 or i == 5) else headerBold))

    # freeze first row
    worksheet.freeze_panes(1, 0)

    # set autofilter
    worksheet.autofilter(0, 0, 1, 5)

    # fill in entries
    if arguments.splitByExtension:
        for number, filetype, path, filename, filedate, filesize in tuple(files[ext]):
            worksheet.write(row, col, number, numberSpace)
            worksheet.write(row, col + 1, filetype)

            if ext in ['mrxs', 'vsi']:
                if re.match('.+' + uniqueSuffix + '$', path):
                    shortPath = re.match('^(.+)\\\.+' + uniqueSuffix + '$', path).group(1)
                    worksheet.write(row, col + 2, shortPath)
                elif arguments.links and counter['URLs'] > 0:
                    worksheet.write_url(row, col + 2, paths['windows'] + '\\' + path, string=path)
                    counter['URLs'] -= 1
                else:
                    worksheet.write(row, col + 2, path)
            else:
                worksheet.write(row, col + 2, path)
            if arguments.links and counter['URLs'] > 0:
                if ext in ['mrxs', 'vsi']:
                    worksheet.write_url(row, col + 3, paths['windows'] + '\\' + re.match('^(.+)\\\.+$', path).group(1) + '\\' + filename, string=filename)
                else:
                    worksheet.write_url(row, col + 3, paths['windows'] + '\\' + path + '\\' + filename, string=filename)
                counter['URLs'] -= 1
            else:
                worksheet.write(row, col + 3, filename)
            worksheet.write(row, col + 4, filedate)
            worksheet.write(row, col + 5, filesize, numberSpace)
            row += 1
    else:
        for number, filetype, path, filename, filedate, filesize in tuple(files[ext]):
            worksheet.write(row, col, row, numberSpace)
            worksheet.write(row, col + 1, filetype)

            if ext in ['mrxs', 'vsi']:
                if re.match('.+' + uniqueSuffix + '$', path):
                    shortPath = re.match('^(.+)\\\.+' + uniqueSuffix + '$', path).group(1)
                    worksheet.write(row, col + 2, shortPath)
                elif arguments.links and counter['URLs'] > 0:
                    worksheet.write_url(
                        row, col + 2, paths['windows'] + '\\' + path, string=path)
                    counter['URLs'] -= 1
                else:
                    worksheet.write(row, col + 2, path)
            else:
                worksheet.write(row, col + 2, path)
            if arguments.links and counter['URLs'] > 0:
                if ext in ['mrxs', 'vsi']:
                    worksheet.write_url(row, col + 3, paths['windows'] + '\\' + re.match(
                        '^(.+)\\\.+$', path).group(1) + '\\' + filename, string=filename)
                else:
                    worksheet.write_url(
                        row, col + 3, paths['windows'] + '\\' + path + '\\' + filename, string=filename)
                counter['URLs'] -= 1
            else:
                worksheet.write(row, col + 3, filename)
            worksheet.write(row, col + 4, filedate)
            worksheet.write(row, col + 5, filesize, numberSpace)
            row += 1
    if row == 1:
        worksheet.write(row, col, 'no files found')

    # adjust length for thousands separators
    maxLengths[ext]['size'] += math.floor((maxLengths[ext]['size'] - 1) / 3)

    # adjust column widths
    if arguments.splitByExtension:
        worksheet.set_column(0, 0, len(str(counter[ext])) + math.floor((len(str(counter[ext])) - 1) / 3) + 3)
        worksheet.set_column(1, 1, max(len(ext), len(workbookHeader[1])) + 2)
        worksheet.set_column(2, 2, max(maxLengths[ext]['path'], len(workbookHeader[2])) + 2)
        worksheet.set_column(3, 3, max(maxLengths[ext]['name'], len(workbookHeader[3])) + 2)
        worksheet.set_column(4, 4, max(maxLengths[ext]['date'], len(workbookHeader[4])) + 2)
        worksheet.set_column(5, 5, max(maxLengths[ext]['size'], len(workbookHeader[5])) + 2)
    if arguments.verbose:
        print(' done')

# adjust column widths
if not arguments.splitByExtension:
    worksheet.set_column(0, 0, len(str(row)) + math.floor((len(str(row)) - 1) / 3) + 3)
    worksheet.set_column(1, 1, max([len(i) for i in arguments.extensions] + [len(workbookHeader[1])]) + 2)
    worksheet.set_column(
        2, 2, max([maxLengths[i]['path'] for i in maxLengths] + [len(workbookHeader[2])]) + 2)
    worksheet.set_column(
        3, 3, max([maxLengths[i]['name'] for i in maxLengths] + [len(workbookHeader[3])]) + 2)
    worksheet.set_column(
        4, 4, max([maxLengths[i]['date'] for i in maxLengths] + [len(workbookHeader[4])]) + 2)
    worksheet.set_column(
        5, 5, max([maxLengths[i]['size'] for i in maxLengths] + [len(workbookHeader[5])]) + 2)

# close workbook
workbook.close()

if arguments.verbose:
    print('  total entries:', counter['all'] - counter['other'])
print('All done and exiting.\n')
