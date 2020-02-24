#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Rob Spears (GitHub: Forty9Unbeaten)"

import argparse
import os
from zipfile import ZipFile

# parse command-line arguments
parser = argparse.ArgumentParser(
    description='Find a string of text in a dotm file')
parser.add_argument('text', help='The text to search for')
parser.add_argument(
    '--directory', '-d',
    help='The directory to search (default is current directory)',
    dest='folder',
    default='.')
args = parser.parse_args()


def main():

    print('\n\tFiles containing ' + args.text)
    print('\t-------------------\n')
    files_searched = 0
    file_matches = 0

    # traverse the file tree and find files with .dotm extension
    for path, folders, files in os.walk(args.folder):
        for file in files:
            files_searched += 1

            # find files with .dotm file extensions
            if '.dotm' in file:
                file_path = os.path.join(path, file)
                document = ZipFile(file_path, 'r').read('word/document.xml')
                byte_string = args.text.encode('utf-8')

                # find files with arg text in them
                if byte_string in document:
                    text_index = document.index(byte_string)
                    sample_text = document[text_index -
                                           35:text_index+35].decode('utf-8')

                    # print file information
                    print('\tFile Path:\t' + file_path)
                    print('\tSample:\t...' + sample_text + '...\n')
                    print('\t*****************\n')
                    file_matches += 1

    # print scraping report
    print('\tFiles Found:\t' + str(file_matches))
    print('\tFiles Searched:\t' + str(files_searched) + '\n')


if __name__ == '__main__':
    main()
