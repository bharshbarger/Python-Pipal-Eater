#!/usr/bin/env python

import argparse
import os
import re
import sys
import time

from reportgen import Reportgen

class Pipal_Eater(object):

    def __init__(self):
        self.file = None
        self.verbose = False
        self.version = '0.1'
        self.pipal_file_content = {}
        self.start_time = time.time()
        self.report_generator_module = Reportgen()

    def signal_handler(self, signal, frame):
        print('You pressed Ctrl+C! Exiting...')
        sys.exit(0)

    def cls(self):
        os.system('cls' if os.name == 'nt' else 'clear')

    def cmdargs(self):
        parser = argparse.ArgumentParser()
        parser.add_argument('-f', '--file', nargs=1, metavar='pipal.txt' ,help='The file containing raw pipal output')
        parser.add_argument('-v', '--verbose', help='Optionally enable verbosity', action='store_true')
        self.args = parser.parse_args()  

    def read_file(self):
        if self.args.verbose is True:
            print('[+] Opening file {}'.format(self.args.file[0]))

        try:
            with open(self.args.file[0]) as f:
                self.pipal_file_content = (f.readlines())
                self.pipal_file_content = [x.strip() for x in self.pipal_file_content]

        except Exception as e:
            print('\n[!] Couldn\'t open file: \'{}\' Error:{}'.format(self.args.file[0],e))
            sys.exit(0)

        if self.args.verbose is True:
            for line in self.pipal_file_content:
                print(''.join(line))

    def parse(self):
        for i, line in enumerate(self.pipal_file_content):
            if 'Total entries' in line:
                self.total = line
            if 'Total unique' in line:
                self.unique = line
            #read 11 lines starting with this heading, always 10 long so range 11 works
            if 'Top 10 passwords' in line:               
                self.top_10 = []
                for z in range(11):
                    self.top_10.append(self.pipal_file_content[(i + z) % len(self.pipal_file_content)])

            #read 11 lines starting with this heading, always 10 long so range 11 works
            if 'Top 10 base words' in line:
                self.top_10_base = []
                for z in range(11):
                    self.top_10_base.append(self.pipal_file_content[(i + z) % len(self.pipal_file_content)])

            
            if 'length ordered' in line:
                self.lengths = []
                for z in range(11):
                    self.lengths.append(self.pipal_file_content[(i + z) % len(self.pipal_file_content)])


            if 'count ordered' in line:
                self.counts = line
            if 'One to six characters' in line:
                self.one_to_six = line
            if 'One to eight characters' in line:
                self.one_to_eight = line
            if 'More than eight' in line:
                self.more_than_eight = line
            if 'Single digit on the end' in line:
                self.trailing_digit = line
            if 'Last number' in line:
                self.trailing_number = line
            if 'Last digit' in line:
                self.last_1digit = line
            if 'Last 2 digits' in line:
                self.last_2digit = line
            if 'Last 3 digits' in line:
                self.last_3digit = line
            if 'Last 4 digits' in line:
                self.last_4digit = line
            if 'Last 5 digits ' in line:
                self.last_5digit = line
            if 'Character sets' in line:
                self.charset = line
            if 'Character set ordering' in line:
                self.charset_ordering = line

    def report(self):
        """run the docx report. text files happen in the respective functions"""
        self.report_generator_module.run(\
                self.total,\
                self.unique,\
                self.top_10,\
                self.top_10_base,\
                self.lengths,\
                self.counts,\
                self.one_to_six,\
                self.one_to_eight,\
                self.more_than_eight,\
                self.trailing_digit,\
                self.trailing_number,\
                self.last_1digit,\
                self.last_2digit,\
                self.last_3digit,\
                self.last_4digit,\
                self.last_5digit,\
                self.charset,\
                self.charset_ordering)



    def end(self):
        """ending stuff, right now just shows how long script took to run"""
        print('\nCompleted in {:.2f} seconds\n'.format(time.time() - self.start_time))

def main():

    run = Pipal_Eater()
    run.cls()
    run.cmdargs()
    run.read_file()
    run.parse()
    run.report()
    run.end()

if __name__ == '__main__':
    main()
