#!usr/bin/env python3.4
# -*- coding: utf-8 -*-
# #  FileName    : common.py
# #  Author      : Administrator
# #  Description : 
# #  Time        : 2016/11/10

import getopt
import sys

USAGE_STRING = u'''
Example:

python generateDB.py [-d lanzhou/beijing/shanghai | --all]

python generateDB.py [-h | -d | --all ] [value]

-h (--help):
            Dispaly usage info

-d (--depot):
            Add a single depot dwg to database

--all:
            Add all depot dwgs to database

-o(--output):
            Set product dir
'''
class Options(object):
    def __int__(self):
        self.xls2json = False

OPTIONS = Options()

def Usage(docString):
    print (docString.rstrip("\n"))
    print (USAGE_STRING)

def ParseOptions(argv, docString,
                 extra_opts = "", extra_long_opts = [],
                 extra_opts_handler = None):
    try:
        opts, args = getopt.getopt(
                    argv, "h"+extra_opts,
                    []+list(extra_long_opts))

    except getopt.GetoptError as err:
        print (">>> %s <<<" %err)
        sys.exit(2)
    for p, v in (opts):
        if p in ("-h", "--help"):
            Usage(docString)
            sys.exit()
        elif p in ("-v", "--verbose"):
            OPTIONS.verbose = True
        else:
            if extra_opts_handler is None or not extra_opts_handler(p, v):
                assert False, "unknown options \"%s\"" % (p, v)
    return args