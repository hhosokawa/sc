import csv
import os
from os.path import isfile, join
import time

# io
input_dir = 'P:/_HHOS/Microsoft'
output_dir = 'P:/_HHOS/Microsoft/'
files = []

############### f(x) ###############

def scan_ms_dir():
    for f in os.listdir(input_dir):
        if isfile(join(input_dir,f)):
            files.append(f)

def move_ms_files():
    pass
        
############### main ###############
        
if __name__ == '__main__':
    t0 = time.clock()
    scan_ms_dir()
    move_ms_files()
    t1 = time.clock()
    print 'clean_bot.py complete.', t1-t0
