#interpretation needs to be written 'whole chromosome(wc) gain(loss/CN-LOH)' in any upper/lowercase
#copy and past interpretation and nomenclature columns into wchange.text file
import re
import os
#change directory if needed
#os.chdir(r'C:\Users\CytoLab\Desktop\Python programs')
file = open('wcchange.txt')

for line in file:
    line = line.replace('-', '~')
    if line.lower().startswith("whole chromosome loss") and line.lower().endswith("hmz\n"):
        lo = re.search(r'(.*)\t(.*)\s(\d{1,2}|\w)(.*)(x|\s)(.*)\n', line)
        buildlo = lo.group(2)
        chromelo = lo.group(3)
        cnlo = lo.group(5)
        countlo = lo.group(6)
        print(buildlo + " " + "(" + chromelo + ")" + cnlo + "x1")
    elif line.lower().startswith("whole chromosome cn-loh") and line.lower().endswith("hmz\n"):
        g = re.search(r'(.*)\t(.*)\s(\d{1,2}|\w)(.*)(x|\s)(.*)\n', line)
        buildg = g.group(2)
        chromeg = g.group(3)
        cng = g.group(5)
        countg = g.group(6)
        print(buildg + " " + "(" + chromeg + ")" + cng + "x2 hmz")
    elif line.lower().startswith("whole chromosome gain") and line.lower().endswith("hmz\n"):
        g = re.search(r'(.*)\t(.*)\s(\d{1,2}|\w)(.*)(x|\s)(.*)\n', line)
        buildg = g.group(2)
        chromeg = g.group(3)
        cng = g.group(5)
        countg = g.group(6)
        print(buildg + " " + "(" + chromeg + ")" + cng + "x3")
    elif line.lower().startswith(("wc", "whole chromosome")):
        w = re.search(r'(.*)\t(.*)\s(\d{1,2}|\w)(.*)(x|\s)(.*)\n', line)
        build = w.group(2)
        chrome = w.group(3)
        cn = w.group(5)
        count = w.group(6)
        print(build + " " + "(" + chrome + ")" + cn + count)
    elif line.startswith('\t') and line.endswith("hmz\n"):
        w = re.search(r'(.*)\t(.*)\s(\d{1,2}|\w)(.*)(x|\s)(.*)\n', line)
        build = w.group(2)
        chrome = w.group(3)
        coor = w.group(4)
        cn = w.group(5)
        count = w.group(6)
        print(build + " " + chrome + coor + cn + "x2 " + count)
    elif line.startswith('\t'):
        t = re.search(r'\t(.*)', line)
        same = t.group(1)
        print(same)
    else:
        l = re.search(r'(.*)\t(.*)', line)
        same2 = l.group(2)
        print(same2)
