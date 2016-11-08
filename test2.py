# coding: utf-8
import pandas as pd
from collections import defaultdict
import re
import sys

reload(sys)
sys.setdefaultencoding('utf-8')


df = pd.read_excel("topic_old_20161104.xlsx")


r = re.compile("[，,]")
# print len(df.values)

input = []

for line in df.values:
    for item in line[1].split('，'):
        if item.strip():
            t = tuple([line[0].strip(), item.strip()])
            input.append(t)

print(len(input))

out = defaultdict(set)
for k, v in input:
    out[v].add(k)


for k,v in out.items():
    for item in list(v):
        if item in out:
            out[k].update(out[item])

print len(out)

values = []
for k,v in out.items():
    values.append(list(v))

with open('extract.txt', 'w') as f:
    for k,v in out.items():
        flag = 0
        for value in values:
            if k in value:
                flag = 1
        if flag == 0:
            output = "%s = %s\n" % (k, ",".join(list(v)))
            f.write(output)




