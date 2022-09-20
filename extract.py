 # coding=utf8
import re
line ="cosme-023　●【マーク:バラ】cosme-023　●【内容:Sakiko】●【マーク:バラ】cosme-023　●【内容:Chise】●【マーク:肉球】cosme-023　●【内容:Sakura】●【マーク:サクラ】　"



# m = re.search(r"([A-Za-z\s]+)", w)
# if m:
#     print(m.group(1))


skipwords = "baccarat cosme bvlgari parfum lamy hana bl brpp cos dior elsa tink jillstuart ursla wh bot jas stanley gr parker svt par gt whbl bell slb lmd rap swarovski whrd whpk tate tsurumi 4c waterman ivgt mbk brbl tiffany annasui samantha yl tiffany"

for w in re.split(':| |・|●|-|【|】|\t',line):
    m = re.sub("[^A-Za-z]", "", w.strip())
    if m in skipwords:
        continue
    else:
        print(m)