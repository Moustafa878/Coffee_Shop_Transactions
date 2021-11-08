from typing import Dict
import xlrd
from itertools import combinations


loc = ("CoffeeShopTransactions.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0,0)
supportcount=input("enter minimum support count :")
confidence=input("enter minimum confidence  :")

one_item_set = {}
two_item_set = {}
three_item_set = {}
print('*************************************1***************************************')
for i in range(1,sheet.nrows):
    for j in  range(3,6):
      if sheet.cell_value(i, j) in one_item_set:
         if sheet.cell_value(i, 3) or \
                 (sheet.cell_value(i, 4)!=sheet.cell_value(i, 5) and sheet.cell_value(i, 4)!=sheet.cell_value(i, 3)) or\
                 (sheet.cell_value(i, 5)!=sheet.cell_value(i, 4) and sheet.cell_value(i, 4)!=sheet.cell_value(i, 3)):
           one_item_set[sheet.cell_value(i, j)]+=1
         else:
             continue
      else:
          one_item_set.update({sheet.cell_value(i, j):1})
dellist=[]
newlist=[]

print ('1 item set before deleting :', one_item_set)
for i in list(one_item_set):
   if float(one_item_set[i]) < float(supportcount):
      one_item_set.pop(i)
      dellist.append(i)

print ('1 item set after deleting :', one_item_set)
print( 'deleting item from 1 item set : ',dellist)
print('***************************************2*************************************')
comb1 = combinations(one_item_set.keys(), 2)


for i in list(comb1):
   for k in range(1, sheet.nrows):
    #  if  i[0]  in dellist or i[1] in dellist:
     #     continue
      #else:
        if i[0] in sheet.row_values(k) and i[1] in sheet.row_values(k):
            if i in two_item_set:
               two_item_set[i] +=1

            else:
             two_item_set.update({i: 1})

print ('2 item set before delete', two_item_set)
dellist.clear()
for i in list(two_item_set):
   if float(two_item_set[i]) < float(supportcount):
       two_item_set.pop(i)
       dellist.append(i)


print ('2 item set after deleting : ',two_item_set)
#print(newlist)
print( 'deleting item from 2 item set : ',dellist)
print('***************************************3*************************************')

comb2 = combinations(one_item_set, 3)
x=[]
comb2edit=[]
if dellist==[]:
    comb2edit=comb2
else:
    for i in list(comb2):

       for t in list(dellist):
          if  all(item in i for item in t):
            x.append(i)
            break
          else:
              comb2edit.append(i)
              continue

    for item in comb2edit:
        if item in x:
            comb2edit.remove(item)
    comb2edit = list(dict.fromkeys(comb2edit))




for i in list(comb2edit):
    for k in range(1, sheet.nrows):
               if i[0] in sheet.row_values(k) and i[1] in sheet.row_values(k) and i[2] in sheet.row_values(k):
                   if i in three_item_set:
                        three_item_set[i] +=1

               else:
                 three_item_set.update({i: 1})

print ('3 item set before deleting :',three_item_set)
dellist.clear()
for i in list(three_item_set):
   if float(three_item_set[i]) < float(supportcount):
      three_item_set.pop(i)
      dellist.append(i)
print ('3 item set after deleting :', three_item_set)
print( 'deleting item from 3 item set : ',dellist)
print('***************************************min_conf*************************************')
if not three_item_set=={}:
    all_values = three_item_set.values()
    max_value = max(all_values)
    for i in list(three_item_set):
        if float(three_item_set[i]) < max_value:
            three_item_set.pop(i)


    for i in list(three_item_set):

        for j in list(i):
            x=float(three_item_set[i]/one_item_set[j])
            if (x*100)< float(confidence):
              print ("conf = support",i,'/ support(',j,')=',x,' => ',x*100,' weak ')
            else:
                print("conf = support", i, '/ support(', j, ')=', x, ' => ', x * 100, ' strong ')
        comb3=combinations(i,2)
        for w in list(comb3):
            y=float((three_item_set[i]/(sheet.nrows-1))/(two_item_set[w]/(sheet.nrows-1)))
            if (y*100)<float(confidence):
               print ("conf = support",i,'/ support',w,'=',y,' => ',x*100,' weak ')
            else:
                print("conf = support", i, '/ support', w, '=', y, ' => ', x * 100, ' strong ')
else:
        all_values = two_item_set.values()
        max_value = max(all_values)
        for i in list(two_item_set):
            if float(two_item_set[i]) < max_value:
                two_item_set.pop(i)
        for i in list(two_item_set):

            for j in list(i):
                x = float((two_item_set[i] / (sheet.nrows - 1)) / (one_item_set[j] / (sheet.nrows - 1)))
                if (x * 100) < float(confidence):
                    print("conf = support", i, '/ support(', j, ')=', x, ' => ', x * 100, ' weak ')
                else:
                    print("conf = support", i, '/ support(', j, ')=', x, ' => ', x * 100, ' strong ')
