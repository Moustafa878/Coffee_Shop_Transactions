from typing import Dict
import xlrd
from itertools import combinations


loc = ("demo2.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0,0)
support=input("enter minimum support :")
confidence=input("enter minimum confidence  :")
supportcount=(float(support)*(sheet.nrows-1))/100
print("minimum support count :  ",supportcount)
print('*************************************1***************************************')
my_dict={
    "m":6,
    "o":4,
    "k":8,
    "f":9
}
dellist=[]
newlist=[]
my_dict2 ={}
my_dict3={}
print ('1 item set before deleting :', my_dict)
for i in list(my_dict):
   if float(my_dict[i]) < float(supportcount):
      my_dict.pop(i)
      dellist.append(i)

print ('1 item set after deleting :', my_dict)
print( 'deleting item from 1 item set : ',dellist)
print('***************************************2*************************************')
comb1 = combinations(my_dict.keys(), 2)


for i in list(comb1):
   for k in range(1, sheet.nrows):
    #  if  i[0]  in dellist or i[1] in dellist:
     #     continue
      #else:
        if i[0] in sheet.row_values(k) and i[1] in sheet.row_values(k):
            if i in my_dict2:
               my_dict2[i] +=1

            else:
             my_dict2.update({i: 1})

print ('2 item set before delete', my_dict2)
dellist.clear()
for i in list(my_dict2):
   if float(my_dict2[i]) < float(supportcount):
       my_dict2.pop(i)
       dellist.append(i)

   else:
       for k in list(i):
          newlist.append(k)
newlist=list(dict.fromkeys(newlist))
print ('2 item set after deleting : ',my_dict2)
#print(newlist)
print( 'deleting item from 2 item set : ',dellist)
print('***************************************3*************************************')

comb2 = combinations(newlist, 3)
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



#print('comb2edit',comb2edit)
for i in list(comb2edit):
    for k in range(1, sheet.nrows):
               if i[0] in sheet.row_values(k) and i[1] in sheet.row_values(k) and i[2] in sheet.row_values(k):
                   if i in my_dict3:
                        my_dict3[i] +=1

               else:
                 my_dict3.update({i: 1})

print ('3 item set before deleting :',my_dict3)
dellist.clear()
for i in list(my_dict3):
   if float(my_dict3[i]) < float(supportcount):
      my_dict3.pop(i)
      dellist.append(i)
print ('3 item set after deleting :', my_dict3)
print( 'deleting item from 3 item set : ',dellist)
print('***************************************min_conf*************************************')
all_values = my_dict3.values()
max_value = max(all_values)
for i in list(my_dict3):
    if float(my_dict3[i]) < max_value:
        my_dict3.pop(i)


for i in list(my_dict3):

    for j in list(i):
        x=float((my_dict3[i]/(sheet.nrows-1))/(my_dict[j]/(sheet.nrows-1)))
        if (x*100)< float(confidence):
          print ("conf = support",i,'/ support(',j,')=',x,' => ',x*100,' weak ')
        else:
            print("conf = support", i, '/ support(', j, ')=', x, ' => ', x * 100, ' strong ')
    comb3=combinations(i,2)
    for w in list(comb3):
        y=float((my_dict3[i]/(sheet.nrows-1))/(my_dict2[w]/(sheet.nrows-1)))
        if (y*100)<float(confidence):
           print ("conf = support",i,'/ support',w,'=',y,' => ',x*100,' weak ')
        else:
            print("conf = support", i, '/ support', w, '=', y, ' => ', x * 100, ' strong ')
