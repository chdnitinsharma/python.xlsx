import sys,urllib2,json,xlsxwriter;
import pprint
from StringIO import StringIO


pp = pprint.PrettyPrinter(indent=4);

#Dummy Json will be here
dummyJsonString='[{"description": " Description ","title": "Business / Administration Services"},{"description": " Description1 ","title": "Python Job","brand":"Python"}]';
io = StringIO(dummyJsonString);


#Load your Json From API
payload=json.load(io);

# [(unique key match with you json field , Description on excel Header), ... ]
headerDesc=[("title","Title"),("description","Description"),("brand","Brand")];
header,excel_Header=zip(*headerDesc);


collection=[];
var_len=1;
for collection_ in payload:
	temp={};
	for record in collection_:
		if record in header:
			temp[record]=collection_[record];
	collection.append(temp);


#If you want to remove any unique key not to come in excel
#Add/Remove dynamic header in Excel header
removeHeaderKey=["variant"];
headerDesc=filter(lambda x: x[0] not in removeHeaderKey,headerDesc);
header,excel_Header=zip(*headerDesc);


#Adding dict into list to write records in excel file
record=[];
for t in collection:
	temp={};
	for r in t:
		if r in header:
			temp[r]=t[r];
	record.append(temp);


#create xlsx file
workbook = xlsxwriter.Workbook('leads.xlsx')
worksheet = workbook.add_worksheet()

row = 0

#write header first in excel xlsx
for index,attr in enumerate(header):
	worksheet.write(row,index,excel_Header[index])
	
row=row+1;

#writing rows in file
for attrCollection in record:
	for attr in attrCollection:
		worksheet.write(row,int(header.index(attr)),attrCollection[attr])
	row=row+1;

#file close
workbook.close()

#print record
pp.pprint(record)	

