from urllib import parse 
import xlrd
import xlwt
import re

def decode(input):
	r = parse.unquote(input)
	if re.search('(%\w\w)(%\w\w)',r) != None:
		return decode(r)
	else: 
		return r

if __name__ == '__main__':
	url = "input.xlsx"
	rd = xlrd.open_workbook(url)
	s0 = rd.sheet_by_index(0)
	col1 = s0.col_values(0)
	col2 = s0.col_values(1)
	wt = xlwt.Workbook()
	s1 = wt.add_sheet(u'sheet1',cell_overwrite_ok=True)
	for i in range(s0.nrows):
		s1.write(i,0,col1[i])
		if i == 0:
			s1.write(i,1,col2[i])
		else:
			s1.write(i,1,decode(col1[i]))
	wt.save("output.csv")
	print("Done!")
