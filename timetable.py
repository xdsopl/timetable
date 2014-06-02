#! /usr/bin/python
# timetable - create timetable from spreadsheet using lpsolve
# Written in 2014 by <Ahmet Inan> <ainan@mathematik.uni-freiburg.de>
# To the extent possible under law, the author(s) have dedicated all copyright and related and neighboring rights to this software to the public domain worldwide. This software is distributed without any warranty.
# You should have received a copy of the CC0 Public Domain Dedication along with this software. If not, see <http://creativecommons.org/publicdomain/zero/1.0/>.

from sys import argv, exit
from os.path import splitext
from odf.opendocument import load, OpenDocumentSpreadsheet
from odf.style import Style, TableColumnProperties, TableCellProperties, ParagraphProperties
from odf.text import P
from odf.table import Table, TableColumn, TableRow, TableCell
from lpsolve55 import lpsolve

if len(argv) != 2:
	print "usage: "+argv[0]+" input.ods"
	exit(1)

sheet = load(argv[1]).spreadsheet.getElementsByType(Table)[0]
examiners = []
dates = []
date = ""
times = []
examinees = []
examinee = ""
date_rows = True
repeat_row = 0
for row in sheet.getElementsByType(TableRow):
	if not examiners:
		for cell in row.getElementsByType(TableCell)[1:-1]:
			examiner = unicode(cell)
			examiners += [examiner]
		continue
	pos = 0
	tmp = []
	time = ""
	for cell in row.getElementsByType(TableCell):
		if not pos:
			if repeat_row:
				repeat_row -= 1
				pos += 1
			elif not unicode(cell):
				date_rows = False
				dates += [[date, times]]
				times = []
				date = []
				break
			else:
				if date_rows:
					if times:
						dates += [[date, times]]
					times = []
					date = unicode(cell)
				else:
					examinee = unicode(cell)
			repeat = cell.getAttribute("numberrowsspanned")
			if repeat:
				repeat_row = int(repeat) - 1
		if pos == 1 and date_rows:
			time = unicode(cell)
		repeat = cell.getAttribute("numbercolumnsrepeated")
		if repeat:
			repeat_cell = int(repeat)
		else:
			repeat_cell = 1
		if unicode(cell) != 'x':
			pos += repeat_cell
			continue
		while repeat_cell:
			examiner = examiners[pos-2]
			tmp += [examiner]
			repeat_cell -= 1
			pos += 1
	if time:
		times += [[time, tmp]]
		time = ""
	if examinee:
		examinees += [[examinee, tmp]]
		examinee = ""

'''
for date, times in dates:
	for time, examiners in times:
		print date, time,
		for examiner in examiners:
			print examiner,
		print

for examinee, examiners in examinees:
	print '"'+examinee+'"',
	for examiner in examiners:
		print examiner,
	print

'''

variable = 0
examinations = []
for date, times in dates:
	tmp = []
	for time, examiners in times:
		possibilities = []
		for examinee, wishing in examinees:
			if set(wishing).issubset(examiners):
				possibilities += [[variable, examinee, wishing]]
				variable += 1
		if possibilities:
			tmp += [[time, possibilities]]
	if tmp:
		examinations += [[date, tmp]]


'''
for date, times in examinations:
	for time, possibilities in times:
		print date, time, possibilities
'''

undefined = {}
for candidate, wishing in examinees:
	without_combination = True
	for date, times in examinations:
		for time, possibilities in times:
			for variable, examinee, examiners in possibilities:
				if candidate == examinee:
					without_combination = False
	if without_combination:
		undefined[candidate] = ["without combination", wishing]

max_examiners = 0
date_weights = {}
for date, times in examinations:
	histogram = {}
	for time, possibilities in times:
		for variable, examinee, examiners in possibilities:
			max_examiners = max(max_examiners, len(examiners))
			for examiner in examiners:
				if examiner in histogram:
					histogram[examiner] += 1
				else:
					histogram[examiner] = 1
	date_weights[date] = histogram

function = []
for date, times in examinations:
	weights = date_weights[date]
	for time, possibilities in times:
		for variable, examinee, examiners in possibilities:
			function += [sum(weights.values()) + sum([weights[e] for e in examiners])]

constraints = []
examinees = {}
for date, times in examinations:
	for time, possibilities in times:
		tmp = []
		for variable, examinee, examiners in possibilities:
			if not tmp:
				tmp += [0] * variable
			tmp += [1]
			if examinee in examinees:
				examinees[examinee] += [variable]
			else:
				examinees[examinee] = [variable]
		# one examination at most per timeslot
		constraints += [tmp + [0] * (len(function) - len(possibilities))]

for examinee in examinees:
	tmp = [0] * len(function)
	for variable in examinees[examinee]:
		tmp[variable] = 1
	# one examination at most per examinee
	constraints += [tmp]

lp = lpsolve('make_lp', len(constraints), len(function))
lpsolve('set_verbose', lp, 'IMPORTANT')
lpsolve('set_mat', lp, constraints)
lpsolve('set_rh_vec', lp, [1] * len(constraints))
lpsolve('set_constr_type', lp, ['LE'] * len(constraints))
lpsolve('set_obj_fn', lp, function)
lpsolve('set_maxim', lp)
lpsolve('set_binary', lp, [True] * len(function))
#lpsolve('write_lp', lp, 'a.lp')
lpsolve('solve', lp)
variables = lpsolve('get_variables', lp)[0]

'''
for date, times in examinations:
	for time, possibilities in times:
		for variable, examinee, examiners in possibilities:
			if variables[variable]:
				print date, time, '"'+examinee+'"',
				for examiner in examiners:
					print examiner,
				print
'''

examinees = {}
for date, times in examinations:
	for time, possibilities in times:
		for variable, examinee, examiners in possibilities:
			if variables[variable]:
				examinees[examinee] = examiners

for date, times in examinations:
	for time, possibilities in times:
		for variable, examinee, examiners in possibilities:
			if not examinee in examinees and not examinee in undefined:
				undefined[examinee] = ["without appointment", examiners]

'''
for examinee in undefined:
	print examinee
'''

output = OpenDocumentSpreadsheet()

middle_center = Style(name="middle_center", family="table-cell")
middle_center.addElement(TableCellProperties(verticalalign="middle"))
middle_center.addElement(ParagraphProperties(textalign="center"))
output.styles.addElement(middle_center)

middle_left = Style(name="middle_left", family="table-cell")
middle_left.addElement(TableCellProperties(verticalalign="middle"))
middle_left.addElement(ParagraphProperties(textalign="left"))
output.styles.addElement(middle_left)

date_style = Style(name="date", family="table-column")
date_style.addElement(TableColumnProperties(columnwidth="5cm"))
output.automaticstyles.addElement(date_style)

time_style = Style(name="time", family="table-column")
time_style.addElement(TableColumnProperties(columnwidth="1.5cm"))
output.automaticstyles.addElement(time_style)

examinee_style = Style(name="examinee", family="table-column")
examinee_style.addElement(TableColumnProperties(columnwidth="7cm"))
output.automaticstyles.addElement(examinee_style)

examiner_style = Style(name="examiner", family="table-column")
examiner_style.addElement(TableColumnProperties(columnwidth="4cm"))
output.automaticstyles.addElement(examiner_style)

table = Table()
table.addElement(TableColumn(stylename=date_style))
table.addElement(TableColumn(stylename=time_style))
table.addElement(TableColumn(stylename=examinee_style))
table.addElement(TableColumn(numbercolumnsrepeated=max_examiners, stylename=examiner_style))

def homogenize(examiners):
	tmp = examiners + [""] * (max_examiners - len(examiners))
	if homogenize.previous:
		for k in range(max_examiners):
			for j in range(max_examiners):
				for i in range(max_examiners):
					if tmp[j] == homogenize.previous[i]:
						if i != j:
							tmp[i], tmp[j] = tmp[j], tmp[i]
	homogenize.previous = tmp
	return tmp
homogenize.previous = []

first_cell = None
first_date = None
number_rows_spanned = 0
for date, times in examinations:
	for time, possibilities in times:
		for variable, examinee, examiners in possibilities:
			if variables[variable]:
				tr = TableRow()
				table.addElement(tr)
				tc = TableCell(stylename=middle_center)
				if first_date == date:
					number_rows_spanned += 1
				else:
					if number_rows_spanned > 1:
						first_cell.setAttribute("numberrowsspanned", number_rows_spanned)
					number_rows_spanned = 1
					first_date = date
					first_cell = tc
				tr.addElement(tc)
				tc.addElement(P(text=date))
				tc = TableCell(stylename=middle_center)
				tr.addElement(tc)
				tc.addElement(P(text=time))
				tc = TableCell(stylename=middle_left)
				tr.addElement(tc)
				tc.addElement(P(text=examinee))
				for examiner in homogenize(examiners):
					tc = TableCell(stylename=middle_center)
					tr.addElement(tc)
					tc.addElement(P(text=examiner))

if number_rows_spanned > 1:
	first_cell.setAttribute("numberrowsspanned", number_rows_spanned)

for examinee in undefined:
	tr = TableRow()
	table.addElement(tr)
	tc = TableCell(stylename=middle_center)
	tr.addElement(tc)
	tc.addElement(P(text=undefined[examinee][0]))
	tc = TableCell(stylename=middle_center)
	tr.addElement(tc)
	tc = TableCell(stylename=middle_left)
	tr.addElement(tc)
	tc.addElement(P(text=examinee))
	for examiner in homogenize(undefined[examinee][1]):
		tc = TableCell(stylename=middle_center)
		tr.addElement(tc)
		tc.addElement(P(text=examiner))

output.spreadsheet.addElement(table)
name, ext = splitext(argv[1])
output.save(name+"-timetable"+ext)

