import xlsxwriter as xls
import random as rand
from collections import Counter

'''
Andrew Nash
python excel writer Assignment 2
This program creates a list of random numbers sorts it and counts the number of times an int is found
Then it charts it.
'''
# create the workbook and worksheet assignments
workbook = xls.Workbook('ExcelAssignment1.xlsx')
worksheet = workbook.add_worksheet()

# Create a new Chart object.
chart = workbook.add_chart({'type': 'bar'})

# Write some data to add to plot on the chart.
data = []
foo = []
# random init
rand.seed(128)

# fill foo with random data then add the sorted list to data and the unsorted to data
for x in range(50):
    foo.append(rand.randint(0,50))
data.append(sorted(foo))
data.append(foo)

# write my name and set a column width
worksheet.write(0, 0, 'Andrew Nash')
worksheet.set_column('A:A', 20)
# write my lists to columns on the worksheet
worksheet.write_column('B1', data[0])
worksheet.write_column('C1', data[1])

# Configure the chart
# add data series for both columns

chart.add_series({'values': '=Sheet1!$B$1:$B$50', 'name': 'Sorted values'})
chart.add_series({'values': '=Sheet1!$C$1:$C$50', 'name': 'random values'})

# add labels
chart.set_x_axis({'name': 'Cell number', 'name_font': {'size': 14, 'bold': True}})

# Insert the chart into the worksheet.
worksheet.insert_chart('E2', chart)

# create another chart
chart = workbook.add_chart({'type': 'bar', 'name': 'Number of Times int was found'})

# count the number of times an int is found
boo = Counter(foo)
# clear foo
foo = [0] * 50

# remove the zeroth element from boo as it will always be zero
del boo[0]
# for each key value in the boo dictionary add to the list
for k,v in boo.items():
    foo[k-1] = v

# write this list to a new column
worksheet.write_column('D1', foo)
# add the chart to the worksheet
chart.add_series({'values': '=Sheet1!$D$1:$D$50', 'name': 'occurrences'})

# add labels
chart.set_x_axis({'name': 'Integer', 'name_font': {'size': 14, 'bold': True}})

# Insert the chart into the worksheet.
worksheet.insert_chart('E25', chart)

# close the connection
workbook.close()