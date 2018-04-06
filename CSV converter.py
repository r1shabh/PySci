__author__ = 'rishabh anand'


import numpy as np
import h5py
import xlsxwriter as xls


# Create a new excel file and add a worksheet
workbook = xls.Workbook('C_nodes.xlsx') #excel file where data is being written


#read hdf5 file
a = h5py.File('C_nodes.hdf5', 'r')
a.keys()



# loop through every frame
for i in range(750, 783):
    print i
    worksheet = workbook.add_worksheet()
    row = 2
    column = 1
    # write to excel file
    dataset = '/' + str(i)
    # get the names of the datasets from a group (where a group is a frame)
    frame_1 = a.get(dataset) #frame in hdf5 file from which data is being read
    # convert it to an array
    array_1 = np.array(frame_1)
    table_row = 1
    V = list(frame_1.get('V'))
    x = list(frame_1.get('x'))
    y = list(frame_1.get('y'))
    z = list(frame_1.get('z'))
    hdg = list(frame_1.get('hdg'))
    hdg_rate = list(frame_1.get('hdg_rate'))
    tid = list(frame_1.get('tid'))
    VX = list(frame_1.get('vx'))
    VY = list(frame_1.get('vy'))
    VZ = list(frame_1.get('vz'))
    worksheet.write('A1', dataset)
    worksheet.write(row, column, 'V')
    row += 1
    

    for item in V:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
    column += 2
    row = 2
    table_row = 1
    worksheet.write(row, column, 'hdg')
    row += 1
    for item in hdg:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
    column += 2
    row = 2
    table_row = 1
    worksheet.write(row, column, 'hdg_rate')
    row += 1
    for item in hdg_rate:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
    column += 2
    row = 2
    table_row = 1
    worksheet.write(row, column, 'tid')
    row += 1
    for item in tid:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
    column += 2
    row = 2
    table_row = 1
    worksheet.write(row, column, 'vx')
    row += 1
    for item in VX:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
    column += 2
    row = 2
    table_row = 1
    worksheet.write(row, column, 'vy')
    row += 1
    for item in VY:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
    column += 2
    row = 2
    table_row = 1
    worksheet.write(row, column, 'vz')
    row += 1
    for item in VZ:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
    column += 2
    row = 2
    table_row = 1
    worksheet.write(row, column, 'x')
    row += 1
    for item in x:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
    column += 2
    row = 2
    table_row = 1
    worksheet.write(row, column, 'y')
    row += 1
    for item in y:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
    column += 2
    row = 2
    table_row = 1
    worksheet.write(row, column, 'z')
    row += 1
    for item in z:
        worksheet.write(row, column, table_row)
        worksheet.write(row, column + 1, item)
        table_row += 1
        row += 1
        
workbook.close()
    

