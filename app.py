import openpyxl
xfile = openpyxl.load_workbook('test.xlsx')
year='III/V'
specialization='CSE'
sec='C3'
timing="1pm-2pm"

sheet = xfile.get_sheet_by_name('Sheet1')
sheet['B1'] = 'Kannan Ramu'
sheet['E1']='22.07.22'
lis=[121212,233232,232323,232323,23232,232323,232323,233322,121212,233232,232323,232323,23232,232323,232323,233322]
for i in range(3,len(lis)+3):
    sheet['A{}'.format(i)]=year
    sheet['B{}'.format(i)]=specialization
    sheet['C{}'.format(i)]=sec
    sheet['D{}'.format(i)]=timing
    sheet['E{}'.format(i)]=lis[i-3]

xfile.save('text2.xlsx')