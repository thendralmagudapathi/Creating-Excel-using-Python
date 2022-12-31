import xlsxwriter

entry = [
    {
        'S.No' : '1',
        'Assembly Parts' : 'Lever SA',
        'Standard Part(Yes/No)' : 'No',
        'Revision Number' : '2',
        'Status' : 'Not Released',
        'Part Number' : 'L1254',
        'Weight in grams' : '120',
        'Material' : 'SS316',
        'Quantity' : '1',
        'Notes' : '1. Remove Burns 2.Chamfer the sharp edges',
        'Date' : '22/10/2023',
        'Designer' : 'James',
        'Detailed' : 'John',
        'Approved by' : 'Michael',
        'Custom Scale' : '0.5',
        'Customer paper size' : 'A1',
        'Orientation' : ' '
    },
    {
        'S.No' : '2', 
        'Assembly Parts' : 'Bolt_M6',
        'Standard Part(Yes/No)' : 'Yes',
        'Revision Number' : '10',
        'Status' : ' ',
        'Part Number' : 'M6B234',
        'Weight in grams' : '8',
        'Material' : ' ',
        'Quantity' : ' ',
        'Notes' : ' ',
        'Date' : ' ',
        'Designer' : ' ',
        'Detailed' : ' ',
        'Approved by' : ' ',
        'Custom Scale' : ' ',
        'Customer paper size' : ' ',
        'Orientation' : ' ',
    },
    {
        'S.No' : '3',
        'Assembly Parts' : 'Lever SA',
        'Standard Part(Yes/No)' : 'No',
        'Revision Number' : '2',
        'Status' : 'Not Released',
        'Part Number' : 'L1254',
        'Weight in grams' : '120',
        'Material' : 'SS316',
        'Quantity' : '1',
        'Notes' : '1. Remove Burns 2.Chamfer the sharp edges',
        'Date' : '22/10/2023',
        'Designer' : 'James',
        'Detailed' : 'John',
        'Approved by' : 'Michael',
        'Custom Scale' : '0.5',
        'Customer paper size' : 'A1',
        'Orientation' : ' '
    }
]

workbook = xlsxwriter.Workbook('Creating Excel Task.xlsx')
worksheet = workbook.add_worksheet("Sheet 1")

worksheet.write(0,0, "S.No")
worksheet.write(0,1, "Assembly parts")
worksheet.write(0,2, "Standard part")
worksheet.write(0,3, "Revision Number")
worksheet.write(0,4, "Status")
worksheet.write(0,5, "Part Number")
worksheet.write(0,6, "Weight")
worksheet.write(0,7, "Material")
worksheet.write(0,8, "Quantity")
worksheet.write(0,9, "Notes")
worksheet.write(0,10, "Date")
worksheet.write(0,11, "Designed By")
worksheet.write(0,12, "Detailed")
worksheet.write(0,13, "Approved By")
worksheet.write(0,14, "Custom Scale")
worksheet.write(0,15, "Custom Paper Size")
worksheet.write(0,16, "Orientation")

for i, e  in enumerate(entry):
     worksheet.write(i+1, 0, e["S.No"])
     worksheet.write(i+1, 1, e["Assembly Parts"])
     worksheet.write(i+1, 2, e["Standard Part(Yes/No)"])
     worksheet.write(i+1, 3, e["Revision Number"])
     worksheet.write(i+1, 4, e["Status"])
     worksheet.write(i+1, 5, e["Part Number"])
     worksheet.write(i+1, 6, e["Weight in grams"])
     worksheet.write(i+1, 7, e["Material"])
     worksheet.write(i+1, 8, e["Quantity"])
     worksheet.write(i+1, 9, e["Notes"])
     worksheet.write(i+1, 10, e["Date"])
     worksheet.write(i+1, 11, e["Designer"])
     worksheet.write(i+1, 12, e["Detailed"])
     worksheet.write(i+1, 13, e["Approved by"])
     worksheet.write(i+1, 14, e["Custom Scale"])
     worksheet.write(i+1, 15, e["Customer paper size"])
     worksheet.write(i+1, 16, e["Orientation"])

workbook.close()