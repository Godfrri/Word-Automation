from docxtpl import DocxTemplate, InlineImage
from docx2pdf import convert

reportDoc = DocxTemplate('Template/reportTmpl.docx')

# Product = [ ('Furazolidone Tablets USP 100 mg (180 ml)', 2016003, '36M', 'RT',  '2023-01-29'),
#               ('Furazolidone Tablets USP 100 mg (180 ml)', 2016004, '36M', 'RT', '2023-01-29'),
#               ('H.C.T Tablets B.P 50 mg', 2022002, '36M', 'RT', '2023-01-29'),
#               ('Nalidixic Acid Tablets B.P 500 mg', 2033001, '36M', 'RT', '2023-01-26')]

Product = [ ('Furazolidone Tablets USP 100 mg (180 ml)', 2016003, '36M',  '2023-01-29'),
            ('Furazolidone Tablets USP 100 mg (180 ml)', 2016004, '36M', '2023-01-29'),
            ('H.C.T Tablets B.P 50 mg', 2022002, '36M', '2023-01-29'),
            ('Nalidixic Acid Tablets B.P 500 mg', 2033001, '36M', '2023-01-26')]

salesRows = []

for x in range(len(Product)):
    salesRows.append({  'no' : x+1, 
                        'name' : Product[x][0],
                        'cPu' : Product[x][1],
                        'nUnits' : Product[x][2],
                        'revenue' : Product[x][3]})


itemRow = ['Test_1', 'Test_2', 'Test_3']

content= {
    'reportDtStr': '23-Dec-2022',
    'salesTblRows': salesRows,
    'topItemsRows': itemRow,
    'trendImg': InlineImage(reportDoc, 'Template/download.png')
}

reportDoc.render(content)
reportDoc.save('Output/report.docx')
convert('Output/report.docx', 'Output/report.pdf')
