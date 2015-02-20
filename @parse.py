# -*- coding: cp1251 -*-
import xlwt

# =========== INPUTS: ====================

MTSBILL = 'bill_1.csv'
PARSED = 'stat_1.dif'

# ========================================
font0 = xlwt.Font()
font0.bold = True
style0 = xlwt.XFStyle()
style0.font = font0

f = open(MTSBILL, 'rU')
res = open(PARSED, 'w')
pn = 0

for line in f:
    if line.find('Контракт № ') >= 0:
        # Closing the previous detailisation file
        if pn <> 0 and pn <> 'MTS connect':
            wb.save(str(cid) + '.xls')
            print 'Processed ' + str(cid) + ' file'
        L = line.split()
        cid = int(L[2])
        if line.find('Номер телефону:') >= 0:
            pn = int(L[5])
        else:
            pn = 'MTS connect'
        if pn <> 'MTS connect':
            # new det file
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Detailing')
            row = column = 0
            header = ['Event type','Realm','Destination/Sender','Date','Time','Duration/Amount','Cost','Billed']
            for h in header:
                ws.write(row, column, h, style0)
                ws.col(column).width = 4000
                column += 1
            row = 1
            column = 0
    if pn <> 0 and pn <> 'MTS connect' and line[0:2] == ',,':
        # Detailisation line detected
        current = unicode(line[2:], 'cp1251')
        current = current.split(',')
        for c in current:
            ws.write(row, column, c)
            column += 1
        column = 0
        row += 1
        
    if line.find('ЗАГАЛОМ ЗА КОНТРАКТОМ (БЕЗ ПДВ ТА ПФ):,,,') >=0:
        total = float(line[41:-1])
        #print line,
        res.write(str(cid)+ ';' + str(round(total,2)) + '\n')

# end of file
wb.save(str(cid)+'.xls')
res.close()
f.close()
