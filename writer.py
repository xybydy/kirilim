import os

import xlsxwriter
from sqlalchemy import desc, asc

import db

NUM_FORMAT = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
PERC_FORMAT = '0%'
GRI_BG = '#D9D9D9'
EY_SARI = '#FFFF99'
TOP_HEAD = '#F2F2F2'
BLUE = '#0000D4'
TBS_PEMBE = '#FDE9D9'
PL_HESAPLAR = ['UA', 'UB', 'X', 'VA', 'VC', 'VE']

directory = os.path.abspath(".")
cikti_folder = os.path.join(directory, "cikti")


def create_workbook(isim):
    global workbook
    workbook = xlsxwriter.Workbook(os.path.join(cikti_folder, '{}.xlsx'.format(isim)))


def create_tbs():
    worksheet = workbook.add_worksheet('TBs')
    red = workbook.add_format({'bold': True, 'font_color': 'red'})
    header = workbook.add_format({'bold': True, 'bg_color': GRI_BG, 'top': 1, 'bottom': 1})
    isim = workbook.add_format({'bg_color': GRI_BG})
    data = workbook.add_format({'bg_color': TBS_PEMBE, 'num_format': NUM_FORMAT})

    worksheet.set_zoom(80)
    worksheet.hide_gridlines(2)

    row = 1
    column = 0

    worksheet.write_string('B1', 'Control', red)
    worksheet.write_string(1, 0, 'Acc', header)
    worksheet.write_string(1, 1, 'Description', header)

    column += 2

    for k, v in db.periods.items():

        if v is '':
            continue

        worksheet.write_string(row, column, '%s' % v, header)
        column += 1

    row += 1

    ana_hesaplar = db.session.query(db.Hesaplar).filter_by(len=3).order_by(db.Hesaplar.number).all()

    for item in ana_hesaplar:
        worksheet.write_string(row, 0, item.number, isim)
        worksheet.write_string(row, 1, item.name, isim)
        worksheet.write_number(row, 2, item.py1, data)
        worksheet.write_number(row, 3, item.py2, data)
        worksheet.write_number(row, 4, item.cy, data)
        row += 1

    worksheet.write_formula(0, 2, '=SUM(C3:C%s)' % row, workbook.add_format({'num_format': NUM_FORMAT}))
    worksheet.write_formula(0, 3, '=SUM(D3:D%s)' % row, workbook.add_format({'num_format': NUM_FORMAT}))
    worksheet.write_formula(0, 4, '=SUM(E3:E%s)' % row, workbook.add_format({'num_format': NUM_FORMAT}))


def create_comperative():
    worksheet = workbook.add_worksheet('Comperative TB')
    worksheet.set_zoom(80)
    worksheet.hide_gridlines(2)

    row = 6
    col = 1

    money = workbook.add_format({'num_format': NUM_FORMAT})
    border = workbook.add_format()
    border.set_top(1)
    border.set_bottom(1)
    border.set_bold()
    border.set_num_format(NUM_FORMAT)
    border.set_bg_color(GRI_BG)

    worksheet.write_string(row, col, 'Acc.', border)
    worksheet.write_string(row, col + 1, 'Description', border)
    worksheet.write_string(6, 6, 'Change', border)

    col += 2

    create_headline(worksheet)

    for k, v in db.periods.items():

        if v is '':
            continue

        worksheet.write_string(row, col, '%s' % v, border)
        col += 1

    row += 1

    ana_hesaplar = db.session.query(db.Hesaplar).filter_by(len=3).order_by(db.Hesaplar.number).all()

    for item in ana_hesaplar:
        worksheet.write_string(row, 1, item.number)
        worksheet.write_string(row, 2, item.name)
        worksheet.write_number(row, 3, item.py1, money)
        worksheet.write_number(row, 4, item.py2, money)
        worksheet.write_number(row, 5, item.cy, money)
        worksheet.write_formula(row, 6, '=F%s-E%s' % (row + 1, row + 1), money)
        row += 1


def create_headline(worksheet, tip='bd'):
    # 3 satir 7 sutun kadar
    topline = workbook.add_format()
    topline.set_bg_color(GRI_BG)
    right = workbook.add_format()
    right.set_right(1)
    bottom = workbook.add_format()
    bottom.set_bottom(1)
    bottom.set_bg_color(GRI_BG)
    right.set_bg_color(GRI_BG)

    double = workbook.add_format()
    double.set_bg_color(GRI_BG)
    double.set_bottom(1)
    double.set_right(1)

    ff_1 = workbook.add_format({'bold': True, 'bg_color': GRI_BG})
    ff_2 = workbook.add_format({'bold': True, 'bg_color': GRI_BG, 'bottom': 1})

    for x in xrange(3):
        for y in xrange(9):
            if x == 2 and y == 8:
                worksheet.write_blank(x, y, None, double)
            elif y == 1 and x == 0:
                worksheet.write_string(x, y, db.tanimlar['company'], ff_1)
            elif y == 1 and x == 1:
                worksheet.write_string(x, y, db.periods['cy'], ff_1)
            elif y == 1 and x == 2:
                if tip == 'bd':
                    worksheet.write_string(x, y, 'BD of Accounts', ff_2)
                elif tip == 'tb':
                    worksheet.write_string(x, y, 'All TBs', ff_2)
                else:
                    baslik = db.session.query(db.Lead).filter_by(lead_code=tip).first().name
                    worksheet.write_string(x, y, baslik, ff_2)
            elif x == 2:
                worksheet.write_blank(x, y, None, bottom)
            elif y == 8:
                worksheet.write_blank(x, y, None, right)
            else:
                worksheet.write_blank(x, y, None, topline)


def create_alltb():
    worksheet = workbook.add_worksheet('TB Mapping')
    worksheet.set_zoom(80)
    worksheet.hide_gridlines(2)
    # 1,86 3,86
    row = 5
    col = 4

    border = workbook.add_format()
    border.set_top(1)
    border.set_bottom(1)
    border.set_bold()
    border.set_num_format(NUM_FORMAT)
    border.set_bg_color(GRI_BG)

    bold = workbook.add_format()
    bold.set_bold()

    money = workbook.add_format()
    money.set_num_format(NUM_FORMAT)

    worksheet.set_column(0, 0, 1.86)
    worksheet.set_column(0, 0, 3.86)

    create_headline(worksheet, 'tb')

    for x in xrange(1, 4):
        if x is 1:
            worksheet.write_blank(row, x, None,
                                  workbook.add_format({'bg_color': GRI_BG, 'left': 1, 'top': 1, 'bottom': 1}))
        else:
            worksheet.write_blank(row, x, None, border)

    for k, v in db.periods.items():
        if col is 6:
            worksheet.write_string(row, col, '%s' % v, workbook.add_format(
                {'bold': True, 'bg_color': GRI_BG, 'right': 1, 'top': 1, 'bottom': 1}))
        else:
            worksheet.write_string(row, col, '%s' % v, border)
        col += 1

    for i in xrange(6, 11):
        worksheet.write_blank(i, 1, None, workbook.add_format({'left': 1}))
        if i is 10:
            worksheet.write_blank(i, 1, None, workbook.add_format({'left': 1, 'bottom': 1}))

    worksheet.write_string(6, 2, 'Assets', bold)
    worksheet.write_string(7, 2, 'Liabilities', bold)
    worksheet.write_string(8, 2, 'Shareholders\' equity', bold)
    worksheet.write_string(9, 2, 'Income & expenses', bold)
    worksheet.write_string(10, 2, 'Suspense accounts', workbook.add_format({'bold': 1, 'bottom': 1}))
    worksheet.write_formula(6, 4, '=SUM(E16:E24)', money)
    worksheet.write_formula(7, 4, '=SUM(E25:E29)', money)
    worksheet.write_formula(8, 4, '=E30', money)
    worksheet.write_formula(9, 4, '=SUM(E31:E35)', money)
    worksheet.write_formula(10, 4, '=E36+E37', workbook.add_format({'bottom': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_formula(6, 5, '=SUM(F16:F24)', money)
    worksheet.write_formula(7, 5, '=SUM(F25:F29)', money)
    worksheet.write_formula(8, 5, '=F30', money)
    worksheet.write_formula(9, 5, '=SUM(F31:F35)', money)
    worksheet.write_formula(10, 5, '=F36+F37', workbook.add_format({'bottom': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_formula(6, 6, '=SUM(G16:G24)', workbook.add_format({'right': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_formula(7, 6, '=SUM(G25:G29)', workbook.add_format({'right': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_formula(8, 6, '=G30', workbook.add_format({'right': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_formula(9, 6, '=SUM(G31:G35)', workbook.add_format({'right': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_formula(10, 6, '=G36+G37', workbook.add_format({'right': 1, 'bottom': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_blank(12, 1, None, workbook.add_format({'left': 1, 'bottom': 1, 'top': 1}))
    worksheet.write_string(12, 2, 'Control',
                           workbook.add_format({'top': 1, 'bottom': 1, 'bold': True, 'num_format': NUM_FORMAT}))
    worksheet.write_blank(12, 3, workbook.add_format({'top': 1, 'bottom': 1}))
    worksheet.write_formula(12, 4, '=SUM(E7:E11)',
                            workbook.add_format({'top': 1, 'bottom': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_formula(12, 5, '=SUM(F7:F11)',
                            workbook.add_format({'top': 1, 'bottom': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_formula(12, 6, '=SUM(G7:G11)',
                            workbook.add_format({'right': 1, 'bottom': 1, 'top': 1, 'num_format': NUM_FORMAT}))

    worksheet.write_blank(10, 3, None, workbook.add_format({'bottom': 1}))
    worksheet.write_blank(12, 3, None, workbook.add_format({'bottom': 1, 'top': 1}))

    worksheet.write_blank(14, 1, None, workbook.add_format(
        {'bg_color': GRI_BG, 'top': 1, 'bottom': 1, 'left': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_string(14, 2, 'Account Groups', border)
    worksheet.write_string(14, 3, 'Code', border)

    row = 14
    col = 4

    for k, v in db.periods.items():
        if col is 6:
            worksheet.write_string(row, col, '%s' % v, workbook.add_format(
                {'top': 1, 'bottom': 1, 'right': 1, 'bg_color': GRI_BG, 'bold': True}))
        else:
            worksheet.write_string(row, col, '%s' % v, border)
        col += 1

    row = 15

    lead_b = 15

    for item in db.session.query(db.Lead).group_by(db.Lead.lead_code).order_by(db.Lead.lead_code).all():
        if item.lead_code == 'Unmapped':
            continue

        if item.lead_code is not "X":
            # worksheet.write_blank(row,0,None,workbook.add_format({'left':1,'bg_color':EY_SARI}))
            worksheet.write_blank(row, 1, None, workbook.add_format({'left': 1, 'bg_color': EY_SARI}))
            worksheet.write_string(row, 2, item.name, workbook.add_format({'bg_color': EY_SARI}))
            worksheet.write_string(row, 3, item.lead_code,
                                   workbook.add_format({'bg_color': EY_SARI, 'bold': 1, 'font_color': 'red'}))
        else:
            if not len(db.session.query(db.Hesaplar).filter_by(lead_code='Unmapped').all()) > 0:
                # workbook.write_blank(row,0,None,workbook.add_format({'left':1,'bg_color':EY_SARI,'bottom':1}))
                worksheet.write_blank(row, 1, None, workbook.add_format({'left': 1, 'bg_color': EY_SARI, 'bottom': 1}))
                worksheet.write_string(row, 2, item.name, workbook.add_format({'bg_color': EY_SARI, 'bottom': 1}))
                worksheet.write_string(row, 3, item.lead_code, workbook.add_format(
                    {'bg_color': EY_SARI, 'bold': 1, 'font_color': 'red', 'bottom': 1}))
        row += 1

    if len(db.session.query(db.Hesaplar).filter_by(lead_code='Unmapped').all()) > 0:
        worksheet.write_blank(row, 1, None, workbook.add_format({'left': 1, 'bg_color': EY_SARI, 'bottom': 1}))
        worksheet.write_string(row, 2, 'Unmapped Accounts', workbook.add_format({'bg_color': EY_SARI, 'bottom': 1}))
        worksheet.write_string(row, 3, 'UNMAPPED',
                               workbook.add_format({'bg_color': EY_SARI, 'bold': 1, 'font_color': 'red', 'bottom': 1}))

    row += 2
    lead_s = row - 1

    col = 4

    worksheet.write_string(row, 1, 'Acc',
                           workbook.add_format({'bg_color': GRI_BG, 'bold': True, 'left': 1, 'bottom': 1, 'top': 1}))
    worksheet.write_string(row, 2, 'Description', border)
    worksheet.write_blank(row, 3, None, border)

    for k, v in db.periods.items():

        if v is '':
            continue

        if v is db.periods.get('cy'):
            worksheet.write_string(row, col, '%s' % v, workbook.add_format(
                {'bg_color': GRI_BG, 'bold': True, 'bottom': 1, 'top': 1, 'right': 1, 'align': 'right'}))

        worksheet.write_string(row, col, '%s' % v, border)
        col += 1

    row += 1

    hesap_b = row + 1

    for item in db.session.query(db.Hesaplar).filter_by(len=3).order_by(db.Hesaplar.number).all():
        worksheet.write_string(row, 1, item.number, workbook.add_format({'bg_color': EY_SARI, 'left': 1}))
        worksheet.write_string(row, 2, item.name, workbook.add_format({'bg_color': EY_SARI}))
        worksheet.write_string(row, 3, item.lead_code,
                               workbook.add_format({'bg_color': EY_SARI, 'bold': True, 'font_color': 'red'}))
        worksheet.write_formula(row, 4, "=SUMIF('TBs'!A:A,B%s,'TBs'!C:C)" % (row + 1),
                                workbook.add_format({'bg_color': EY_SARI, 'num_format': NUM_FORMAT}))
        worksheet.write_formula(row, 5, "=SUMIF('TBs'!A:A,B%s,'TBs'!D:D)" % (row + 1),
                                workbook.add_format({'bg_color': EY_SARI, 'num_format': NUM_FORMAT}))
        worksheet.write_formula(row, 6, "=SUMIF('TBs'!A:A,B%s,'TBs'!E:E)" % (row + 1),
                                workbook.add_format({'bg_color': EY_SARI, 'right': 1, 'num_format': NUM_FORMAT}))
        row += 1

    for i in xrange(1, 7):
        worksheet.write_blank(row, i, None, workbook.add_format({'top': 1}))

    hesap_s = row

    for x in xrange(lead_b, lead_s):
        if x < hesap_b - 4:
            worksheet.write_formula(x, 4, '=SUMIF(D%s:D%s,D%s,E%s:E%s)' % (hesap_b, hesap_s, x + 1, hesap_b, hesap_s),
                                    workbook.add_format({'num_format': NUM_FORMAT, 'bg_color': EY_SARI}))
            worksheet.write_formula(x, 5, '=SUMIF(D%s:D%s,D%s,F%s:F%s)' % (hesap_b, hesap_s, x + 1, hesap_b, hesap_s),
                                    workbook.add_format({'num_format': NUM_FORMAT, 'bg_color': EY_SARI}))
            worksheet.write_formula(x, 6, '=SUMIF(D%s:D%s,D%s,G%s:G%s)' % (hesap_b, hesap_s, x + 1, hesap_b, hesap_s),
                                    workbook.add_format({'num_format': NUM_FORMAT, 'right': 1, 'bg_color': EY_SARI}))
        else:
            worksheet.write_formula(x, 4, '=SUMIF(D%s:D%s,D%s,E%s:E%s)' % (hesap_b, hesap_s, x + 1, hesap_b, hesap_s),
                                    workbook.add_format({'bottom': 1, 'num_format': NUM_FORMAT, 'bg_color': EY_SARI}))
            worksheet.write_formula(x, 5, '=SUMIF(D%s:D%s,D%s,F%s:F%s)' % (hesap_b, hesap_s, x + 1, hesap_b, hesap_s),
                                    workbook.add_format({'bottom': 1, 'num_format': NUM_FORMAT, 'bg_color': EY_SARI}))
            worksheet.write_formula(x, 6, '=SUMIF(D%s:D%s,D%s,G%s:G%s)' % (hesap_b, hesap_s, x + 1, hesap_b, hesap_s),
                                    workbook.add_format(
                                        {'bottom': 1, 'num_format': NUM_FORMAT, 'right': 1, 'bg_color': EY_SARI}))


def create_lead(hesap):
    worksheet = workbook.add_worksheet('%s Lead' % hesap)
    worksheet.set_zoom(80)
    worksheet.hide_gridlines(2)

    money = workbook.add_format({'num_format': NUM_FORMAT})
    border = workbook.add_format()
    border.set_top(1)
    border.set_bottom(1)
    border.set_bold()
    border.set_bg_color(GRI_BG)
    percn = workbook.add_format({'bold': True, 'font_color': BLUE, 'align': 'right', 'num_format': PERC_FORMAT})

    dipler = workbook.add_format(
        {'top': 1, 'bottom': 1, 'bold': True, 'num_format': NUM_FORMAT, 'bg_color': EY_SARI})

    row = 6
    col = 1

    worksheet.write_string(row, col, 'Acc.', border)
    worksheet.write_string(row, col + 1, 'Description', border)

    col += 2

    for k, v in db.periods.items():

        if v is '':
            continue

        worksheet.write_string(row, col, '%s' % v, border)
        col += 1

    worksheet.write_string(row, col, 'WP Ref', workbook.add_format(
        {'top': 1, 'bottom': 1, 'bold': True, 'bg_color': GRI_BG, 'color': 'red'}))
    worksheet.write_string(row, col + 1, '% change', border)

    for item in db.session.query(db.Hesaplar).filter_by(len=3, lead_code=hesap).order_by(db.Hesaplar.number).all():
        row += 1
        worksheet.write_string(row, 1, '%s' % item.number)
        worksheet.write_string(row, 2, '%s' % item.name)
        worksheet.write_number(row, 3, item.py1, money)
        worksheet.write_number(row, 4, item.py2, money)
        worksheet.write_number(row, 5, item.cy, money)
        if hesap in PL_HESAPLAR and db.periods.get("cy")[3:5] != "12":
            worksheet.write_formula(row, 7, '=IF(D%s=0,"INF",IF(F%s=0,"-100",(F%s-D%s)/F%s))' % (
                row + 1, row + 1, row + 1, row + 1, row + 1), percn)
        else:
            worksheet.write_formula(row, 7, '=IF(E%s=0,"INF",IF(F%s=0,"-100",(F%s-E%s)/F%s))' % (
                row + 1, row + 1, row + 1, row + 1, row + 1), percn)

    row += 2

    worksheet.write_blank(row, 1, None, dipler)
    worksheet.write_string(row, 2, 'TOTAL', dipler)

    worksheet.write_formula(row, 3, '=SUM(D8:D%s)' % row, dipler)
    worksheet.write_formula(row, 4, '=SUM(E8:E%s)' % row, dipler)
    worksheet.write_formula(row, 5, '=SUM(F8:F%s)' % row, dipler)
    worksheet.write_blank(row, 6, None, dipler)
    worksheet.write_blank(row, 7, None, dipler)

    row += 2

    worksheet.write_string(row, 1, hesap,
                           workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1, 'left': 1}))
    worksheet.write_string(row, 2, 'Check',
                           workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1}))
    worksheet.write_formula(row, 3, "=VLOOKUP($B%s,'TB Mapping'!$D:$G,2,0)-D%s" % (row + 1, row - 1),
                            workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1}))
    worksheet.write_formula(row, 4, "=VLOOKUP($B%s,'TB Mapping'!$D:$G,3,0)-E%s" % (row + 1, row - 1),
                            workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1}))
    worksheet.write_formula(row, 5, "=VLOOKUP($B%s,'TB Mapping'!$D:$G,4,0)-F%s" % (row + 1, row - 1),
                            workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1}))
    worksheet.write_blank(row, 6, None, workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1}))
    worksheet.write_blank(row, 7, None,
                          workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1, 'right': 1}))
    create_headline(worksheet, hesap)


def create_breakdown(hesap):
    worksheet = workbook.add_worksheet('%s1 - BD' % hesap)
    worksheet.set_zoom(80)
    worksheet.hide_gridlines(2)

    money = workbook.add_format({'num_format': NUM_FORMAT})
    header = workbook.add_format()
    header.set_top(1)
    header.set_bottom(1)
    header.set_bold()
    header.set_bg_color(GRI_BG)
    color = workbook.add_format()
    color.set_bold()
    color.set_color('red')

    percn = workbook.add_format({'bold': True, 'font_color': BLUE, 'align': 'right', 'num_format': PERC_FORMAT})
    dipler = workbook.add_format(
        {'top': 1, 'bottom': 1, 'bold': True, 'num_format': NUM_FORMAT, 'bg_color': EY_SARI})
    create_headline(worksheet)

    row = 6
    col = 1
    start = 8

    for item in db.session.query(db.Hesaplar).filter_by(lead_code=hesap, len=3).order_by(db.Hesaplar.number).all():
        worksheet.write_string(row, col, 'Acc.', header)
        worksheet.write_string(row, col + 1, 'Descripton', header)
        col += 2
        for k, v in db.periods.items():

            if v is '':
                continue

            worksheet.write_string(row, col, '%s' % v, header)
            col += 1
        worksheet.write_string(row, col, 'Change', header)
        worksheet.write_string(row, col + 1, 'Flux', header)

        col = 1

        list_of_bds = db.session.query(db.Hesaplar).filter_by(ana_hesap=item.number, bd=True).order_by(
            desc(db.Hesaplar.cy)).all() if (item.lead_code not in {"M", "N", "O", "P", "T", "UA"} or (
            item.ana_hesap in {"302", "402", "322", "422", "437", "337", "610", "611", "612", "653", "642", "654",
                               "655",
                               "659", "680", "681", "682", "689", "656", "657", "660", "661"} and item.lead_code in {
                "UB",
                "VE",
                "M",
                "N",
                "UA"})) else db.session.query(
            db.Hesaplar).filter_by(ana_hesap=item.number, bd=True).order_by(asc(db.Hesaplar.cy)).all()

        for eben in list_of_bds:
            row += 1
            worksheet.write_string(row, 1, '%s' % eben.number)
            worksheet.write_string(row, 2, '%s' % eben.name)
            worksheet.write_number(row, 3, eben.py1, money)
            worksheet.write_number(row, 4, eben.py2, money)
            worksheet.write_number(row, 5, eben.cy, money)
            if hesap in PL_HESAPLAR and db.periods.get("cy")[3:5] != "12":
                worksheet.write_formula(row, 6, '=F%s-D%s' % (row + 1, row + 1), money)
                worksheet.write_formula(row, 7, '=IF(D%s=0,"INF",IF(F%s=0,"-100",(F%s-D%s)/F%s))' % (
                    row + 1, row + 1, row + 1, row + 1, row + 1), percn)
            else:
                worksheet.write_formula(row, 6, '=F%s-E%s' % (row + 1, row + 1), money)
                worksheet.write_formula(row, 7, '=IF(E%s=0,"INF",IF(F%s=0,"-100",(F%s-E%s)/F%s))' % (
                    row + 1, row + 1, row + 1, row + 1, row + 1), percn)

        row += 2

        for i in xrange(1, 8):
            if i == 2:
                worksheet.write_string(row, i, 'Total', dipler)
                worksheet.write_string(row + 2, i, 'Check', dipler)
            elif i == 3:
                worksheet.write_formula(row, i, '=SUM(D%s:D%s)' % (start, row - 1), dipler)
                worksheet.write_string(row + 1, i, '|--- %s Lead ---|' % hesap, color)
                worksheet.write_formula(row + 2, i,
                                        "=IFERROR(VLOOKUP(\"%s\",'%s Lead'!B:F,3,0),0)-D%s" % (
                                            item.number, hesap, row + 1), dipler)
            elif i == 4:
                worksheet.write_formula(row, i, '=SUM(E%s:E%s)' % (start, row - 1), dipler)
                worksheet.write_string(row + 1, i, '|--- %s Lead ---|' % hesap, color)
                worksheet.write_formula(row + 2, i,
                                        "=IFERROR(VLOOKUP(\"%s\",'%s Lead'!B:F,4,0),0)-E%s" % (
                                            item.number, hesap, row + 1), dipler)
            elif i == 5:
                worksheet.write_formula(row, i, '=SUM(F%s:F%s)' % (start, row - 1), dipler)
                worksheet.write_string(row + 1, i, '|--- %s Lead ---|' % hesap, color)
                worksheet.write_formula(row + 2, i,
                                        "=IFERROR(VLOOKUP(\"%s\",'%s Lead'!B:F,5,0),0)-F%s" % (
                                            item.number, hesap, row + 1), dipler)
            else:
                worksheet.write_blank(row, i, None, dipler)
                worksheet.write_blank(row + 2, i, None, dipler)

        row += 6
        start = row + 2


def create_leads():
    for k in db.session.query(db.Lead).group_by(db.Lead.lead_code).order_by(db.Lead.lead_code).all():
        if k.lead_code == 'Unmapped':
            continue
        if not db.session.query(db.Hesaplar).filter_by(lead_code=k.lead_code).first():
            continue

        create_lead(k.lead_code)
        create_breakdown(k.lead_code)

    if len(db.session.query(db.Lead).filter_by(lead_code='Unmapped').all()) > 0:
        create_lead('Unmapped')
        create_breakdown('Unmapped')


def create_a4():
    # bitmedi daha
    file_name = 'a4'

    if db.tanimlar['company'] and db.periods['cy']:
        file_name = '{0} - {1}'.format(db.tanimlar['company'].encode('utf-8'), db.periods['cy'])

    create_workbook(file_name)
    create_tbs()
    create_comperative()
    create_alltb()
    # fill_the_blanks()
    create_leads()
    workbook.close()
