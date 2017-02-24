__author__ = 'fatihka'

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

# ASCII A harfi 65'den baslar. hr kalkip ascii'dan devam edeyim


W_ZOOM = 70

dir = os.path.abspath(".")
cikti_folder = os.path.join(dir, "cikti")

if not os.path.exists(cikti_folder):
    os.mkdir(cikti_folder)


def create_workbook(isim):
    global workbook
    workbook = xlsxwriter.Workbook(os.path.join(cikti_folder, '{}.xlsx'.format(isim)))


def create_tbs():
    worksheet = workbook.add_worksheet('TBs')
    red = workbook.add_format({'bold': True, 'font_color': 'red'})
    header = workbook.add_format({'bold': True, 'bg_color': GRI_BG, 'top': 1, 'bottom': 1})
    isim = workbook.add_format({'bg_color': GRI_BG})
    data = workbook.add_format({'bg_color': TBS_PEMBE, 'num_format': NUM_FORMAT})

    worksheet.set_zoom(W_ZOOM)
    worksheet.hide_gridlines(2)

    worksheet.write_string(0, 1, 'Control', red)
    worksheet.write_string(1, 0, 'Acc', header)
    worksheet.write_string(1, 1, 'Description', header)

    row = 1
    column = 2

    for k in db.periodss:
        worksheet.write_string(row, column, '%s' % k, header)
        column += 1

    row += 1

    ana_hesaplar = db.session.query(db.Hesaplar).filter_by(len=3).order_by(db.Hesaplar.number).all()

    for item in ana_hesaplar:
        worksheet.write_row(row, 0, [item.number, item.name], isim)

        for period in db.periodss:
            worksheet.write_number(row, db.periodss.index(period) + 2, getattr(item, period), data)

        row += 1

    for period in range(db.len_periods):
        worksheet.write_formula(0, period + 2, '=SUM({0}3:{0}{1})'.format(chr(65 + 2 + period), row),
                                workbook.add_format({'num_format': NUM_FORMAT}))


def create_comperative():
    """
    en son buna bakicam
    """
    worksheet = workbook.add_worksheet('Comperative TB')
    worksheet.set_zoom(W_ZOOM)
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

    for k in db.periodss:
        worksheet.write_string(row, col, '%s' % k, border)
        col += 1

    row += 1

    ana_hesaplar = db.session.query(db.Hesaplar).filter_by(len=3).order_by(db.Hesaplar.number).all()

    for item in ana_hesaplar:
        worksheet.write_string(row, 1, item.number)
        worksheet.write_string(row, 2, item.name)
        for period in db.periodss:
            worksheet.write_number(row, db.periodss.index(period) + 3, getattr(item, period), money)
        worksheet.write_formula(row, 6, '=F%s-E%s' % (row + 1, row + 1), money)
        row += 1


def create_headline(worksheet, type='bd', en=9):
    # 3 satir 9 sutun kadar
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

    for x in range(3):
        for y in range(en):
            if x == 2 and y == en - 1:
                worksheet.write_blank(x, y, None, double)
            elif y == 1 and x == 0:
                worksheet.write_string(x, y, db.tanimlar['company'], ff_1)
            elif y == 1 and x == 1:
                worksheet.write_string(x, y, db.periodss[-1], ff_1)
            elif y == 1 and x == 2:
                if type == 'bd':
                    worksheet.write_string(x, y, 'BD of Accounts', ff_2)
                elif type == 'tb':
                    worksheet.write_string(x, y, 'All TBs', ff_2)
                else:
                    baslik = db.session.query(db.Lead).filter_by(lead_code=type).first().name
                    worksheet.write_string(x, y, baslik, ff_2)
            elif x == 2:
                worksheet.write_blank(x, y, None, bottom)
            elif y == en - 1:
                worksheet.write_blank(x, y, None, right)
            else:
                worksheet.write_blank(x, y, None, topline)


def create_alltb():
    worksheet = workbook.add_worksheet('TB Mapping')
    worksheet.set_zoom(W_ZOOM)
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

    # worksheet.set_column(0, 0, 1.86)
    # worksheet.set_column(0, 0, 3.86)

    create_headline(worksheet, 'tb')

    unmp = len(db.session.query(db.Hesaplar).filter_by(lead_code='Unmapped').all())

    for x in range(1, 4):
        if x is 1:
            worksheet.write_blank(row, x, None,
                                  workbook.add_format({'bg_color': GRI_BG, 'left': 1, 'top': 1, 'bottom': 1}))
        else:
            worksheet.write_blank(row, x, None, border)

    for k in db.periodss:
        if col is db.periodss[-1]:
            worksheet.write_string(row, col, '%s' % k, workbook.add_format(
                {'bold': True, 'bg_color': GRI_BG, 'right': 1, 'top': 1, 'bottom': 1}))
        else:
            worksheet.write_string(row, col, '%s' % k, border)

        col += 1

    for i in range(6, 11):
        worksheet.write_blank(i, 1, None, workbook.add_format({'left': 1}))
        if i is 10:
            worksheet.write_blank(i, 1, None, workbook.add_format({'left': 1, 'bottom': 1}))
    # todo: buralari sigorta ve diger sirketleri goz onunde bulundrarak rakam kismini dinamic hale getirmek lazim.
    worksheet.write_string(6, 2, 'Assets', bold)
    worksheet.write_string(7, 2, 'Liabilities', bold)
    worksheet.write_string(8, 2, 'Shareholders\' equity', bold)
    worksheet.write_string(9, 2, 'Income & expenses', bold)
    worksheet.write_string(10, 2, 'Suspense accounts', workbook.add_format({'bold': 1, 'bottom': 1}))
    worksheet.write_blank(10, 3, None, workbook.add_format({'bottom': 1}))

    for col in range(db.len_periods):
        if col == db.len_periods - 1:
            fformat = workbook.add_format({'num_format': NUM_FORMAT, 'right': 1})

            worksheet.write_formula(6, col + 4, '=SUM({0}16:{0}24)'.format(chr(65 + 4 + col)),
                                    fformat)
            worksheet.write_formula(7, col + 4, '=SUM({0}25:{0}29)'.format(chr(65 + 4 + col)),
                                    fformat)
            worksheet.write_formula(8, col + 4, '={}30'.format(chr(65 + 4 + col)),
                                    fformat)
            worksheet.write_formula(9, col + 4, '=SUM({0}31:{0}35)'.format(chr(65 + 4 + col)),
                                    fformat)
            # todo: alttaki guncellenecek
            worksheet.write_formula(10, col + 4, '={0}36+{0}37'.format(chr(65 + 4 + col)),
                                    workbook.add_format({'bottom': 1, 'num_format': NUM_FORMAT, 'right': 1}))
        else:
            worksheet.write_formula(6, col + 4, '=SUM({0}16:{0}24)'.format(chr(65 + 4 + col)), money)
            worksheet.write_formula(7, col + 4, '=SUM({0}25:{0}29)'.format(chr(65 + 4 + col)), money)
            worksheet.write_formula(8, col + 4, '={}30'.format(chr(65 + 4 + col)), money)
            worksheet.write_formula(9, col + 4, '=SUM({0}31:{0}35)'.format(chr(65 + 4 + col)), money)
            # todo: alttaki guncellenecek
            worksheet.write_formula(10, col + 4, '={0}36+{0}37'.format(chr(65 + 4 + col)),
                                    workbook.add_format({'bottom': 1, 'num_format': NUM_FORMAT}))

    worksheet.write_blank(12, 1, None, workbook.add_format({'left': 1, 'bottom': 1, 'top': 1}))
    worksheet.write_string(12, 2, 'Control',
                           workbook.add_format({'top': 1, 'bottom': 1, 'bold': True, 'num_format': NUM_FORMAT}))
    worksheet.write_blank(12, 3, None, workbook.add_format({'top': 1, 'bottom': 1}))

    for col in range(db.len_periods):
        if col == db.len_periods - 1:
            worksheet.write_formula(12, col + 4, '=SUM({0}7:{0}11)'.format(chr(65 + 4 + col)),
                                    workbook.add_format({'top': 1, 'right': 1, 'bottom': 1, 'num_format': NUM_FORMAT}))
        else:
            worksheet.write_formula(12, col + 4, '=SUM({0}7:{0}11)'.format(chr(65 + 4 + col)),
                                    workbook.add_format({'top': 1, 'bottom': 1, 'num_format': NUM_FORMAT}))

    worksheet.write_blank(14, 1, None, workbook.add_format(
        {'bg_color': GRI_BG, 'top': 1, 'bottom': 1, 'left': 1, 'num_format': NUM_FORMAT}))
    worksheet.write_string(14, 2, 'Account Groups', border)
    worksheet.write_string(14, 3, 'Code', border)

    row = 14
    col = 4

    for k in db.periodss:
        if col is db.len_periods + 3:
            worksheet.write_string(row, col, '%s' % k, workbook.add_format(
                {'top': 1, 'bottom': 1, 'right': 1, 'bg_color': GRI_BG, 'bold': True}))
        else:
            worksheet.write_string(row, col, '%s' % k, border)
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
            if not unmp > 0:
                # workbook.write_blank(row,0,None,workbook.add_format({'left':1,'bg_color':EY_SARI,'bottom':1}))
                worksheet.write_blank(row, 1, None, workbook.add_format({'left': 1, 'bg_color': EY_SARI, 'bottom': 1}))
                worksheet.write_string(row, 2, item.name, workbook.add_format({'bg_color': EY_SARI, 'bottom': 1}))
                worksheet.write_string(row, 3, item.lead_code, workbook.add_format(
                    {'bg_color': EY_SARI, 'bold': 1, 'font_color': 'red', 'bottom': 1}))
            else:
                worksheet.write_blank(row, 1, None, workbook.add_format({'left': 1, 'bg_color': EY_SARI}))
                worksheet.write_string(row, 2, item.name, workbook.add_format({'bg_color': EY_SARI}))
                worksheet.write_string(row, 3, item.lead_code,
                                       workbook.add_format({'bg_color': EY_SARI, 'bold': 1, 'font_color': 'red'}))
        row += 1

    if unmp > 0:
        worksheet.write_blank(row, 1, None, workbook.add_format({'left': 1, 'bg_color': EY_SARI, 'bottom': 1}))
        worksheet.write_string(row, 2, 'Unmapped Accounts', workbook.add_format({'bg_color': EY_SARI, 'bottom': 1}))
        worksheet.write_string(row, 3, 'Unmapped',
                               workbook.add_format({'bg_color': EY_SARI, 'bold': 1, 'font_color': 'red', 'bottom': 1}))
        lead_s = row + 1
    else:
        lead_s = row
    # else:
    #     lead_s = row

    row += 2
    col = 4

    worksheet.write_string(row, 1, 'Acc',
                           workbook.add_format({'bg_color': GRI_BG, 'bold': True, 'left': 1, 'bottom': 1, 'top': 1}))
    worksheet.write_string(row, 2, 'Description', border)
    worksheet.write_blank(row, 3, None, border)

    for v in db.periodss:
        if v is db.periodss[-1]:
            worksheet.write_string(row, col, '%s' % v, workbook.add_format(
                {'bg_color': GRI_BG, 'bold': True, 'bottom': 1, 'top': 1, 'right': 1, 'align': 'right'}))

        worksheet.write_string(row, col, '%s' % v, border)
        col += 1

    row += 1

    hesap_b = row + 1
    # todo:buralardan devam
    for item in db.session.query(db.Hesaplar).filter_by(len=3).order_by(db.Hesaplar.number).all():
        worksheet.write_string(row, 1, item.number, workbook.add_format({'bg_color': EY_SARI, 'left': 1}))
        worksheet.write_string(row, 2, item.name, workbook.add_format({'bg_color': EY_SARI}))
        worksheet.write_string(row, 3, item.lead_code,
                               workbook.add_format({'bg_color': EY_SARI, 'bold': True, 'font_color': 'red'}))

        for period in range(db.len_periods):
            if period is db.len_periods - 1:
                worksheet.write_formula(row, period + 4,
                                        "=SUMIF('TBs'!A:A,B{0},'TBs'!{1}:{1})".format(row + 1, chr(67 + period)),
                                        workbook.add_format(
                                            {'bg_color': EY_SARI, 'num_format': NUM_FORMAT, 'right': 1}))
            else:
                worksheet.write_formula(row, period + 4,
                                        "=SUMIF('TBs'!A:A,B{0},'TBs'!{1}:{1})".format(row + 1, chr(67 + period)),
                                        workbook.add_format({'bg_color': EY_SARI, 'num_format': NUM_FORMAT}))

        row += 1

    for i in range(1, db.len_periods + 4):
        worksheet.write_blank(row, i, None, workbook.add_format({'top': 1}))

    hesap_s = row

    for x in range(lead_b, lead_s):
        if x < hesap_b:
            for col in range(db.len_periods):
                if col is db.len_periods - 1 and x == lead_s - 1:
                    fformat = workbook.add_format(
                        {'num_format': NUM_FORMAT, 'bg_color': EY_SARI, 'bottom': 1, 'right': 1})
                elif col is db.len_periods - 1:
                    fformat = workbook.add_format(
                        {'num_format': NUM_FORMAT, 'bg_color': EY_SARI, 'right': 1})
                elif x == lead_s - 1:
                    fformat = workbook.add_format(
                        {'num_format': NUM_FORMAT, 'bg_color': EY_SARI, 'bottom': 1})
                else:
                    fformat = workbook.add_format(
                        {'num_format': NUM_FORMAT, 'bg_color': EY_SARI})

                worksheet.write_formula(x, col + 4,
                                        '=SUMIF(D{1}:D{2},D{3},{0}{1}:{0}{2})'.format(chr(65 + 4 + col), hesap_b,
                                                                                      hesap_s,
                                                                                      x + 1), fformat)


def create_lead(hesap):
    worksheet = workbook.add_worksheet('%s Lead' % hesap)
    worksheet.set_zoom(W_ZOOM)
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
    col = 3
    last_index = db.len_periods + 4

    worksheet.write_string(row, 1, 'Acc.', border)
    worksheet.write_string(row, 2, 'Description', border)

    for v in db.periodss:
        worksheet.write_string(row, col, '%s' % v, border)
        col += 1

    worksheet.write_string(row, col, 'WP Ref', workbook.add_format(
        {'top': 1, 'bottom': 1, 'bold': True, 'bg_color': GRI_BG, 'color': 'red'}))

    if db.len_periods > 1:
        worksheet.write_string(row, col + 1, '% change', border)

    for item in db.session.query(db.Hesaplar).filter_by(len=3, lead_code=hesap).order_by(db.Hesaplar.number).all():
        row += 1
        worksheet.write_string(row, 1, '%s' % item.number)
        worksheet.write_string(row, 2, '%s' % item.name)
        for period in db.periodss:
            worksheet.write_number(row, db.periodss.index(period) + 3, getattr(item, period), money)

        if db.len_periods > 2:
            if db.periodss[-1][3:5] != db.periodss[-2][3:5] and hesap in PL_HESAPLAR:
                # if hesap in PL_HESAPLAR and db.periodss[-1][3:5] != "12":
                worksheet.write_formula(row, last_index,
                                        '=IF({1}{0}=0,"INF",IF({2}{0}=0,"-100",({2}{0}-{1}{0})/{2}{0}))'.
                                        format(row + 1, chr(65 + last_index - 4), chr(65 + last_index - 2)), percn)
            else:
                worksheet.write_formula(row, last_index,
                                        '=IF({1}{0}=0,"INF",IF({2}{0}=0,"-100",({2}{0}-{1}{0})/{2}{0}))'.
                                        format(row + 1, chr(65 + last_index - 3), chr(65 + last_index - 2)), percn)
        elif db.len_periods == 2:
            worksheet.write_formula(row, last_index,
                                    '=IF({1}{0}=0,"INF",IF({2}{0}=0,"-100",({2}{0}-{1}{0})/{2}{0}))'.
                                    format(row + 1, chr(65 + last_index - 3), chr(65 + last_index - 2)), percn)

    row += 2

    worksheet.write_blank(row, 1, None, dipler)
    worksheet.write_string(row, 2, 'TOTAL', dipler)

    for i in range(3, db.len_periods + 3):
        worksheet.write_formula(row, i, '=SUM({0}8:{0}{1})'.format(chr(65 + i), row), dipler)

    worksheet.write_blank(row, db.len_periods + 3, None, dipler)
    worksheet.write_blank(row, db.len_periods + 4, None, dipler)

    row += 2

    worksheet.write_string(row, 1, hesap,
                           workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1, 'left': 1}))
    worksheet.write_string(row, 2, 'Check',
                           workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1}))

    for i in range(3, db.len_periods + 3):
        worksheet.write_formula(row, i,
                                "=VLOOKUP($B{2},'TB Mapping'!$D:${0},{1},0)-{4}{3}".format(chr(65 + last_index - 1),
                                                                                           i - 1,
                                                                                           row + 1,
                                                                                           row - 1, chr(65 + i)),
                                workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1}))

    worksheet.write_blank(row, db.len_periods + 3, None,
                          workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1}))
    worksheet.write_blank(row, db.len_periods + 4, None,
                          workbook.add_format({'color': 'red', 'bold': True, 'top': 1, 'bottom': 1, 'right': 1}))

    create_headline(worksheet, hesap, last_index + 1)


def create_breakdown(hesap):
    worksheet = workbook.add_worksheet('%s1 - BD' % hesap)
    worksheet.set_zoom(W_ZOOM)
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

    row = 6
    col = 1
    start = 8

    tip = db.tanimlar['optional'] if not db.tanimlar['optional'] == 'NO' else 'NO'

    if tip is not "NO":
        last_index = len(db.periodss) + 4
    else:
        last_index = len(db.periodss) + 3

    # todo: bi fonksiyon yaz, formatli bir sekilde seri olarak excele deger girsin. orjinalinde var ama formatlar ayni oluyor.

    create_headline(worksheet, en=last_index + 2)
    for item in db.session.query(db.Hesaplar).filter_by(lead_code=hesap, len=3).order_by(db.Hesaplar.number).all():
        worksheet.write_string(row, col, 'Acc.', header)

        if tip is not 'NO':
            worksheet.write_string(row, col + 1, tip, header)
            col += 1

        worksheet.write_string(row, col + 1, 'Descripton', header)
        col += 2

        for v in db.periodss:
            worksheet.write_string(row, col, '%s' % v, header)
            col += 1

        if db.len_periods > 1:
            worksheet.write_string(row, col, 'Change', header)
            worksheet.write_string(row, col + 1, 'Flux', header)

        col = 1

        list_of_bds = db.session.query(db.Hesaplar).filter_by(ana_hesap=item.number, bd=True).order_by(
            desc(getattr(db.Hesaplar, db.periodss[-1]))).all() if (
            item.lead_code not in {"M", "N", "O", "P", "T", "UA"} or (
                # todo opsiyon eklemek lazim, database'e eklenebilir reversable diye
                item.ana_hesap in {"302", "402", "322", "422", "437", "337", "610", "611", "612", "653", "642",
                                   "654", "655", "659", "680", "681", "682", "689", "656", "657", "660",
                                   "661"} and item.lead_code in {
                    "UB", "VE", "M", "N", "UA"})) else db.session.query(
            db.Hesaplar).filter_by(ana_hesap=item.number, bd=True).order_by(
            asc(getattr(db.Hesaplar, db.periodss[-1]))).all()

        for eben in list_of_bds:
            row += 1
            col = 1
            worksheet.write_string(row, col, eben.number)

            if tip is not 'NO':
                col += 1
                if eben.optional:
                    try:
                        worksheet.write_string(row, col, eben.optional)
                    except:
                        worksheet.write_blank(row, col, None)

            col += 1
            worksheet.write_string(row, col, eben.name)

            for period in db.periodss:
                if tip is not "NO":
                    worksheet.write_number(row, db.periodss.index(period) + 4, getattr(eben, period), money)
                else:
                    worksheet.write_number(row, db.periodss.index(period) + 3, getattr(eben, period), money)

                if db.len_periods > 2:
                    if db.periodss[-1][3:5] != db.periodss[-2][3:5] and hesap in PL_HESAPLAR:
                        # if hesap in PL_HESAPLAR and db.periodss[-1][3:5] != "12":
                        worksheet.write_formula(row, last_index,
                                                '={1}{0}-{2}{0}'.format(row + 1, chr(65 + last_index - 1),
                                                                        chr(65 + last_index - 3)), money)
                        worksheet.write_formula(row, last_index + 1,
                                                '=IF({1}{0}=0,"INF",IF({2}{0}=0,"-100",({2}{0}-{1}{0})/{2}{0}))'.
                                                format(row + 1, chr(65 + last_index - 3), chr(65 + last_index - 1)),
                                                percn)
                    else:
                        worksheet.write_formula(row, last_index,
                                                '={1}{0}-{2}{0}'.format(row + 1, chr(65 + last_index - 1),
                                                                        chr(65 + last_index - 2)), money)
                        worksheet.write_formula(row, last_index + 1,
                                                '=IF({1}{0}=0,"INF",IF({2}{0}=0,"-100",({2}{0}-{1}{0})/{2}{0}))'.
                                                format(row + 1, chr(65 + last_index - 2), chr(65 + last_index - 1)),
                                                percn)
                elif db.len_periods == 2:
                    worksheet.write_formula(row, last_index,
                                            '={1}{0}-{2}{0}'.format(row + 1, chr(65 + last_index - 2),
                                                                    chr(65 + last_index - 3)), money)
                    worksheet.write_formula(row, last_index + 1,
                                            '=IF({1}{0}=0,"INF",IF({2}{0}=0,"-100",({2}{0}-{1}{0})/{2}{0}))'.
                                            format(row + 1, chr(65 + last_index - 3), chr(65 + last_index - 2)),
                                            percn)

                    # if hesap in PL_HESAPLAR and db.periodss[-1][3:5] != "12":
                    #     worksheet.write_formula(row, 6, '=F%s-D%s' % (row + 1, row + 1), money)
                    #     worksheet.write_formula(row, 7, '=IF(D%s=0,"INF",IF(F%s=0,"-100",(F%s-D%s)/F%s))' % (
                    #         row + 1, row + 1, row + 1, row + 1, row + 1), percn)
                    # else:
                    #     worksheet.write_formula(row, 6, '=F%s-E%s' % (row + 1, row + 1), money)
                    #     worksheet.write_formula(row, 7, '=IF(E%s=0,"INF",IF(F%s=0,"-100",(F%s-E%s)/F%s))' % (
                    #         row + 1, row + 1, row + 1, row + 1, row + 1), percn)

        row += 2

        if tip is not 'NO':
            ran = db.len_periods + 6
        else:
            ran = db.len_periods + 5

        if tip is 'NO':
            for i in range(1, ran):
                if i == 2:
                    worksheet.write_string(row, i, 'Total', dipler)
                    worksheet.write_string(row + 2, i, 'Check', dipler)

                elif i == 1 or i > db.len_periods + 2:
                    worksheet.write_blank(row, i, None, dipler)
                    worksheet.write_blank(row + 2, i, None, dipler)
                else:
                    worksheet.write_formula(row, i, '=SUM({0}{1}:{0}{2})'.format(chr(65 + i), start, row - 1), dipler)
                    worksheet.write_string(row + 1, i, '|--- %s Lead ---|' % hesap, color)
                    worksheet.write_formula(row + 2, i,
                                            "=IFERROR(VLOOKUP(\"{0}\",'{1} Lead'!B:{2},{5},0),0)-{4}{3}".format(
                                                item.number, hesap, chr(65 + db.len_periods + 2), row + 1, chr(65 + i),
                                                i),
                                            dipler)
        else:
            for i in range(1, ran):
                if i == 3:
                    worksheet.write_string(row, i, 'Total', dipler)
                    worksheet.write_string(row + 2, i, 'Check', dipler)

                elif i in (1, 2) or i > db.len_periods + 3:
                    worksheet.write_blank(row, i, None, dipler)
                    worksheet.write_blank(row + 2, i, None, dipler)
                else:
                    worksheet.write_formula(row, i, '=SUM({0}{1}:{0}{2})'.format(chr(65 + i), start, row - 1), dipler)
                    worksheet.write_string(row + 1, i, '|--- %s Lead ---|' % hesap, color)
                    worksheet.write_formula(row + 2, i,
                                            "=IFERROR(VLOOKUP(\"{0}\",'{1} Lead'!B:{2},{5},0),0)-{4}{3}".format(
                                                item.number, hesap, chr(65 + db.len_periods + 2), row + 1, chr(65 + i),
                                                                                                  i - 1),
                                            dipler)

        row += 6
        start = row + 2
        col = 1


def create_leads():
    for k in db.session.query(db.Lead).group_by(db.Lead.lead_code).order_by(db.Lead.lead_code).all():
        if k.lead_code == 'Unmapped':
            continue
        if not db.session.query(db.Hesaplar).filter_by(lead_code=k.lead_code).first():
            continue
        # todo: alltbs tepedeki kucuk tablo daha dynamic hale gelmeli
        create_lead(k.lead_code)
        create_breakdown(k.lead_code)

    if len(db.session.query(db.Lead).filter_by(lead_code='Unmapped').all()) > 0:
        create_lead('Unmapped')
        create_breakdown('Unmapped')


def create_a4():
    # bitmedi daha da neyin bitmedigini unuttum amk
    file_name = 'a4'

    if db.tanimlar['company']:
        file_name = '{0} - {1}'.format(db.tanimlar['company'].encode('utf-8'), db.periodss[-1])

    create_workbook(file_name)
    create_tbs()
    # create_comperative()
    create_alltb()
    create_leads()
    workbook.close()


if __name__ == '__main__':
    def define_variables():
        from xlrd import open_workbook
        excel_file = open_workbook('a3m.xlsx', encoding_override='utf-8')
        sheet = excel_file.sheet_by_name('Instruction')

        try:
            db.tanimlar['optional'] = sheet.cell_value(7, 2)
        except:
            pass

        db.tanimlar['company'] = sheet.cell(6, 2).value

        for sheet in excel_file.sheets()[2:]:
            db.periodss.append(sheet.name)
        db.Hesaplar = db.make_hesaplar()
        db.create_db()
        db.len_periods = len(db.periodss)


    define_variables()
    create_a4()
