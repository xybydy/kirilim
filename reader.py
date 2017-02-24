# -*- coding: utf-8 -*-
__author__ = 'fatihka'
from sqlalchemy import func, and_
from xlrd import open_workbook

import db
from utils import flush

# home = os.getenv('HOME')
# desktop = os.path.join(home, 'Desktop')
# file = os.path.join(desktop, 'a3.xlsx')

bakiye = ["bakiye", "balance"]
hesap_adi = ['description', 'hesap adi', 'hesap adı', 'hesap', 'aciklama', 'açıklama']
sum_var = False


def parse_excel_file(file):
    excel_file = open_workbook(file, encoding_override='utf-8')
    flush('Opening excel file...')

    def prepare_mapping():
        s = excel_file.sheet_by_name("Mapping")

        for x in range(1, s.nrows):
            db.session.add(
                db.Lead(account=s.cell(x, 0).value, account_name=s.cell(x, 1).value, lead_code=s.cell(x, 2).value,
                        name=s.cell(x, 3).value))
        flush('Parsing mappings...')
        db.session.commit()

    def define_variables():
        flush('Defining Company names and audit periods...')
        sheet = excel_file.sheet_by_name('Instruction')

        db.tanimlar['company'] = sheet.cell_value(6, 2)
        try:
            db.tanimlar['optional'] = sheet.cell_value(7, 2) or "NO"
        except IndexError:
            flush('There is no optional column...')

        for sheet in excel_file.sheets()[2:]:
            db.periodss.append(sheet.name)

    # donus = dict()
    define_variables()
    db.Hesaplar = db.make_hesaplar()
    db.create_db()
    db.len_periods = len(db.periodss)
    prepare_mapping()

    def find_seperator():
        from collections import defaultdict
        pointer = defaultdict(int)
        stop = 0
        for item in excel_file.sheet_by_index(excel_file.nsheets - 1).col(0):
            if stop == 5:
                return max(pointer.keys(), key=lambda q: pointer[q])
            if not str(item.value).isalnum():
                for char in str(item.value):
                    if not char.isalnum():
                        pointer[char] += 1
                        stop += 1

    for i in range(2, excel_file.nsheets):
        sheet = excel_file.sheet_by_index(i)
        key = sheet.name

        flush('Parsing %s...' % key)

        headers = [col.value.strip() for col in sheet.row(0)]

        desc_col = 0
        bal_col = 0
        opt_col = 999

        for y in range(len(headers)):
            if headers[y].lower() in bakiye:
                bal_col = y
            if headers[y].lower() in hesap_adi:
                desc_col = y
            if headers[y] == db.tanimlar['optional']:
                opt_col = y

        for row in range(1, sheet.nrows):
            temp = {}
            # todo float diye gorulen kalemlerde seperator farkedilip aslinda str'ye cevrilip seperator replace edilebilir.

            for col in range(sheet.ncols):
                if col is 0:
                    temp["number"] = str(sheet.cell_value(row, col))  # burasi duzeltilecek bug var.
                elif col == bal_col:
                    temp[key] = sheet.cell_value(row, col)
                elif col == opt_col:
                    temp['optional'] = sheet.cell_value(row, col)
                elif col == desc_col:
                    temp['name'] = sheet.cell_value(row, col)

            param = 'optional' if not db.tanimlar['optional'] == 'NO' else 'NO'

            if param is 'NO':
                gecici = db.session.query(db.Hesaplar).filter_by(number=temp['number']).first()
            else:
                hink = db.session.query(db.Hesaplar).filter_by(number=temp['number'],
                                                               **{param: temp[param]}).first()
                gecici = None
                if hink is not None:
                    if int(getattr(hink, key)) is 0:
                        gecici = hink

            ana_hesap = temp['number'][:3]
            lead_cod = None

            if db.session.query(db.Lead).filter_by(account=ana_hesap).first() is not None:
                lead_cod = db.session.query(db.Lead).filter_by(account=ana_hesap).first().lead_code

            if gecici is not None:
                setattr(gecici, key, temp[key])
                flush(temp['number'] + ' updated...', wait=0.0001, code='blue')
            else:
                db.session.add(db.Hesaplar(**temp, ana_hesap=ana_hesap, lead_code=lead_cod, len=len(temp['number'])))
                flush(temp['number'] + ' added...', wait=0.0001, code='blue')

        db.session.commit()
        flush('Period %s parsed...' % key)


def find_bds():
    flush('Determining breakdowns...')
    accounts = db.session.query(db.Hesaplar).all()
    for k in accounts:
        if db.tanimlar['optional'] != 'NO' and sum_var is False:
            if len(k.number) > 3:
                k.bd = True
        elif db.tanimlar['optional'] != 'NO' and sum_var is True:
            if len(db.session.query(db.Hesaplar).filter(and_(db.Hesaplar.number.startswith(k.number),
                                                             db.Hesaplar.optional == k.optional)).all()) <= 1 \
                    and len(k.number) > 3:
                k.bd = True
        else:
            if len(db.session.query(db.Hesaplar).filter(db.Hesaplar.number.startswith(k.number)).all()) <= 1:
                k.bd = True

    db.session.commit()
    flush('Breakdowns are set!')


def summary_check():
    if db.session.query(db.Hesaplar).filter_by(len=3).first() is not None:
        flush('Data has main accounts.')
        sum_var = True
        return True
    flush('There is no main accounts!', err=True)


def create_summary_accs():
    flush('Creating main accounts...')

    query = db.session.query(db.Hesaplar.ana_hesap,
                             db.Hesaplar.lead_code,
                             *[func.sum(getattr(db.Hesaplar, '{}'.format(period))).label('{}'.format(period)) for period
                               in db.periodss]).group_by('ana_hesap').all()

    for k in query:
        unmapped = None

        if k.lead_code == 'Unmapped':
            unmapped = db.session.query(db.Hesaplar).filter_by(len=3,
                                                               ana_hesap=k.ana_hesap).first() or db.session.query(
                db.Hesaplar).filter_by(ana_hesap=k.ana_hesap).first()

        source = db.session.query(db.Lead).filter_by(account=k.ana_hesap).first()

        main_source = unmapped or source
        name = main_source.name

        t = db.Hesaplar(number=k.ana_hesap, ana_hesap=k.ana_hesap, name=name, lead_code=main_source.lead_code, len=3)

        for q in db.periodss:
            [setattr(t, q, getattr(k, q))]
            db.session.add(t)
            db.session.commit()


def fix_mainaccs():
    flush('Changing the names of main accounts...')
    hesaplar = db.session.query(db.Hesaplar).filter_by(len=3)
    for k in hesaplar.all():
        item = db.session.query(db.Lead).filter_by(account=k.ana_hesap).first()
        if item:
            k.name = item.account_name

    db.session.commit()


def delete_zeros(exceptions=None):
    if exceptions is None:
        exceptions = ['900']

    flush('Deleting accounts with zero balances in all periods...')
    query = db.session.query(db.Hesaplar).filter_by(**{k: 0 for k in db.periodss}).all()
    for item in query:
        if item.number[:3] in exceptions:
            continue
        db.session.delete(item)

    db.session.commit()


def create_or_parse_sum():
    if summary_check():
        return True
    else:
        create_summary_accs()


# db.periodss = ['31.10.2015', '31.12.2015', '31.12.2016']
# db.create_db()
# db.Hesaplar = db.make_hesaplar()
# create_summary_accs()

if __name__ == '__main__':
    parse_excel_file("a3m.xlsx")
    delete_zeros()
    create_or_parse_sum()
    fix_mainaccs()
    find_bds()
