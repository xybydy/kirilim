# -*- coding: utf-8 -*-

__author__ = 'fatihka'

from sqlalchemy import func
from xlrd import open_workbook

import db
from utils import flush


# home = os.getenv('HOME')
# desktop = os.path.join(home, 'Desktop')
# file = os.path.join(desktop, 'a3.xlsx')


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

        db.tanimlar['company'] = sheet.cell(6, 2).value

        for sheet in excel_file.sheets()[2:]:
            db.periodss.append(sheet.name)

    donus = dict()
    define_variables()
    db.Hesaplar = db.make_hesaplar()
    db.create_db()
    prepare_mapping()

    for i in range(2, excel_file.nsheets):
        sheet = excel_file.sheet_by_index(i)
        key = sheet.name

        flush('Parsing %s...' % key)

        for row in range(1, sheet.nrows):
            temp = []

            hucre = str(int(sheet.cell_value(row, 0))) if type(sheet.cell_value(row, 0)) is float else \
                sheet.cell_value(row, 0)

            if hucre not in donus: donus[hucre] = dict()

            for col in range(sheet.ncols):
                if col is 0 and type(sheet.cell_value(row, col)) is float:
                    temp.append(str(int(sheet.cell_value(row, col))))  # burasi duzeltilecek bug var.
                else:
                    temp.append(sheet.cell_value(row, col))

            gecici = db.session.query(db.Hesaplar).filter_by(number=hucre).first()
            ana_hesap = hucre[:3]
            lead_cod = None
            if db.session.query(db.Lead).filter_by(account=ana_hesap).first() is not None:
                lead_cod = db.session.query(db.Lead).filter_by(account=ana_hesap).first().lead_code

            if gecici is not None:
                setattr(gecici, key, temp[4])
                flush(hucre + ' updated...', wait=0.0001, code='blue')
            else:
                db.session.add(
                    db.Hesaplar(number=hucre, name=temp[1], len=len(hucre), ana_hesap=ana_hesap, lead_code=lead_cod,
                                **{key: temp[4]}))
                flush(hucre + ' added...', wait=0.0001, code='blue')

        db.session.commit()
        flush('Period %s parsed...' % key)


def find_bds():
    flush('Determining breakdowns...')
    accounts = db.session.query(db.Hesaplar).all()
    for k in accounts:
        if len(db.session.query(db.Hesaplar).filter(db.Hesaplar.number.startswith(k.number)).all()) <= 1:
            k.bd = True
    db.session.commit()
    flush('Breakdowns are set!')


def summary_check():
    if db.session.query(db.Hesaplar).filter_by(len=3).first() is not None:
        flush('Data has main accounts.')
        return True
    flush('There is no main accounts!', err=True)


def create_summary_accs():
    flush('Creating main accounts...')

    # sums = ', '.join(['SUM("{0}") AS "{0}"'.format(k) for k in db.periodss])

    # raw_query = text("SELECT ana_hesap, lead_code, %s from hesaplar WHERE bd=1 GROUP BY ana_hesap" % sums)

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


def delete_zeros(exceptions=['900']):
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
