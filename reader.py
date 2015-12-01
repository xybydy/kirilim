# -*- coding: utf-8 -*-


from sqlalchemy import func
from xlrd import open_workbook

import db

__author__ = 'fatihka'

# home = os.getenv('HOME')
# desktop = os.path.join(home, 'Desktop')
# file = os.path.join(desktop, 'a3.xlsx')

verbose = True

SINIFLANDIRMA = {u'C': u'Cash and cash equivalents', u'D': u'Marketable securities', u'E': u'Trade receivables',
                 u'F': u'Inventories', u'G': u'Prepaid expenses and other current assets', u'H': u'Investments',
                 u'J': u'Other non-current assets',
                 u'K': u'Property plant and equipment & Intangible assets',
                 u'I': u'Related party balances', u'M': u'Short-term borrowings',
                 u'N': u'Trade payables', u'P': u'Other current payables and accrued liabilities',
                 u'O': u'Income taxes payable and Tax charge',
                 u'R': u'Other non-current liabilities', u'T': u"Shareholder's equity", u'UA': u'Sales',
                 u'VA': u'Cost of sales', u'VC': u'Operating expenses', u'VE': u'Financial income and expense',
                 u'UB': u'Other income and expenses', u'X': u'Suspense accounts', u'Unmapped': u'Unmapped Accounts'}

hesap_kodlari = {u'O': [u'371', u'370', u'691'], u'D': [u'111', u'110', u'112', u'119', u'118'],
                 u'VA': [u'623', u'622', u'621', u'620'], u'C': [u'102', u'100', u'108'],
                 u'VC': [u'630', u'631', u'632'],
                 u'E': [u'120', u'121', u'122', u'123', u'124', u'127', u'128', u'129', u'101', u'222', u'221', u'220'],
                 u'VE': [u'661', u'660', u'642', u'646', u'656', u'657', u'647'],
                 u'G': [u'137', u'136', u'135', u'139', u'138', u'126', u'199', u'198', u'195', u'194', u'197', u'196',
                        u'191', u'190', u'193', u'192', u'179', u'178', u'170', u'180', u'181'],
                 u'F': [u'151', u'150', u'153', u'152', u'155', u'154', u'157', u'156', u'159', u'158'],
                 u'H': [u'245', u'244', u'247', u'246', u'241', u'240', u'243', u'242', u'249', u'248'],
                 u'J': [u'281', u'280', u'298', u'299', u'297', u'294', u'295', u'292', u'293', u'291', u'271', u'272',
                        u'278', u'279', u'229', u'226', u'224', u'277', u'239', u'235', u'236', u'237'],
                 u'M': [u'407', u'405', u'402', u'401', u'400', u'409', u'408', u'308', u'309', u'300', u'301', u'302',
                        u'303', u'304', u'305', u'306'],
                 u'N': [u'426', u'422', u'429', u'421', u'420', u'103', u'322', u'320', u'321', u'326', u'329'],
                 u'P': [u'340', u'349', u'372', u'379', u'472', u'393', u'399', u'368', u'369', u'360', u'361', u'449',
                        u'440', u'380',
                        u'381', u'339', u'335', u'337', u'336', u'373', u'392', u'391', u'397', u'350', u'358'],
                 u'R': [u'492', u'493', u'438', u'436', u'437', u'499', u'481', u'480', u'479'],
                 u'I': [u'133', u'132', u'131', u'130', u'231', u'232', u'233', u'432', u'433', u'431', u'331', u'333',
                        u'332'],
                 u'X': [u'762', u'760', u'740', u'713', u'711', u'710', u'770', u'752', u'733', u'780', u'781', u'782',
                        u'741', u'742', u'730', u'761', u'750', u'722', u'723', u'720', u'721', u'734', u'751', u'731',
                        u'771', u'772', u'712', u'732'],
                 u'K': [u'265', u'264', u'269', u'268', u'263', u'262', u'261', u'260', u'267',
                        u'255', u'258', u'259', u'252', u'253', u'250', u'251', u'256', u'257', u'254'],
                 u'UB': [u'697', u'698', u'671', u'682', u'641', u'640', u'643', u'645', u'644', u'653', u'659', u'654',
                         u'679', u'649', u'655', u'689', u'681', u'680'],
                 u'UA': [u'601', u'602', u'600', u'612', u'610', u'611'],

                 u'T': [u'541', u'548', u'549', u'570', u'522', u'529', u'591', u'590', u'525', u'520', u'521', u'523',
                        u'580', u'542', u'540', u'502', u'500', u'501'], u'Unmapped': [u'900']}

ACCOUNT_NAMES = {u'265': {'Name': u'PP&E Acquired Through Financial Leases', 'Lead': u'K'},
                 u'762': {'Name': u'Selling & Marketing Expenses (Variances)', 'Lead': u'X'},
                 u'760': {'Name': u'Selling & Marketing Expenses', 'Lead': u'X'},
                 u'127': {'Name': u'Other Trade Receivables', 'Lead': u'E'},
                 u'661': {'Name': u'Long-Term Debt Expense', 'Lead': u'VE'},
                 u'660': {'Name': u'Short-Term Debt Expense', 'Lead': u'VE'},
                 u'133': {'Name': u'Due From Subsidiaries', 'Lead': u'I'},
                 u'132': {'Name': u'Due From Participations', 'Lead': u'I'},
                 u'131': {'Name': u'Due From Shareholders', 'Lead': u'I'},
                 u'130': {'Name': u'Due From Intergroup Companies', 'Lead': u'I'},
                 u'137': {'Name': u'Rediscount Of Other Notes Receivables', 'Lead': u'G'},
                 u'136': {'Name': u'Other Miscellaneous Receivables', 'Lead': u'G'},
                 u'135': {'Name': u'Due From Personnel', 'Lead': u'G'},
                 u'139': {'Name': u'Allowances For Other Doubtful Receivables', 'Lead': u'G'},
                 u'138': {'Name': u'Other Doubtful Receivables', 'Lead': u'G'},
                 u'492': {'Name': u'Vat Deferred To Following Years', 'Lead': u'R'},
                 u'493': {'Name': u'Installation Participation', 'Lead': u'R'},
                 u'691': {'Name': u'Tax Charge', 'Lead': u'O'},
                 u'697': {'Name': u'Inflation Adjustment of Construction Contracts', 'Lead': u'UB'},
                 u'740': {'Name': u'Cost of Services Rendered ', 'Lead': u'X'},
                 u'698': {'Name': u'Inflation Adjustment ', 'Lead': u'UB'},
                 u'407': {'Name': u'Other Marketable Securities Issued', 'Lead': u'M'},
                 u'405': {'Name': u'Bonds Issued', 'Lead': u'M'}, u'541': {'Name': u'Special Reserves', 'Lead': u'T'},
                 u'340': {'Name': u'Advances Taken', 'Lead': u'P'},
                 u'402': {'Name': u'Rediscount on Financial Leasing Payables', 'Lead': u'M'},
                 u'401': {'Name': u'Financial Leasing Payables', 'Lead': u'M'},
                 u'400': {'Name': u'Bank Loans', 'Lead': u'M'},
                 u'281': {'Name': u'Long-Term Accrued Income', 'Lead': u'J'},
                 u'280': {'Name': u'Long-Term Prepaid Expenses', 'Lead': u'J'},
                 u'548': {'Name': u'Other Profit Reserves', 'Lead': u'T'},
                 u'549': {'Name': u'Special Funds', 'Lead': u'T'},
                 u'349': {'Name': u'Other Advances Received', 'Lead': u'P'},
                 u'409': {'Name': u'Other Financial Payables', 'Lead': u'M'},
                 u'408': {'Name': u'Premiums On Marketable Securities Issued', 'Lead': u'M'},
                 u'570': {'Name': u'Retained Earnings', 'Lead': u'T'},
                 u'522': {'Name': u'Revaluation Fund -F/A', 'Lead': u'T'},
                 u'523': {'Name': u'Revaluation Fund - Part.', 'Lead': u'T'},
                 u'713': {'Name': u'Direct Materil (Quantity Variances)', 'Lead': u'X'},
                 u'712': {'Name': u'Direct Materil (Price Variances)', 'Lead': u'X'},
                 u'711': {'Name': u'Direct Material (Reflected)', 'Lead': u'X'},
                 u'710': {'Name': u'Direct Material', 'Lead': u'X'},
                 u'120': {'Name': u'Trade Receivables', 'Lead': u'E'},
                 u'121': {'Name': u'Notes Receivable', 'Lead': u'E'},
                 u'122': {'Name': u'Rediscount On Notes Receivable', 'Lead': u'E'},
                 u'123': {'Name': u'Cheques Receivable (Maturity)', 'Lead': u'E'},
                 u'124': {'Name': u'Unearned Leasing Interest Income', 'Lead': u'E'},
                 u'126': {'Name': u'Deposits & Guarantees Given', 'Lead': u'G'},
                 u'264': {'Name': u'Leasehold Improvements', 'Lead': u'K'},
                 u'128': {'Name': u'Doubtful Trade Rec.', 'Lead': u'E'},
                 u'129': {'Name': u'Allowances For Doubtful Trade Receivables', 'Lead': u'E'},
                 u'269': {'Name': u'Advances Given', 'Lead': u'K'},
                 u'268': {'Name': u'Accumulated Amortization', 'Lead': u'K'},
                 u'422': {'Name': u'Rediscount On Notes Payable', 'Lead': u'N'},
                 u'682': {'Name': u'KKEG', 'Lead': u'UB'},
                 u'379': {'Name': u'Other Allowances', 'Lead': u'P'},
                 u'298': {'Name': u'Allowances For The Decline In The Value Of Inv.', 'Lead': u'J'},
                 u'299': {'Name': u'Accumulated Depreciation', 'Lead': u'J'},
                 u'371': {'Name': u'Prepaid Tax ', 'Lead': u'O'},
                 u'297': {'Name': u'Other Miscellaneous Assets', 'Lead': u'J'},
                 u'294': {'Name': u'Fixed Assets & Inventories To Be Sold', 'Lead': u'J'},
                 u'295': {'Name': u'Prepaid Taxes & Funds', 'Lead': u'J'}, u'292': {'Name': u'Other Vat', 'Lead': u'J'},
                 u'293': {'Name': u'Long-Term Inventories', 'Lead': u'J'},
                 u'291': {'Name': u'Long-Term Vat Deductible', 'Lead': u'J'},
                 u'591': {'Name': u'Net Loss For The Current Period', 'Lead': u'T'},
                 u'590': {'Name': u'Net Profit For The Current Period', 'Lead': u'T'},
                 u'199': {'Name': u'Allowances For Other Current Assets', 'Lead': u'G'},
                 u'198': {'Name': u'Other Miscellaneous Current Assets', 'Lead': u'G'},
                 u'195': {'Name': u'Job Advances', 'Lead': u'G'}, u'194': {'Name': u'SCT Deductible', 'Lead': u'G'},
                 u'197': {'Name': u'Count & Delivery Shortages', 'Lead': u'G'},
                 u'196': {'Name': u'Advances Given To Personnel', 'Lead': u'G'},
                 u'191': {'Name': u'Vat Deductible', 'Lead': u'G'},
                 u'190': {'Name': u'Vat To Be Transferred', 'Lead': u'G'},
                 u'193': {'Name': u'Prepaid Taxes & Funds', 'Lead': u'G'}, u'192': {'Name': u'Other Vat', 'Lead': u'G'},
                 u'393': {'Name': u'Head-Office & Branch Current Accounts', 'Lead': u'P'},
                 u'271': {'Name': u'Research Expenses', 'Lead': u'J'},
                 u'272': {'Name': u'Preparation & Development Expenses', 'Lead': u'J'},
                 u'111': {'Name': u'Private Sector Bonds', 'Lead': u'D'},
                 u'110': {'Name': u'Share Certificates', 'Lead': u'D'},
                 u'112': {'Name': u'Public Sector Bonds', 'Lead': u'D'},
                 u'278': {'Name': u'Accumulated Depletion', 'Lead': u'J'},
                 u'279': {'Name': u'Advances Given', 'Lead': u'J'},
                 u'399': {'Name': u'Other Miscellaneous Liabilities', 'Lead': u'P'},
                 u'119': {'Name': u'Allow. For The Decline In The Value of Securities', 'Lead': u'D'},
                 u'118': {'Name': u'Other Securities', 'Lead': u'D'},
                 u'752': {'Name': u'Research and Development (Variances)', 'Lead': u'X'},
                 u'647': {'Name': u'Rediscount Income', 'Lead': u'VE'},
                 u'429': {'Name': u'Othertrade Payable', 'Lead': u'N'},
                 u'611': {'Name': u'Sales Discounts', 'Lead': u'UA'},
                 u'520': {'Name': u'Premiums On Sale Of Share Certificates', 'Lead': u'T'},
                 u'521': {'Name': u'Gain On Canceled Share Capital', 'Lead': u'T'},
                 u'368': {'Name': u'Rescheduled Taxes', 'Lead': u'P'},
                 u'369': {'Name': u'Other Liabilities', 'Lead': u'P'},
                 u'421': {'Name': u'Notes Payable', 'Lead': u'N'}, u'420': {'Name': u'Trade Payable', 'Lead': u'N'},
                 u'255': {'Name': u'Furniture & Fixture', 'Lead': u'K'},
                 u'529': {'Name': u'Other Capital Reserves', 'Lead': u'T'},
                 u'360': {'Name': u'Taxes & Funds Payable', 'Lead': u'P'},
                 u'426': {'Name': u'Deposits & Guarantees Received', 'Lead': u'N'},
                 u'308': {'Name': u'Premiums On Marketable Securities Issued', 'Lead': u'M'},
                 u'309': {'Name': u'Other Financial Payables', 'Lead': u'M'},
                 u'449': {'Name': u'Other Advances Received', 'Lead': u'P'},
                 u'440': {'Name': u'Advances Received From Customer', 'Lead': u'P'},
                 u'580': {'Name': u'Prior Year Losses', 'Lead': u'T'},
                 u'542': {'Name': u'General Reserves', 'Lead': u'T'},
                 u'300': {'Name': u'Bank Loans', 'Lead': u'M'},
                 u'301': {'Name': u'Financial Leasing Payables', 'Lead': u'M'},
                 u'302': {'Name': u'Rediscount on Financial Leasing Payables', 'Lead': u'M'},
                 u'303': {'Name': u'Current Portion of Long-term Bank Loans', 'Lead': u'M'},
                 u'304': {'Name': u'Bonds Principle & Interest Payable', 'Lead': u'M'},
                 u'305': {'Name': u'Bonds & Notes Issued', 'Lead': u'M'},
                 u'306': {'Name': u'Other Marketable Securities Issued', 'Lead': u'M'},
                 u'245': {'Name': u'Subsidiaries', 'Lead': u'H'},
                 u'244': {'Name': u'Allowance For Participations', 'Lead': u'H'},
                 u'247': {'Name': u'Allowance For  Investments', 'Lead': u'H'},
                 u'246': {'Name': u'S/C Commitment To Subsidiaries', 'Lead': u'H'},
                 u'241': {'Name': u'Allow. For The Decline In The Value Of L/T M/S', 'Lead': u'H'},
                 u'240': {'Name': u'Long-Term Marketable Securities', 'Lead': u'H'},
                 u'243': {'Name': u'S/C Commitment To Participations', 'Lead': u'H'},
                 u'242': {'Name': u'Participations', 'Lead': u'H'}, u'102': {'Name': u'Cash In Banks', 'Lead': u'C'},
                 u'103': {'Name': u'Payment Orders', 'Lead': u'N'}, u'100': {'Name': u'Cash On Hand', 'Lead': u'C'},
                 u'101': {'Name': u'Cheques Receivable', 'Lead': u'E'},
                 u'249': {'Name': u'Allowances For Other Financial Assets', 'Lead': u'H'},
                 u'248': {'Name': u'Other Financial  Assets', 'Lead': u'H'},
                 u'380': {'Name': u'Short-Term Deferred Income', 'Lead': u'P'},
                 u'372': {'Name': u'Allowances For Retirement Pay', 'Lead': u'P'},
                 u'780': {'Name': u'Financial Expenses', 'Lead': u'X'},
                 u'781': {'Name': u'Financial Expenses (Reflected)', 'Lead': u'X'},
                 u'782': {'Name': u'Financial Expenses (Variances)', 'Lead': u'X'},
                 u'381': {'Name': u'Accrued Expenses', 'Lead': u'P'},
                 u'751': {'Name': u'Research and Development (Reflected)', 'Lead': u'X'},
                 u'108': {'Name': u'Other Cash And Cash Equivalents', 'Lead': u'C'},
                 u'641': {'Name': u'Dividend Income From Subsidiaries', 'Lead': u'UB'},
                 u'640': {'Name': u'Dividend Income From Participations', 'Lead': u'UB'},
                 u'643': {'Name': u'Commission Income', 'Lead': u'UB'},
                 u'642': {'Name': u'Interest Income', 'Lead': u'VE'},
                 u'645': {'Name': u'Gain On Sale Of M/S', 'Lead': u'UB'},
                 u'644': {'Name': u'Provisions No Longer Required', 'Lead': u'UB'},
                 u'438': {'Name': u'Rescheduled Payables To Government', 'Lead': u'R'},
                 u'646': {'Name': u'F/X Gain', 'Lead': u'VE'},
                 u'436': {'Name': u'Other Miscellaneous Payables', 'Lead': u'R'},
                 u'437': {'Name': u'Rediscount On Other Notes Payable', 'Lead': u'R'},
                 u'601': {'Name': u'Sales -Export', 'Lead': u'UA'},
                 u'432': {'Name': u'Due To Participations - L/T', 'Lead': u'I'},
                 u'433': {'Name': u'Due To Subsidiaries - L/T', 'Lead': u'I'},
                 u'431': {'Name': u'Due To Shareholders - L/T', 'Lead': u'I'},
                 u'623': {'Name': u'Cost Of Other Sales', 'Lead': u'VA'},
                 u'622': {'Name': u'Cost Of Services Rendered', 'Lead': u'VA'},
                 u'621': {'Name': u'Cost Of Merchandises Sold', 'Lead': u'VA'},
                 u'620': {'Name': u'Cost Of Goods Sold', 'Lead': u'VA'},
                 u'339': {'Name': u'Other Payables', 'Lead': u'P'},
                 u'741': {'Name': u'Cost of Services Rendered (Reflected)', 'Lead': u'X'},
                 u'335': {'Name': u'Due To Personnel', 'Lead': u'P'},
                 u'337': {'Name': u'Rediscount On Other Notes Payable', 'Lead': u'P'},
                 u'336': {'Name': u'Other Miscellaneous Payable', 'Lead': u'P'},
                 u'331': {'Name': u'Due To Shareholders', 'Lead': u'I'},
                 u'742': {'Name': u'Cost of Services Rendered (Variances)', 'Lead': u'X'},
                 u'333': {'Name': u'Due To Participations', 'Lead': u'I'},
                 u'332': {'Name': u'Due To Investments', 'Lead': u'I'},
                 u'258': {'Name': u'Construction In Progress', 'Lead': u'K'},
                 u'259': {'Name': u'Advances Given', 'Lead': u'K'},
                 u'179': {'Name': u'Advances Given To Subcontractors', 'Lead': u'G'},
                 u'178': {'Name': u'Construction Inflation Adjustment Spread Over Years', 'Lead': u'G'},
                 u'252': {'Name': u'Buildings', 'Lead': u'K'},
                 u'253': {'Name': u'Machinery Installation & Equipment', 'Lead': u'K'},
                 u'250': {'Name': u'Land', 'Lead': u'K'}, u'251': {'Name': u'Land Improvements', 'Lead': u'K'},
                 u'256': {'Name': u'Other Tangible Assets', 'Lead': u'K'},
                 u'257': {'Name': u'Accumulated Depreciation', 'Lead': u'K'},
                 u'254': {'Name': u'Vehicles', 'Lead': u'K'},
                 u'170': {'Name': u'Construction & Repair Costs Spread Over Years', 'Lead': u'G'},
                 u'499': {'Name': u'Other Miscellenous Long-Term Liabilities', 'Lead': u'R'},
                 u'657': {'Name': u'Rediscount Expenses', 'Lead': u'VE'},
                 u'602': {'Name': u'Other Revenues', 'Lead': u'UA'},
                 u'540': {'Name': u'Legal Reserves', 'Lead': u'T'},
                 u'761': {'Name': u'Selling & Marketing Expenses (Reflected)', 'Lead': u'X'},
                 u'731': {'Name': u'Manufacturing Overheads (Reflected)', 'Lead': u'X'},
                 u'730': {'Name': u'Manufacturing Overheads', 'Lead': u'X'},
                 u'180': {'Name': u'Short-Term Prepaid Expenses', 'Lead': u'G'},
                 u'181': {'Name': u'Accrued Income', 'Lead': u'G'},
                 u'734': {'Name': u'Manufacturing Overheads (Capacity Variances)', 'Lead': u'X'},
                 u'502': {'Name': u'Positive Inflation Adjustment Differences', 'Lead': u'T'},
                 u'500': {'Name': u'Share Capital', 'Lead': u'T'},
                 u'501': {'Name': u'Unpaid Share Capital', 'Lead': u'T'},
                 u'630': {'Name': u'Research & Development Expenses', 'Lead': u'VC'},
                 u'631': {'Name': u'Selling & Marketing Expenses', 'Lead': u'VC'},
                 u'632': {'Name': u'General & Admin Expenses', 'Lead': u'VC'},
                 u'750': {'Name': u'Research and Development', 'Lead': u'X'},
                 u'370': {'Name': u'Provision For Taxation', 'Lead': u'O'},
                 u'653': {'Name': u'Commission Expenses', 'Lead': u'UB'},
                 u'229': {'Name': u'Allowances For Doubtful Trade Receivables', 'Lead': u'J'},
                 u'226': {'Name': u'Deposits & Guarantees Given', 'Lead': u'J'},
                 u'224': {'Name': u'Unearned Leasing Interest Income', 'Lead': u'J'},
                 u'222': {'Name': u'Rediscount On Notes Receivable', 'Lead': u'E'},
                 u'221': {'Name': u'Notes Receivable', 'Lead': u'E'},
                 u'220': {'Name': u'Trade Receivables', 'Lead': u'E'},
                 u'373': {'Name': u'Allowance For Cost ', 'Lead': u'P'},
                 u'392': {'Name': u'Other Vat Payable', 'Lead': u'P'}, u'391': {'Name': u'Vat Payable', 'Lead': u'P'},
                 u'722': {'Name': u'Direct Labour (Wage Variances)', 'Lead': u'X'},
                 u'723': {'Name': u'Direct Labour (Time Variances)', 'Lead': u'X'},
                 u'720': {'Name': u'Direct Labor', 'Lead': u'X'},
                 u'721': {'Name': u'Direct Labor Reflected to Production', 'Lead': u'X'},
                 u'397': {'Name': u'Count & Delivery Surplus', 'Lead': u'P'},
                 u'659': {'Name': u'Other Operational Expense', 'Lead': u'UB'},
                 u'656': {'Name': u'F/X Loss', 'Lead': u'VE'},
                 u'151': {'Name': u'Work-In-Process', 'Lead': u'F'},
                 u'150': {'Name': u'Raw Materials And Supplies', 'Lead': u'F'},
                 u'153': {'Name': u'Merchandises', 'Lead': u'F'}, u'152': {'Name': u'Finished Goods', 'Lead': u'F'},
                 u'155': {'Name': u'Auxiliary Materials', 'Lead': u'F'}, u'154': {'Name': u'By-Products', 'Lead': u'F'},
                 u'157': {'Name': u'Other Inventories', 'Lead': u'F'},
                 u'156': {'Name': u'Packaging Materials', 'Lead': u'F'},
                 u'159': {'Name': u'Advances Given To Suppliers', 'Lead': u'F'},
                 u'158': {'Name': u'Allowances For The Decline In The Value Of Inv.', 'Lead': u'F'},
                 u'277': {'Name': u'Other Depletable Assets', 'Lead': u'J'},
                 u'361': {'Name': u'Social Security Premium Payable', 'Lead': u'P'},
                 u'239': {'Name': u'Allowances For Doubtful Receivables', 'Lead': u'J'},
                 u'235': {'Name': u'Due From Personnel', 'Lead': u'J'},
                 u'236': {'Name': u'Other Miscellaneous Receivables', 'Lead': u'J'},
                 u'237': {'Name': u'Rediscount On Other Notes Receivables', 'Lead': u'J'},
                 u'231': {'Name': u'Due From Share Holders', 'Lead': u'I'},
                 u'232': {'Name': u'Due From Participants', 'Lead': u'I'},
                 u'233': {'Name': u'Due From Subsidiaries', 'Lead': u'I'},
                 u'600': {'Name': u'Sales -Local', 'Lead': u'UA'},
                 u'322': {'Name': u'Rediscount On Notes Payable', 'Lead': u'N'},
                 u'320': {'Name': u'Trade Payables', 'Lead': u'N'}, u'321': {'Name': u'Notes Payable', 'Lead': u'N'},
                 u'326': {'Name': u'Deposits & Guarantees Received', 'Lead': u'N'},
                 u'329': {'Name': u'Other Trade Payables', 'Lead': u'N'},
                 u'771': {'Name': u'General & Admintrative Expenses (Reflected)', 'Lead': u'X'},
                 u'770': {'Name': u'General & Admintrative Expenses', 'Lead': u'X'},
                 u'654': {'Name': u'Provisions', 'Lead': u'UB'},
                 u'772': {'Name': u'General & Admintrative Expenses (Variances)', 'Lead': u'X'},
                 u'612': {'Name': u'Other Discounts', 'Lead': u'UA'}, u'610': {'Name': u'Sales Return', 'Lead': u'UA'},
                 u'733': {'Name': u'Manufacturing Overheads (Productivity Variances)', 'Lead': u'X'},
                 u'679': {'Name': u'Other Extra-Ordinary Income', 'Lead': u'UB'},
                 u'671': {'Name': u'Prior Period Income', 'Lead': u'UB'},
                 u'649': {'Name': u'Other Operational Income', 'Lead': u'UB'},
                 u'655': {'Name': u'Loss From Sale Of Marketable Securities', 'Lead': u'UB'},
                 u'481': {'Name': u'Long-Term Accrued Expenses', 'Lead': u'R'},
                 u'480': {'Name': u'Long-Term Deferred Income', 'Lead': u'R'},
                 u'263': {'Name': u'Research & Development Expenses', 'Lead': u'K'},
                 u'262': {'Name': u'Pre-Operating Expenses', 'Lead': u'K'},
                 u'732': {'Name': u'Manufacturing Overheads (Budget Variances)', 'Lead': u'X'},
                 u'261': {'Name': u'Goodwill', 'Lead': u'K'},
                 u'472': {'Name': u'Allowance For Retirement Pay', 'Lead': u'P'},
                 u'260': {'Name': u'Rights', 'Lead': u'K'},
                 u'689': {'Name': u'Other Extra-Ordinary Expenses', 'Lead': u'UB'},
                 u'350': {'Name': u'Contruction & Repair Progress Billings', 'Lead': u'P'},
                 u'267': {'Name': u'Other Intangible Assets', 'Lead': u'K'},
                 u'479': {'Name': u'Other Allowances', 'Lead': u'R'},
                 u'681': {'Name': u'Prior Period Expenses', 'Lead': u'UB'},
                 u'680': {'Name': u'Idle Capacity Expenses', 'Lead': u'UB'},
                 u'358': {'Name': u'Construction Inflation Adjustment Spread Over Years', 'Lead': u'P'}}


def prepare_db():
    for k, v in hesap_kodlari.iteritems():
        for a in v:
            db.session.add(db.Lead(account=a, lead_code=k, name=SINIFLANDIRMA[k]))

    db.session.commit()


def parse_excel_file(inp):
    """
    YENI
    :param inp: String, Path to excel file.
    :return: Array
    """

    excel_file = open_workbook(inp, encoding_override='utf-8')

    donus = dict()

    def define_variables():

        sheet = excel_file.sheet_by_index(0)

        py1 = sheet.cell(7, 2).value if not sheet.cell(7, 2).ctype is 0 else ''

        db.periods['cy'] = sheet.cell(9, 2).value
        db.periods['py2'] = sheet.cell(8, 2).value
        db.periods['py1'] = py1
        db.tanimlar['company'] = sheet.cell(6, 2).value

    define_variables()

    for i in xrange(1, 4):
        sheet = excel_file.sheet_by_index(i)
        key = sheet.name

        if key == 'CY Detailed TB':
            key = db.periods['cy']
        elif key == 'PY2 Detailed TB':
            key = db.periods['py2']
        elif key == 'PY1 Detailed TB':
            key = db.periods['py1']

        for row in xrange(sheet.nrows):
            temp = []

            if row == 0:
                continue

            hucre = unicode(int(sheet.cell_value(row, 0))) if type(
                sheet.cell_value(row, 0)) is float else sheet.cell_value(row, 0)

            if hucre not in donus:
                donus[hucre] = dict()

            for col in xrange(sheet.ncols):
                if col is 0 and type(sheet.cell_value(row, col)) is float:
                    temp.append(unicode(int(sheet.cell_value(row, col))))  # burasi duzeltilecek bug var.
                else:
                    temp.append(sheet.cell_value(row, col))

            gecici = db.session.query(db.Hesaplar).filter_by(number=hucre).first()
            ana_hesap = hucre[:3]
            lead_cod = None
            if db.session.query(db.Lead).filter_by(account=ana_hesap).first() is not None:
                lead_cod = db.session.query(db.Lead).filter_by(account=ana_hesap).first().lead_code

            if gecici is not None:
                if key is db.periods['cy']:
                    gecici.cy = temp[4]
                elif key is db.periods['py2']:
                    gecici.py2 = temp[4]
                else:
                    gecici.py1 = temp[4]
                print hucre, ' update'
            else:
                if key is db.periods['cy']:
                    db.session.add(
                        db.Hesaplar(number=hucre, len=len(hucre), lead_code=lead_cod, ana_hesap=ana_hesap, name=temp[1],
                                    cy=temp[4]))
                elif key is db.periods['py2']:
                    db.session.add(
                        db.Hesaplar(number=hucre, len=len(hucre), name=temp[1], lead_code=lead_cod, ana_hesap=ana_hesap,
                                    py2=temp[4]))
                else:
                    db.session.add(
                        db.Hesaplar(number=hucre, name=temp[1], len=len(hucre), ana_hesap=ana_hesap, lead_code=lead_cod,
                                    py1=temp[4]))
                print hucre, ' add'

        db.session.commit()


def find_bds():
    accounts = db.session.query(db.Hesaplar).all()
    for k in accounts:
        if len(db.session.query(db.Hesaplar).filter(db.Hesaplar.number.startswith(k.number)).all()) <= 1:
            k.bd = True
    db.session.commit()


def summary_check():
    if db.session.query(db.Hesaplar).filter_by(len=3).first() is not None:
        global has_sum
        has_sum = True
        return True


def create_summary_accs():
    query = db.session.query(db.Hesaplar.ana_hesap,
                             db.Hesaplar.lead_code,
                             func.sum(db.Hesaplar.py1).label('py1'),
                             func.sum(db.Hesaplar.py2).label('py2'),
                             func.sum(db.Hesaplar.cy).label('cy'),
                             ).group_by('ana_hesap')

    for k in query.all():
        unmapped = None

        if k.lead_code == 'Unmapped':
            unmapped = db.session.query(db.Hesaplar).filter_by(len=3,
                                                               ana_hesap=k.ana_hesap).first() or db.session.query(
                db.Hesaplar).filter_by(ana_hesap=k.ana_hesap).first()

        source = db.session.query(db.Lead).filter_by(account=k.ana_hesap).first()

        main_source = unmapped or source
        name = main_source.name

        db.session.add(
            db.Hesaplar(number=k.ana_hesap, ana_hesap=k.ana_hesap, name=name, lead_code=main_source.lead_code, cy=k.cy,
                        py1=k.py1, py2=k.py2, len=3))

    db.session.commit()


def fix_mainaccs():
    """
    Ana hesap isimlerini ingilizce yapiyor
    :return: None
    """
    query = db.session.query(db.Hesaplar).filter_by(len=3)

    for k in query.all():
        if k.ana_hesap in ACCOUNT_NAMES:
            item = ACCOUNT_NAMES[k.ana_hesap]['Name']
            k.name = item

    db.session.commit()


def delete_zeros(exceptions=None):
    if not exceptions:
        exceptions = ['900']
    query = db.session.query(db.Hesaplar).filter_by(py1=0, py2=0, cy=0).all()
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


'''
prepare_db()
parse_excel_file(file)
fix_mainaccs()
find_bds()
create_a4()
'''
