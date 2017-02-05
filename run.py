__author__ = 'fatihka'

import os

from reader import parse_excel_file, fix_mainaccs, find_bds, delete_zeros, create_or_parse_sum
from writer import create_a4

# home = os.getenv('HOME')
# desktop = os.path.join(home, 'Downloads')
# file = os.path.join(desktop, 'A3 Deva.xlsx')

dir = os.path.abspath(".")
file = os.path.join(dir, "a3.xlsx")
# prepare_mapping(file)
# prepare_db()
parse_excel_file(file)
delete_zeros()
create_or_parse_sum()
fix_mainaccs()
find_bds()
create_a4()

# delete_zeros()
# create_a4()
