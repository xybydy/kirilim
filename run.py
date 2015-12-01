import os

import click

from reader import prepare_db, parse_excel_file, delete_zeros, create_or_parse_sum, fix_mainaccs, find_bds
from writer import create_a4


def runbaby(f):
    prepare_db()
    parse_excel_file(f)
    delete_zeros()
    create_or_parse_sum()
    fix_mainaccs()
    find_bds()
    create_a4()


def kir_falan(a):
    cwd = os.getcwdu()

    if a[0] is '.':
        for f in os.listdir(cwd):
            ext = f.split('.')[-1:]

            if ext is not 'xlsx':
                continue
            else:
                inp = os.path.join(cwd, f)
                if os.path.exists(inp):
                    runbaby(inp)
    else:
        for param in a:
            if 'xlsx' not in param:
                param += '.xlsx'

            inp = os.path.join(cwd, param)
            print inp
            if os.path.exists(inp):
                runbaby(inp)


@click.group()
def cli():
    pass


kir_help = '''
You can enter multiple filenames. Just place a dot instead of a filename to perform breakdown on all files on working directory. File extension is optional\n
i.e. run.py kir .\tor\trun.py kir abc qwe zxc aaa.xlsx
'''


@cli.command(short_help='Perform breakdown procedures.', help=kir_help)
@click.option('--persistent', '-p', is_flag=True, help='Persistent database')
@click.option('--debug', '-d', default=False, is_flag=True, help='Debug Mode.')
@click.argument('filename', nargs=-1)
@click.pass_context
def kir(ctx, debug, persistent, filename):
    if len(filename) == 0:
        print ctx.get_help()
    else:
        kir_falan(filename)


if __name__ == '__main__':
    cli()