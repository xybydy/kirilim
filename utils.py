import sys
from time import sleep

from colored import stylize, fg, attr


def flush(msg, err=None, fast=None, wait=0, code='reg'):
    codes = dict(
        error=fg('red') + attr('bold'),
        reg=fg(28) + attr('bold'),
        blue=fg('blue')
    )

    if err:
        if fast:
            print(stylize('\n[-] {0}'.format(msg), codes['error']), end='')
        else:
            sys.stdout.write(stylize('\n[-] ', codes['error']))
            for char in msg:
                sys.stdout.write(stylize('%s' % char, codes['error']))
                sys.stdout.flush()
                sleep(wait)
    else:
        if fast:
            print(stylize('\n[+] {0}'.format(msg), codes[code]), end='')
        else:
            print(stylize('\n[+] ', codes[code]), end='')
            for char in msg:
                sys.stdout.write(stylize('%s' % char, codes[code]))
                sys.stdout.flush()
                sleep(wait)
