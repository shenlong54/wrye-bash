# -*- coding: utf-8 -*-
#
# GPL License and Copyright Notice ============================================
#  This file is part of Wrye Bash.
#
#  Wrye Bash is free software; you can redistribute it and/or
#  modify it under the terms of the GNU General Public License
#  as published by the Free Software Foundation; either version 2
#  of the License, or (at your option) any later version.
#
#  Wrye Bash is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#
#  You should have received a copy of the GNU General Public License
#  along with Wrye Bash; if not, write to the Free Software Foundation,
#  Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
#
#  Wrye Bash copyright (C) 2005-2009 Wrye, 2010-2015 Wrye Bash Team
#  https://github.com/wrye-bash
#
# =============================================================================

"""This module parses the command line that was used to start Wrye Bash."""

import argparse

def parse():
    """Helper function to define commandline arguments"""
    parser = argparse.ArgumentParser()

    #### Groups ####
    def arg(group, dashed, descr, dest, action='store', default=None):
        group.add_argument(dashed, descr, dest=dest, action=action,
                           default='' if default is None else default,
                           help=h) # so we can wrap help but not too much

    ### Path Group ###
    pathGroup = parser.add_argument_group("Path Arguments",
        r"""All path arguments must be absolute paths and use either forward
        slashes (/) or two backward slashes (\\). All of these can also be
        set in the ini (where you can also use relative paths) and if set
        in both cmd line takes precedence.""")
    # oblivionPath #
    h = """Specifies the game directory (the one containing the game's exe).
    Use this argument if Bash is located outside of the game directory,
    and the --game argument failed to find it."""
    arg(pathGroup, '-o', '--oblivionPath', dest='oblivionPath')

    ### User Path Group ###
    userPathGroup = parser.add_argument_group("User Directory Arguments",
        """These arguments allow you to specify your user directories in
        several ways. These are only useful if the regular procedure for
        getting the user directory fails. And even in that case, the user
        is probably better off installing win32com.""")
    # personalPath #
    h = r"""Specify the user's personal directory. (Like "C:\\Documents and
    Settings\\Wrye\\My Documents") If you need to set this then you probably
    need to set -l too"""
    arg(userPathGroup, '-p', '--personalPath', dest='personalPath')
    # userPath #
    h = """Specify the user profile path. May help if HOMEDRIVE and/or
    HOMEPATH are missing from the user's environment"""
    arg(userPathGroup, '-u', '--userPath', dest='userPath')
    # localAppDataPath #
    h = """Specify the user's local application data directory.If you need
    to set this then you probably need to set -p too."""
    arg(userPathGroup, '-l', '--localAppDataPath', dest='localAppDataPath')

    ### Backup Group ###
    backupGroup = parser.add_argument_group("Backup and Restore Arguments",
        """These arguments allow you to do backup and restore settings
        operations.""")
    # backup #
    h = """Backup all Bash settings to an archive file before the app
    launches. You have to specify the filepath with the -f/--filename
    option. If also -r is specified Bash will not start."""
    arg(backupGroup, '-b', '--backup', dest='backup', action='store_true',
        default=False)
    # restore #
    h = """Restore all Bash settings from an archive file before the app
    launches. You have to specify the filepath with the -f/--filename
    option. If also -b is specified Bash will not start."""
    arg(backupGroup, '-r', '--restore', dest='restore', action='store_true',
        default=False)
    # filename #
    h = """The file to use with the -r or -b options. Must end in '.7z' (or
    be a valid directory path for -r) and be a valid path and for -r exist
    and for -b not already exist."""
    arg(backupGroup, '-f', '--filename', dest='filename')
    # quietquit #
    h = """Close Bash after creating or restoring backup and do not display
    any prompts or message dialogs."""
    arg(backupGroup, '-q', '--quiet-quit', dest='quietquit',
        action='store_true', default=False)

    #### Individual Arguments ####
    parser.add_argument('-d', '--debug',
                        action='store_true',
                        default=False,
                        dest='debug',
                        help='Useful if bash is crashing on startup or if '
                             'you want to print a lot of information'
                             ' (e.g. while developing or '
                             'debugging).')
    parser.set_defaults(mode=0)
    parser.add_argument('-C', '--Cbash-mode',
                        action='store_const',
                        const=2,
                        dest='mode',
                        help='enables CBash and uses CBash to build bashed '
                             'patch.')
    parser.add_argument('-P', '--Python-mode',
                        action='store_const',
                        const=1,
                        dest='mode',
                        help='disables CBash and uses python code to build '
                             'bashed patch.')
    parser.add_argument('--no-uac',
                        action='store_true',
                        dest='noUac',
                        help='suppress the prompt to restart in admin mode '
                             'when UAC is detected.')
    parser.add_argument('--uac',
                        action='store_true',
                        dest='uac',
                        help='always start in admin mode if UAC protection is '
                             'detected.')
    parser.add_argument('--genHtml',
                        default=None,
                        help=argparse.SUPPRESS)
    parser.add_argument('-L', '--Language',
                        action='store',
                        default='',
                        dest='language',
                        help='Specify the user language overriding the system '
                             'language settings.')
    # parse and error check backup options
    args = parser.parse_args()
    if args.backup and args.restore:
        parser.error('You specified both backup and restore')
    elif (args.backup or args.restore) and not args.filename:
        parser.error('You must specify a filename for use with backup/restore')
    return args

_short_to_long = dict(
    [('-o', '--oblivionPath'), ('-p', '--personalPath'), ('-u', '--userPath'),
     ('-l', '--localAppDataPath'), ('-b', '--backup'), ('-r', '--restore'),
     ('-q', '--quiet-quit'), ('-f', '--filename'),
     ('-d', '--debug'), ('-C', '--Cbash-mode'), ('-P', '--Python-mode'),
     ('-L', '--Language'), ])

def convert_to_long_options(sys_argv):
    sys_argv = list(sys_argv)
    for j, arg in enumerate(sys_argv):
        if arg in _short_to_long:
            sys_argv[j] = _short_to_long[arg]
    return sys_argv
