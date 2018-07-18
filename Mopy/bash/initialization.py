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
"""Functions for initializing Bash data structures on boot"""
import balt
import bush
import env
from bass import dirs
from bolt import GPath
from bosh import OblivionIni, ConfigHelpers
from env import get_personal_path, get_local_app_data_path, test_permissions
from exception import BoltError, PermissionError, NonExistentDriveError

def get_path_from_ini(bash_ini_, section_key, option_key):
    if not bash_ini_ or not bash_ini_.has_option(section_key, option_key):
        return None
    get_value = bash_ini_.get(section_key, option_key)
    return GPath(get_value.strip()) if get_value != u'.' else None

def getPersonalPath(bash_ini_, my_docs_path):
    #--Determine User folders from Personal and Local Application Data directories
    #  Attempt to pull from, in order: Command Line, Ini, win32com, Registry
    if my_docs_path:
        my_docs_path = GPath(my_docs_path)
        sErrorInfo = _(u"Folder path specified on command line (-p)")
    else:
        my_docs_path = get_path_from_ini(bash_ini_, 'General', 'sPersonalPath')
        if my_docs_path:
            sErrorInfo = _(
                u"Folder path specified in bash.ini (%s)") % u'sPersonalPath'
        else:
            my_docs_path, sErrorInfo = get_personal_path()
    #  If path is relative, make absolute
    if not my_docs_path.isabs():
        my_docs_path = dirs['app'].join(my_docs_path)
    #  Error check
    if not my_docs_path.exists():
        raise BoltError(u"Personal folder does not exist.\n"
                        u"Personal folder: %s\nAdditional info:\n%s"
                        % (my_docs_path.s, sErrorInfo))
    return my_docs_path

def getLocalAppDataPath(bash_ini_, app_data_local_path):
    #--Determine User folders from Personal and Local Application Data directories
    #  Attempt to pull from, in order: Command Line, Ini, win32com, Registry
    if app_data_local_path:
        app_data_local_path = GPath(app_data_local_path)
        sErrorInfo = _(u"Folder path specified on command line (-l)")
    else:
        app_data_local_path = get_path_from_ini(bash_ini_, u'General',
                                                u'sLocalAppDataPath')
        if app_data_local_path:
            sErrorInfo = _(u"Folder path specified in bash.ini (%s)") % u'sLocalAppDataPath'
        else:
            app_data_local_path, sErrorInfo = get_local_app_data_path()
    #  If path is relative, make absolute
    if not app_data_local_path.isabs():
        app_data_local_path = dirs['app'].join(app_data_local_path)
    #  Error check
    if not app_data_local_path.exists():
        raise BoltError(u"Local AppData folder does not exist.\nLocal AppData folder: %s\nAdditional info:\n%s"
                        % (app_data_local_path.s, sErrorInfo))
    return app_data_local_path

def getOblivionModsPath(bash_ini_):
    ob_mods_path = get_path_from_ini(bash_ini_, u'General', u'sOblivionMods')
    if ob_mods_path:
        src = [u'[General]', u'sOblivionMods']
    else:
        ob_mods_path = GPath(GPath(u'..').join(u'%s Mods' % bush.game.fsName))
        src = u'Relative Path'
    if not ob_mods_path.isabs(): ob_mods_path = dirs['app'].join(ob_mods_path)
    return ob_mods_path, src

def getBainDataPath(bash_ini_):
    idata_path = get_path_from_ini(bash_ini_, u'General', u'sInstallersData')
    if idata_path:
        src = [u'[General]', u'sInstallersData']
        if not idata_path.isabs(): idata_path = dirs['app'].join(idata_path)
    else:
        idata_path = dirs['installers'].join(u'Bash')
        src = u'Relative Path'
    return idata_path, src

def getBashModDataPath(bash_ini_):
    mod_data_path = get_path_from_ini(bash_ini_, u'General', u'sBashModData')
    if mod_data_path:
        if not mod_data_path.isabs():
            mod_data_path = dirs['app'].join(mod_data_path)
        src = [u'[General]', u'sBashModData']
    else:
        mod_data_path, src = getOblivionModsPath(bash_ini_)
        mod_data_path = mod_data_path.join(u'Bash Mod Data')
    return mod_data_path, src

def getLegacyPath(newPath, oldPath):
    return (oldPath,newPath)[newPath.isdir() or not oldPath.isdir()]

def getLegacyPathWithSource(newPath, oldPath, newSrc, oldSrc=None):
    if newPath.isdir() or not oldPath.isdir():
        return newPath, newSrc
    else:
        return oldPath, oldSrc

def initDirs(bashIni_, personal, localAppData):
    #--Oblivion (Application) Directories
    dirs['app'] = bush.gamePath
    dirs['mods'] = dirs['app'].join(u'Data')
    dirs['patches'] = dirs['mods'].join(u'Bash Patches')
    dirs['defaultPatches'] = dirs['mopy'].join(u'Bash Patches', bush.game.fsName)
    dirs['tweaks'] = dirs['mods'].join(u'INI Tweaks')

    #  Personal
    personal = getPersonalPath(bashIni_, personal)
    dirs['saveBase'] = personal.join(u'My Games', bush.game.fsName)

    #  Local Application Data
    localAppData = getLocalAppDataPath(bashIni_, localAppData)
    dirs['userApp'] = localAppData.join(bush.game.fsName)

    # Use local copy of the oblivion.ini if present
    global gameInis
    global oblivionIni
    data_oblivion_ini = dirs['app'].join(bush.game.iniFiles[0])
    use_data_dir = False
    if data_oblivion_ini.exists():
        oblivionIni = OblivionIni(data_oblivion_ini)
        # is bUseMyGamesDirectory set to 0?
        if oblivionIni.getSetting(u'General', u'bUseMyGamesDirectory', u'1') == u'0':
            use_data_dir = True
    if not use_data_dir: # FIXME that's what the code was doing - loaded the ini in saveBase and tried again ??
        # oblivion.ini was not found in the game directory or
        # bUseMyGamesDirectory was not set.  Default to user profile directory
        oblivionIni = OblivionIni(dirs['saveBase'].join(bush.game.iniFiles[0]))
        # is bUseMyGamesDirectory set to 0?
        if oblivionIni.getSetting(u'General', u'bUseMyGamesDirectory', u'1') == u'0':
            use_data_dir = True
    if use_data_dir:
        # Set the save game folder to the Oblivion directory
        dirs['saveBase'] = dirs['app']
        # Set the data folder to sLocalMasterPath
        dirs['mods'] = dirs['app'].join(
            oblivionIni.getSetting(u'General', u'SLocalMasterPath', u'Data'))
        # these are relative to the mods path so they must be updated too
        dirs['patches'] = dirs['mods'].join(u'Bash Patches')
        dirs['tweaks'] = dirs['mods'].join(u'INI Tweaks')
    gameInis = [oblivionIni]
    gameInis.extend(OblivionIni(dirs['saveBase'].join(x)) for x in bush.game.iniFiles[1:])
    #--Mod Data, Installers
    oblivionMods, oblivionModsSrc = getOblivionModsPath(bashIni_)
    dirs['modsBash'], modsBashSrc = getBashModDataPath(bashIni_)
    dirs['modsBash'], modsBashSrc = getLegacyPathWithSource(
        dirs['modsBash'], dirs['app'].join(u'Data', u'Bash'),
        modsBashSrc, u'Relative Path')

    dirs['installers'] = oblivionMods.join(u'Bash Installers')
    dirs['installers'] = getLegacyPath(dirs['installers'],
                                       dirs['app'].join(u'Installers'))

    dirs['bainData'], bainDataSrc = getBainDataPath(bashIni_)

    dirs['bsaCache'] = dirs['bainData'].join(u'BSA Cache')

    dirs['converters'] = dirs['installers'].join(u'Bain Converters')
    dirs['dupeBCFs'] = dirs['converters'].join(u'--Duplicates')
    dirs['corruptBCFs'] = dirs['converters'].join(u'--Corrupt')

    #--Test correct permissions for the directories
    badPermissions = [test_dir for test_dir in dirs.itervalues()
                      if not test_permissions(test_dir)] # DOES NOTHING !!!
    if not test_permissions(oblivionMods):
        badPermissions.append(oblivionMods)
    if badPermissions:
        # Do not have all the required permissions for all directories
        # TODO: make this gracefully degrade.  IE, if only the BAIN paths are
        # bad, just disable BAIN.  If only the saves path is bad, just disable
        # saves related stuff.
        msg = balt.fill(_(u'Wrye Bash cannot access the following paths:'))
        msg += u'\n\n' + u'\n'.join(
            [u' * ' + bad_dir.s for bad_dir in badPermissions]) + u'\n\n'
        msg += balt.fill(_(u'See: "Wrye Bash.html, Installation - Windows Vista/7" for information on how to solve this problem.'))
        raise PermissionError(msg)

    # create bash user folders, keep these in order
    keys = ('modsBash', 'installers', 'converters', 'dupeBCFs', 'corruptBCFs',
            'bainData', 'bsaCache')
    try:
        env.shellMakeDirs([dirs[key] for key in keys])
    except NonExistentDriveError as e:
        # NonExistentDriveError is thrown by shellMakeDirs if any of the
        # directories cannot be created due to residing on a non-existing
        # drive. Find which keys are causing the errors
        badKeys = set()     # List of dirs[key] items that are invalid
        # First, determine which dirs[key] items are causing it
        for key in keys:
            if dirs[key] in e.failed_paths:
                badKeys.add(key)
        # Now, work back from those to determine which setting created those
        msg = _(u'Error creating required Wrye Bash directories.') + u'  ' + _(
            u'Please check the settings for the following paths in your '
            u'bash.ini, the drive does not exist') + u':\n\n'
        relativePathError = []
        if 'modsBash' in badKeys:
            if isinstance(modsBashSrc, list):
                msg += (u' '.join(modsBashSrc) + u'\n    '
                        + dirs['modsBash'].s + u'\n')
            else:
                relativePathError.append(dirs['modsBash'])
        if {'installers', 'converters', 'dupeBCFs', 'corruptBCFs'} & badKeys:
            # All derived from oblivionMods -> getOblivionModsPath
            if isinstance(oblivionModsSrc, list):
                msg += (u' '.join(oblivionModsSrc) + u'\n    '
                        + oblivionMods.s + u'\n')
            else:
                relativePathError.append(oblivionMods)
        if {'bainData', 'bsaCache'} & badKeys:
            # Both derived from 'bainData' -> getBainDataPath
            # Sometimes however, getBainDataPath falls back to oblivionMods,
            # So check to be sure we haven't already added a message about that
            if bainDataSrc != oblivionModsSrc:
                if isinstance(bainDataSrc, list):
                    msg += (u' '.join(bainDataSrc) + u'\n    '
                            + dirs['bainData'].s + u'\n')
                else:
                    relativePathError.append(dirs['bainData'])
        if relativePathError:
            msg += u'\n' + _(u'A path error was the result of relative paths.')
            msg += u'  ' + _(u'The following paths are causing the errors, '
                             u'however usually a relative path should be fine.')
            msg += u'  ' + _(u'Check your setup to see if you are using '
                             u'symbolic links or NTFS Junctions') + u':\n\n'
            msg += u'\n'.join([u'%s' % x for x in relativePathError])
        raise BoltError(msg)

    # Setup LOOT API, needs to be done after the dirs are initialized
    global configHelpers
    configHelpers = ConfigHelpers()
