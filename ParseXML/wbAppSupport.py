'''
Created on Mar 16, 2011

@author: Darendo
'''

import os, sys, shutil
import zipfile as zip


def moveFileEx(sourceFile, destFile):
    
    bRet = False
    sRet = ''    
    symlinks = 0
    
    try:
        if symlinks and os.path.islink(sourceFile):
            linkto = os.readlink(sourceFile)
            os.symlink(linkto, destFile)

        elif os.path.isdir(sourceFile):
            #sRet = "Can't copy %s to %s: %s" % (sourceFile, destFile, 'Source is a folder')
            sRet = 'Source is a folder'
        
        else:
            shutil.copy2(sourceFile, destFile)
            os.remove(sourceFile)

        bRet = True

    except (IOError, os.error), why:
        #sRet = "Can't copy %s to %s: %s" % (sourceFile, destFile, str(why))
        sRet = "Error: %s" % (str(why))

    except:
        sRet = "Unexpected error: %s" % (sys.exc_info()[0])

    return bRet, sRet


def execSQL(sql = '', dbCN = None):

    #printEx('Executing SQL command: %s' % (sql), 4)
    
    execOK = False
    sRet = ''

    if sql == '':
        sRet = 'No SQL command passed'
        return execOK, sRet

    # prepare a cursor object using cursor() method
    cursor = dbCN.cursor()

    # Attempt to execute command
    try:
        #printEx('Executing SQL command ...', 6)
        cursor.execute(sql)

        # Commit your changes in the database
        #printEx('Committing changes ...', 6)
        dbCN.commit()
        execOK = True

    except dbCN.Error, e:
        sRet = "MySQLdb INSERT error %d: %s" % (e.args[0], e.args[1])
             
    except:
        sRet = "Unexpected error: unable to execute SQL: %s; %s" % (sql, sys.exc_info()[0])

    if not execOK:
        try:
            # Rollback in case there is any error
            dbCN.rollback()

        except:
            pass

    #if execOK == True:
    #    printEx('Executed SQL command OK', 4)
    #else:
    #    printEx('Failed to execute SQL command: %s' % (sRet), 4)

    return execOK, sRet


def unzipFile(zip_file, dest_dir=None):
    """Extract files from PKZIP archive.
       Optionally recurse and unzip zipped archives."""

    #printEx('Unzipping a file ...', 4)
    
    sRet = ''
    
    #printEx('Validating source and dest directories ...', 6)
    if dest_dir == None:
        src_dir = os.path.dirname(zip_file)
        file_name = os.path.basename(zip_file)
        dest_dir = os.path.join(src_dir, file_name)

    if os.path.exists(dest_dir) == False:
        os.mkdir(dest_dir)

    #printEx('Getting a handle to the archive ...', 6)
    zfile = zip.ZipFile(zip_file, 'r')

    #printEx('Listing contents of the archive ...', 6)
    names = zfile.namelist()

    #printEx('Extracting contents ...', 6)
    zfile.extractall(dest_dir)

    #printEx('Finished extracting %d files' % (len(names)), 4)

    return names
