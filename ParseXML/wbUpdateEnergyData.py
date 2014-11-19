# To change this template, choose Tools | Templates
# and open the template in the editor.

__author__="Darendo"
__date__ ="$Feb 4, 2011 3:12:50 PM$"

import sys, os
import datetime, time
import MySQLdb


logFile = ''

dbServer = ''
dbDatabase = ''
dbUser = ''
dbPassword = ''



def main():

    printEx("Starting application ...")


    dbCnIsOpen = False

    # Open database connection
    printEx('Opening database connection ...', 2)
    try:
        dbCN = MySQLdb.connect(dbServer,dbUser,dbPassword,dbDatabase)
        dbCnIsOpen = True

    except dbCN.Error, e:
        print "MySQLdb Error %d: %s" % (e.args[0], e.args[1])
             
    except ValueError:
        print "Some kind of value error"

    except:
        print "Unexpected error:", sys.exc_info()[0]


    printEx("Updating EnergyData ...", 2)

    (bRet, sRet) = updateEnergyDb(dbCN)
    if bRet:
        printEx("Updated %s rows in EnergyData" % (sRet), 4)
    else:            
        printEx("Failed to updated EnergyData: %s" % (sRet), 4)


    if dbCnIsOpen == True:
        # disconnect from server
        #printEx('Closing database connection ...', 2)
        dbCN.close()


    printEx("Finished")
    print 'Finished'



def printEx(message = '', indent = 0):

    now = str(datetime.datetime.now())
    ##print now + '   ' + (' ' * abs(indent)) + message.strip() + '\n'
    tmpLog = open(logFile, 'a')
    tmpLog.write(now + '   ' + (' ' * abs(indent)) + message.strip() + '\n')
    tmpLog.close()



def updateEnergyDb(dbCN):

    bRet = False
    sRet = ''
    
    # prepare a cursor object using cursor() method
    cursor = dbCN.cursor()

    # Prepare SQL query to INSERT a record into the database.
    sql = "spUpdateEnergyData"
    #print "updateEnergyDb() SQL Command = ", sql

    try:
        # Execute the SQL command
        cursor.callproc(sql)
        #dbCN.commit()

        bRet = True
        
        results = cursor.fetchall()
        if len(results) > 0:
            for row in results:
                sRet = row[0]

        else:
            sRet = 'No rows in EnergyData table affected'


    except dbCN.Error, e:
        sRet = "MySQLdb error %d: %s" % (e.args[0], e.args[1])
             
    except:
        sRet = "Unexpected error: unable to execute SQL: %s; %s" % (sql, sys.exc_info()[0])

    #if sRet <> '':
    #    print "Results info: ", sRet

    return bRet, sRet



def loadSettings(settingsFile):

    import ConfigParser
    
    config = ConfigParser.RawConfigParser()
    config.read(settingsFile)
    
    # Set the third, optional argument of get to 1 if you wish to use raw mode.
    #print config.get('Section1', 'foo', 0) # -> "Python is fun!"
    #print config.get('Section1', 'foo', 1) # -> "%(bar)s is %(baz)s!"

    global logFile
    logFile = ''
    if config.has_option('Locations', 'Logs'):
        logFile = config.get('Locations', 'Logs', False, '')
        if logFile.startswith('~'):
            logFile = os.path.expanduser(logFile)
    if logFile == '':
        logFile = os.getcwd()
    dt = time.strftime("%Y-%m-%d")
    logFile = os.path.join(logFile, str(dt) + "_LogFile.txt")
    #print "Log File: ", logFile

    dbSection = config.get('Options', 'Database')
    
    global dbServer
    dbServer = config.get(dbSection, 'Server')
    global dbDatabase
    dbDatabase = config.get(dbSection, 'Database')
    global dbUser
    dbUser = config.get(dbSection, 'User')
    global dbPassword
    dbPassword = config.get(dbSection, 'Password')



if __name__ == "__main__":

    loadSettings(os.path.join(os.getcwd(), "ParseXML.ini"))

    main()
