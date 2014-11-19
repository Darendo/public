'''
Created on Mar 16, 2011

@author: Darendo
'''

import sys


def add(inverterMfg, inverterModel, inverterSN, collectorID, dbCN):

    #printEx('Getting the InverterID ...', 4)
    
    inverterID = 0
    sRet = ''
    
    # prepare a cursor object using cursor() method
    cursor = dbCN.cursor()

    # Prepare SQL query to INSERT a record into the database.
    sql = """INSERT INTO Inverters (Manufacturer, Model, SerialNumber, CollectorID) VALUES ('%s', '%s', '%s', %s)""" % (str(inverterMfg), str(inverterModel), str(inverterSN), str(collectorID))
    print "inverter.getID() SQL Command = ", sql

    try:
        # Execute the SQL command
        cursor.execute(sql)
        dbCN.commit()

        (inverterID, sRet) = getID(inverterModel, inverterSN, collectorID, dbCN)

    except dbCN.Error, e:
        sRet = "MySQLdb INSERT error %d: %s" % (e.args[0], e.args[1])
             
    except:
        sRet = "Unexpected error: unable to execute SQL: %s; %s" % (sql, sys.exc_info()[0])

    #print "InverterID = ", inverterID
    #if sRet <> '':
    #    print "Results info: ", sRet

    return inverterID, sRet


def getID(inverterModel, inverterSN, collectorID, dbCN):

    #printEx('Getting the InverterID ...', 4)

    inverterID = 0
    sRet = ''
    
    # prepare a cursor object using cursor() method
    cursor = dbCN.cursor()

    # Prepare SQL query to INSERT a record into the database.
    sql = "SELECT Inverters.InverterID FROM Inverters WHERE Inverters.Model = '" + str(inverterModel) + "' AND Inverters.SerialNumber = '" + str(inverterSN) + "' AND CollectorID = " + str(collectorID)
    #print "inverter.getID() SQL Command = ", sql

    try:
        #printEx('Opening the recordset ...', 6)
        cursor.execute(sql)

        # Fetch all the rows in a list of lists.
        results = cursor.fetchall()
        if len(results) > 0:
            for row in results:
                inverterID = row[0]
        
        else:
            sRet = 'Warning: No matching inverters found'

    except dbCN.Error, e:
        sRet = "MySQLdb INSERT error %d: %s" % (e.args[0], e.args[1])
             
    except:
        sRet = "Unexpected error: unable to execute SQL: %s; %s" % (sql, sys.exc_info()[0])

    #print "InverterID = ", inverterID
    #if sRet <> '':
    #    print "Results info: ", sRet

    return inverterID, sRet
