'''
Created on Mar 16, 2011

@author: Darendo
'''

import sys


def add(collectorMfg, collectorModel, collectorModelAbbr, collectorSN, siteID, dbCN):

    #printEx('Getting the CollectorID ...', 4)
    
    collectorID = 0
    sRet = ''
    
    # prepare a cursor object using cursor() method
    cursor = dbCN.cursor()

    # Prepare SQL query to INSERT a record into the database.
    sql = """INSERT INTO Collectors (Manufacturer, Model, ModelAbbr, SerialNumber, SiteID) VALUES ('%s', '%s', '%s', '%s', %s)""" % (str(collectorMfg), str(collectorModel), str(collectorModelAbbr), str(collectorSN), str(siteID))
    print "collector.add() SQL Command = ", sql

    try:
        # Execute the SQL command
        cursor.execute(sql)
        dbCN.commit()

        (collectorID, sRet) = getID(collectorModelAbbr, collectorSN, siteID, dbCN)

    except dbCN.Error, e:
        sRet = "MySQLdb INSERT error %d: %s" % (e.args[0], e.args[1])
             
    except:
        sRet = "Unexpected error: unable to execute SQL: %s; %s" % (sql, sys.exc_info()[0])

    #print "CollectorID = ", collectorID
    #if sRet <> '':
    #    print "Results info: ", sRet

    return collectorID, sRet


def getID(collectorModelAbbr, collectorSN, siteID, dbCN):

    #printEx('Getting the CollectorID ...', 4)

    collectorID = 0
    sRet = ''

    # prepare a cursor object using cursor() method
    cursor = dbCN.cursor()

    # Prepare SQL query to INSERT a record into the database.
    sql = "SELECT Collectors.CollectorID FROM Collectors WHERE Collectors.ModelAbbr = '" + str(collectorModelAbbr) + "' AND Collectors.SerialNumber = '" + str(collectorSN) + "' AND Collectors.SiteID = '" + str(siteID) + "'"
    #print "collector.getID() SQL Command = ", sql

    try:
        # Execute the SQL command
        #printEx('Opening the recordset ...', 6)
        cursor.execute(sql)

        # Fetch all the rows in a list of lists.
        results = cursor.fetchall()
        if len(results) > 0:
            for row in results:
                collectorID = row[0]
        
        else:
            sRet = 'Warning: No matching collectors found'

    except dbCN.Error, e:
        sRet = "MySQLdb SELECT Error %d: %s" % (e.args[0], e.args[1])

    except:
        sRet = "Unexpected error: unable to execute SQL: %s; %s" % (sql, sys.exc_info()[0])
    
    #print "CollectorID = ", collectorID
    #if sRet <> '':
    #    print "Results info: ", sRet

    return collectorID, sRet
