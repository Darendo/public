'''
Created on Mar 16, 2011

@author: Darendo
'''

import sys


def add(siteShortName, siteLongName, dbCN):

    #printEx('Getting the CollectorID ...', 4)
    
    siteID = 0
    sRet = ''
    
    # prepare a cursor object using cursor() method
    cursor = dbCN.cursor()

    # Prepare SQL query to INSERT a record into the database.
    sql = """INSERT INTO Sites (ShortName, LongName) VALUES ('%s', '%s')""" % (str(siteShortName), str(siteLongName))
    print "site.add() SQL Command = ", sql

    try:
        # Execute the SQL command
        cursor.execute(sql)
        dbCN.commit()

        # Get the new SiteID
        (siteID, sRet) = getID(siteShortName, dbCN)

    except dbCN.Error, e:
        sRet = "MySQLdb INSERT error %d: %s" % (e.args[0], e.args[1])
             
    except:
        sRet = "Unexpected error: unable to execute SQL: %s; %s" % (sql, sys.exc_info()[0])

    #print "SiteID = ", siteID
    #if sRet <> '':
    #    print "Results info: ", sRet

    return siteID, sRet


def getID(siteShortName, dbCN):

    #printEx('Getting the SiteID ...', 4)

    siteID = 0
    sRet = ''
    
    # prepare a cursor object using cursor() method
    cursor = dbCN.cursor()

    # Prepare SQL query to INSERT a record into the database.
    sql = "SELECT Sites.SiteID FROM Sites WHERE Sites.ShortName = '%s'" % (siteShortName)
    #print "wbSite.getID() SQL Command = ", sql

    try:
        #printEx('Opening the recordset ...', 6)
        cursor.execute(sql)

        results = cursor.fetchall()
        if len(results) > 0:
            for row in results:
                siteID = row[0]
        
        else:
            sRet = 'Warning: No matching sites found'

    except:
        sRet = "Error: unable to fetch SiteID for %s" % (siteShortName)

    cursor = None
    
    #print "SiteID = ", siteID
    #if sRet <> '':
    #    print "Results info: ", sRet

    return siteID, sRet
