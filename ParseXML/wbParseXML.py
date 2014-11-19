# To change this template, choose Tools | Templates
# and open the template in the editor.

__author__="Darendo"
__date__ ="$Feb 4, 2011 3:12:50 PM$"

import sys, os
import datetime, time
from xml.sax import make_parser 
from xml.sax.handler import ContentHandler 
from xml.etree import ElementTree as ET
import MySQLdb

import wbAppSupport
import wbSite
import wbCollector
import wbInverter
import wbDataSet


logFile = ''

dbServer = ''
dbDatabase = ''
dbUser = ''
dbPassword = ''

dirSource = ''
dirArchives = ''

excludeList = []

dsLocations = wbDataSet.DataSet.Locations()
dsInfo = wbDataSet.DataSet.Info()


def main(src_dir, arc_dir):

    #printEx('Starting application')
    
    numTotalFilesProcessed = 0
    
    dsLocations.dirSource = src_dir
    #print dsLocations.dirSource

    #printEx('Listing data files ...', 2)
    archiveFilesList = listDataFiles(dsLocations.dirSource)
    numArchives = len(archiveFilesList)
    numArchive = 0
    #printEx('Found %d data files' % (numArchives), 2)


    if numArchives > 0:
        # Open database connection
        #printEx('Opening database connection ...', 2)
        dbCN = MySQLdb.connect(dbServer,dbUser,dbPassword,dbDatabase)
        dbCnIsOpen = True
    else:
        dbCnIsOpen = False
        bRet = False
        

    for eachArchiveFile in archiveFilesList:

        bRet = False
        tmrStart = datetime.datetime.now()
        
        numArchive = numArchive + 1
        #printEx("Processing %s, %d%%" % (eachArchiveFile, ((numArchive/numArchives)*100)), 2)

        #printEx('Parsing file name ...', 2)
        file_parts = eachArchiveFile.split("/")


        dsInfo.zipFile = file_parts[len(file_parts)-1]
        dsInfo.fileName = dsInfo.zipFile[0:27]

        
        dsInfo.siteName = file_parts[len(file_parts)-2]
        (lRet, sRet) = wbSite.getID(dsInfo.siteName, dbCN)
        if lRet <> 0:
            dsInfo.siteID = lRet
        else:
            printEx("Call to wbSite.getID() failed: %s; calling wbSite.add()" % (sRet), 2)
            (lRet, sRet) = wbSite.add(dsInfo.siteName, dsInfo.siteName, dbCN)
            if lRet <> 0:
                dsInfo.siteID = lRet
            else:
                printEx("Call to wbSite.add() failed: %s; skipping %s" % (sRet, eachArchiveFile), 2)
                break
        #printEx('SiteID = %d' % (dsInfo.siteID), 2)


        dsInfo.collectorAbbr = dsInfo.fileName[0:2]
        dsInfo.collectorSN = dsInfo.fileName[2:11]
        (lRet, sRet) = wbCollector.getID(dsInfo.collectorAbbr, dsInfo.collectorSN, dsInfo.siteID, dbCN)
        if lRet <> 0:
            dsInfo.collectorID = lRet
        else:
            printEx("Call to wbCollector.getID() failed: %s; calling wbCollector.add()" % (sRet), 2)
            (lRet, sRet) = wbCollector.add('Sunny', 'WEBBOX-J2', dsInfo.collectorAbbr, dsInfo.collectorSN, dsInfo.siteID, dbCN)
            if lRet <> 0:
                dsInfo.collectorID = lRet
            else:
                printEx("Call to wbCollector.add() failed: %s; skipping %s" % (sRet, eachArchiveFile), 2)
                break
        #printEx('CollectorID = %d' % (dsInfo.collectorID), 2)
    
    
        try:
            tmpDateTime = dsInfo.fileName[12:27]
            dt = datetime.datetime.strptime(tmpDateTime, '%Y%m%d-%H%M%S')
            #tmpDateTime = dt.strftime("%Y-%m-%d %H:%M:%S")
            #dsInfo.dateFile = tmpDateTime
            dsInfo.dateFile = dt

        except:
            printEx("Unexpected error: Failed to get dsInfo.dateCollected from %s; %s" % (dsInfo.zipFile, sys.exc_info()[0]), 2)


        #print "dsInfo.zipFile = ", dsInfo.zipFile
        #print "dsInfo.fileName = ", dsInfo.fileName
        #print "dsInfo.siteName = ", dsInfo.siteName
        #print "dsInfo.siteID = ", dsInfo.siteID
        #print "dsInfo.collectorAbbr = ", dsInfo.collectorAbbr 
        #print "dsInfo.collectorSN = ", dsInfo.collectorSN
        #print "dsInfo.collectorID = ", dsInfo.collectorID
        #print "dsInfo.dateFile = ", dsInfo.dateFile


        #printEx('Setting up the destination directory ...', 2)
        dest_dir = os.path.join(dsLocations.dirSource, dsInfo.siteName, dsInfo.fileName)
        if not dest_dir.endswith("/"):
            dest_dir = dest_dir + "/"
        dsLocations.dirDest = dest_dir
        if not os.path.exists(dsLocations.dirDest):
            os.mkdir(dsLocations.dirDest)
        #print "dsLocations.dirDest = ", dsLocations.dirDest


        #printEx('Unpacking data file ...', 2)
        numFilesUnpacked, filesToProcess = unpackDataFile(eachArchiveFile, dsLocations.dirDest)
        numFilesToProcess = len(filesToProcess)
        #printEx('Unpacked %d files, will process %d of them' % (numFilesUnpacked, numFilesToProcess), 2)
        numFilesProcessed = 0


        if numFilesUnpacked > 0:

            for eachDataFile in filesToProcess:
        
                if eachDataFile.startswith("Mean.") and eachDataFile.endswith(".xml"):

                    bRet = False

                    # Mean.20110211_171507.xml
                    (tmpFilePath, tmpFileName) = os.path.split(eachDataFile)
                    #print tmpFilePath
                    #print tmpFileName
                    try:
                        tmpFileNameParts = tmpFileName.split('.')
                        tmpDate = tmpFileNameParts[1]
                        tmpDate = tmpDate.replace("_", " ")
                        tmpDate = tmpDate.replace("-", " ")
                        dtFile = datetime.datetime.strptime(tmpDate, '%Y%m%d %H%M%S')
                        #print "Date File = ", datetime.date.isoformat(dtFile)
                    except:
                        printEx("Failed to get file date/time from filename in %s; %s" % (eachDataFile, sys.exc_info()[0]), 2)
                        dtFile = None

                    #print "Parsing MEAN data file", eachDataFile
                    #printEx('Parsing %s ...' % (tmpFileName), 2)

                    sql = ''
                    
                    data_file = os.path.join(dsLocations.dirDest, eachDataFile)
                    allData = parseDataFile(data_file)
                    #print "parseDataFile(%s) returned allData[%d]" % (data_file, len(allData))
                    
                    for dsData in allData:

                        #print "Inverter model:   ", dsData.inverterModel
                        #print "Inverter serial:  ", dsData.inverterSN
                        #print "Data Description: ", dsData.datasetAbbr
                        #print "    First = ", dsData.valFirst
                        #print "    Last = ", dsData.valLast
                        #print "    Min = ", dsData.valMin
                        #print "    Max = ", dsData.valMax
                        #print "    Mean = ", dsData.valMean
                        #print "    Base = ", dsData.valBase
                        #print "    Period = ", dsData.period
                        #print "    Date Collected = ", dsData.dateCollected
                        #print "    Date File = ", dsData.dateFile
                        
                        #printEx('Getting InverterID ...', 4)
                        (lRet, sRet) = wbInverter.getID(dsData.inverterModel, dsData.inverterSN, dsInfo.collectorID, dbCN)
                        if lRet <> 0:
                            dsData.inverterID = lRet
                        else:
                            printEx("Call to wbInverter.getID() failed: %s; calling wbInverter.add()" % (sRet), 2)
                            (lRet, sRet) = wbInverter.add('Sunny', dsData.inverterModel, dsData.inverterSN, dsInfo.collectorID, dbCN)
                            if lRet <> 0:
                                dsData.inverterID = lRet
                            else:
                                printEx("Call to wbInverter.add() failed: %s; skipping %s" % (sRet, eachArchiveFile), 2)
                                break
                        #printEx('InverterID = %d' % (dsData.inverterID), 2)


                        #printEx('InverterID = %d' % (dsData.inverterID), 4)
                        tmpDateCollected = dsData.dateCollected
                        #print "tmpDateCollected = ", tmpDateCollected
                        
                        tmpDateLogged = time.strftime("%Y-%m-%d %H:%M:%S")
                        
                        # Prepare SQL query to INSERT a record into the database.
                        if len(sql) == 0:
                            sql = "('%s', %d, %d, %d, '%s', %f, %f, %f, %f, %f, %f, %d, '%s', '%s')" \
                            % (str(tmpDateCollected), dsInfo.siteID, dsInfo.collectorID, dsData.inverterID, dsData.datasetAbbr, float(dsData.valFirst), float(dsData.valLast), float(dsData.valMin), float(dsData.valMax), float(dsData.valMean), float(dsData.valBase), int(dsData.period), dsInfo.zipFile, tmpDateLogged)
                        else:
                            sql = sql + ", \n    ('%s', %d, %d, %d, '%s', %f, %f, %f, %f, %f, %f, %d, '%s', '%s')" \
                            % (str(tmpDateCollected), dsInfo.siteID, dsInfo.collectorID, dsData.inverterID, dsData.datasetAbbr, float(dsData.valFirst), float(dsData.valLast), float(dsData.valMin), float(dsData.valMax), float(dsData.valMean), float(dsData.valBase), int(dsData.period), dsInfo.zipFile, tmpDateLogged)

                    if len(sql) > 0:
                        sql = "INSERT IGNORE INTO rawData (`DateCollected`, `SiteID`, `CollectorID`, `InverterID`, `DatasetName`, `FirstValue`, `LastValue`, `MinValue`, `MaxValue`, `MeanValue`, `BaseValue`, `Period`, `SourceFile`, `DateLogged`) VALUES " + sql
                        #print "Executing SQL INSERT statement: ", sql
                        #printEx("Executing SQL INSERT statement ...", 2)
                        (bRet, sRet) = wbAppSupport.execSQL(sql, dbCN)
                        if bRet == True:
                            tmpSQL = []
                            if '\n' in sql:
                                tmpSQL = sql.splitlines()
                            else:
                                tmpSQL[0] = sql
                            #printEx("SQL = %s" % (tmpSQL[0]), 2)
                            #print "Added data for ", eachDataFile

                        else:
                            #printEx("Error: Failed to execute SQL for %s" % (eachDataFile), 2)
                            printEx("Error: Failed to execute SQL for %s; %s" % (eachDataFile, sRet), 2)
                            printEx("SQL = %s" % (sql), 2)

                    else:
                        printEx("No SQL to execute for %s" % (eachDataFile), 2)
                        bRet = False
                    

                    if bRet == True:
                        numFilesProcessed = numFilesProcessed + 1


                elif eachDataFile.startswith("Log.") and eachDataFile.endswith(".xml"):
                    numFilesToProcess = numFilesToProcess - 1


                elif eachDataFile.startswith("Info.") and eachDataFile.endswith(".xml"):
                    numFilesToProcess = numFilesToProcess - 1


                else:
                    print "Skipping ", eachDataFile


            # Delete the actual XML data files
            #printEx("Deleting temporary data files ...", 2)
            for eachDataFile in os.listdir(dsLocations.dirDest):
                os.remove(os.path.join(dsLocations.dirDest, eachDataFile))
            os.rmdir(dsLocations.dirDest)


            tmrEnd = datetime.datetime.now()
            tmrDuration = tmrEnd - tmrStart

            
            if numFilesProcessed == numFilesToProcess:
                printEx("Processed %d XML files in %s OK in %s" % (numFilesProcessed, eachArchiveFile, str(tmrDuration)), 2)
                print "Processed %d XML files in %s OK in %s" % (numFilesProcessed, eachArchiveFile, str(tmrDuration))
                
            else:
                printEx("Error: Processed %d, expected %d XML files in %s" % (numFilesProcessed, numFilesToProcess, eachArchiveFile), 2)
                print "Error: Processed %d, expected %d XML files in %s" % (numFilesProcessed, numFilesToProcess, eachArchiveFile)

            numTotalFilesProcessed = numTotalFilesProcessed + numFilesProcessed


        else:
            bRet = True
            printEx("Warning: No files unpacked from %s" % (eachArchiveFile))
            print "Warning: No files unpacked from %s" % (eachArchiveFile)


        # Archive the original file
        if bRet == True:
            dirArchive = os.path.join(arc_dir, dsInfo.siteName)
            if not os.path.exists(dirArchive):
                os.mkdir(dirArchive)
            #print "dirArchive = ", dirArchive
            archivedFile = os.path.join(dirArchive, dsInfo.zipFile)
            (bRet, sRet) = wbAppSupport.moveFileEx(eachArchiveFile, archivedFile)
            if bRet == False:
                printEx('Error: Failed to archive %s; %s' % (eachArchiveFile, sRet), '2')

        else:
            printEx('Skipped archiving %s' % (eachArchiveFile), 2)


        ###break


    if dbCnIsOpen == True:

        (bRet, sRet) = updateEnergyDb(dbCN)
        if bRet:
            printEx("Updated %s rows in EnergyData" % (sRet), 2)
            print "Updated %s rows in EnergyData" % (sRet)
        else:            
            printEx("Failed to updated EnergyData: %s" % (sRet), 2)
            print "Failed to updated EnergyData: %s" % (sRet)

        # disconnect from server
        #printEx('Closing database connection ...', 2)
        dbCN.close()


    if bRet == True:
        printEx("Finished")
        print 'Finished OK; processed %s files' % (str(numTotalFilesProcessed))
    else:
        if numArchives > 0:
            print 'Finished with error; failed to process %d of %d archives' % (numArchives-numArchive, numArchives)
        else:
            print 'Nothing to do.'
        


def printEx(message = '', indent = 0):
    #pass
    now = str(datetime.datetime.now())
    ##print now + '   ' + (' ' * abs(indent)) + message.strip() + '\n'
    tmpLog = open(logFile, 'a')
    tmpLog.write(now + '   ' + (' ' * abs(indent)) + message.strip() + '\n')
    tmpLog.close()



def listDataFiles(src_dir):

    #printEx('Listing data files ...', 4)
    
    files_list = []
    
    #printEx('Calling os.listdir()', 6)
    all_files = os.listdir(src_dir)
    #print len(all_files)

    for each_file in all_files:
        #print "Found " + each_file
        if os.path.isfile(os.path.join(src_dir, each_file)):
            if each_file.startswith("wb") and  each_file.endswith(".zip"):
                files_list.append(os.path.join(src_dir, each_file))
            #    print "  Appended " + each_file
            #else:
            #    print "  Ignored " + each_file
        elif os.path.isdir(os.path.join(src_dir, each_file)):
            more_files = listDataFiles(os.path.join(src_dir, each_file))
            for the_file in more_files:
                files_list.append(os.path.join(src_dir, the_file))

    #printEx('Found %d files' % (len(files_list)), 4)

    return files_list


def unpackDataFile(data_file, dest_dir):

    #printEx('Unpacking data file ...', 4)
    
    files_list = []
    num_files = 0
    
    try:
        output_files = wbAppSupport.unzipFile(data_file, dest_dir)
    
    except:
        printEx("There was an exception unzipping %s" % (data_file))

    else:
        for each_file in output_files:
            num_files = num_files + 1
            if (each_file.startswith("Mean.") or each_file.startswith("Log.")) and (each_file.endswith(".xml") or each_file.endswith(".zip")):
            #if (each_file.startswith("Mean.")) and (each_file.endswith(".xml") or each_file.endswith(".zip")):
                #print each_file
                #zfile.extract(each_file, dest_dir)
                if each_file.endswith(".zip"):
                    sub_files_list = wbAppSupport.unzipFile(os.path.join(dest_dir, each_file), dest_dir)
                    os.remove(os.path.join(dest_dir, each_file))
                    for each_sub_file in sub_files_list:
                        files_list.append(each_sub_file)
                elif each_file.endswith(".xml"):
                    files_list.append(each_file)

    return num_files, files_list



def parseDataFile(data_file):

    #printEx('Parsing the XML file ...', 4)

    allData = []

    #printEx('Importing xml.etree.ElementTree as ET ...', 6)
    #from xml.etree import ElementTree as ET

    tree = ET.parse(data_file)

    #printEx('Finding all MeanPublics ...', 6)
    for parent in tree.findall("MeanPublic"):

        #<MeanPublic>
        #  <Key>WR40U08E:2001385972:Backup State</Key>
        #  <First>0</First>
        #  <Last>0</Last>
        #  <Min>0</Min>
        #  <Max>0</Max>
        #  <Mean>0</Mean>
        #  <Base>19</Base>
        #  <Period>300</Period>
        #  <TimeStamp>2011-04-02T12:10:21</TimeStamp>
        #</MeanPublic>


        #printEx('Instantiating dsData object ...', 8)
        dsData = wbDataSet.DataSet.RawData()
        #printEx('Resetting dsData values ...', 8)
        dsData.resetValues()


        # Mean.20110211_171507.xml
        (filePath, fileName) = os.path.split(data_file)
        #print filePath
        #print fileName
        try:
            fileNameParts = fileName.split('.')
            tmp = fileNameParts[1]
            tmp = tmp.replace("_", " ")
            tmp = tmp.replace("-", " ")
            dt = datetime.datetime.strptime(tmp, '%Y%m%d %H%M%S')
            #dsData.dateFile = dt.strftime("%Y-%m-%d %H:%M:%S")
            dsData.dateFile = dt
            #print "    Date File = ", dsData.dateFile
        except:
            printEx("Warning: Failed to get dsData.dateFile from %s" % (data_file), 8)


        #printEx('Enumerating children ...', 8)
        for child in parent:
            #printEx('Evaluating child ...', 10)
            if child.tag == "Key":
                #print child.text
                parts = child.text.split(":")
                dsData.inverterModel = parts[0]
                dsData.inverterSN = parts[1]
                dsData.datasetAbbr = parts[2]
                #print "Inverter model:   ", dsData.inverterModel
                #print "Inverter serial:  ", dsData.inverterSN
                #print "Data Description: ", dsData.datasetAbbr
            elif child.tag == "First":
                dsData.valFirst = child.text
                #print "    First = ", dsData.valFirst
            elif child.tag == "Last":
                dsData.valLast = child.text
                #print "    Last = ", dsData.valLast
            elif child.tag == "Min":
                dsData.valMin = child.text
                #print "    Min = ", dsData.valMin
            elif child.tag == "Max":
                dsData.valMax = child.text
                #print "    Max = ", dsData.valMax
            elif child.tag == "Mean":
                dsData.valMean = child.text
                #print "    Mean = ", dsData.valMean
            elif child.tag == "Base":
                dsData.valBase = child.text
                #print "    Base = ", dsData.valBase
            elif child.tag == "Period":
                dsData.period = child.text
                #print "    Period = ", dsData.period
            elif child.tag == "TimeStamp":
                try:
                    tmp = str(child.text)
                    tmp = tmp.replace('T', ' ')
                    dt = datetime.datetime.strptime(tmp, '%Y-%m-%d %H:%M:%S')
                    #dsData.dateCollected = dt.strftime("%Y-%m-%d %H:%M:%S")
                    dsData.dateCollected = dt
                    #print "    Date Logged = ", dsData.dateCollected
                except:
                    printEx("Warning: Failed to get dsData.dateCollected from %s" % (data_file), 8)
    
        #printEx('Validating data ...', 6)
        if dsData.isValid:
            #print "dsData.isValid = True"
            #if dsData.datasetAbbr == "Balancer" or \
            #dsData.datasetAbbr == "Event-Cnt" or \
            #dsData.datasetAbbr == "Error" or \
            #dsData.datasetAbbr == "Grid Type" or \
            #dsData.datasetAbbr == "Mode" or \
            #dsData.datasetAbbr == "Serial Number" or \
            #dsData.datasetAbbr == "Temperature" or \
            #dsData.datasetAbbr == "Vfan":
            if dsData.datasetAbbr in excludeList:
                pass
            else:
                allData.append(dsData)

        else:
            print "****************************************"
            print "dsData.isValid = False"
            print "Inverter model:   ", dsData.inverterModel
            print "Inverter serial:  ", dsData.inverterSN
            print "Data Description: ", dsData.datasetAbbr
            print "    First = ", dsData.valFirst
            print "    Last = ", dsData.valLast
            print "    Min = ", dsData.valMin
            print "    Max = ", dsData.valMax
            print "    Mean = ", dsData.valMean
            print "    Base = ", dsData.valBase
            print "    Period = ", dsData.period
            print "    Date Logged = ", dsData.dateCollected
            print "****************************************"

    #printEx('Finished parsing the XML file', 4)

    return allData


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


    
def testCode():

    dbCnIsOpen = False

    # Open database connection
    #printEx('Opening database connection ...', 2)
    try:
        dbCN = MySQLdb.connect(dbServer,dbUser,dbPassword,dbDatabase)
        dbCnIsOpen = True

    except dbCN.Error, e:
        print "MySQLdb Error %d: %s" % (e.args[0], e.args[1])
             
    except ValueError:
        print "Some kind of value error"

    except:
        print "Unexpected error:", sys.exc_info()[0]


    #lRet = wbCollector.add('Sunny', 'WEBBOX-J2', 'wb', '150059364', '3', dbCN)
    #print "wbCollector.add('Sunny', 'WEBBOX-J2', 'wb', '150059364', '3') returned ", lRet

    #sRet = updateEnergyDb(dbCN)
    #print sRet


    if dbCnIsOpen == True:
        # disconnect from server
        #printEx('Closing database connection ...', 6)
        #dbCN.commit()
        dbCN.close()    


#def reviewStats(statsFile):
#    pass
#    import pstats
#    p = pstats.Stats(statsFile)
#    p.strip_dirs().sort_stats(-1).print_stats()
#    p.sort_stats('time', 'cum')
#    p.print_stats()
#
#    #p.sort_stats('cumulative').print_stats(10)
#    #p.sort_stats('time').print_stats(10)


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

    global dirSource
    dirSource = config.get('Locations', 'Source')
    if dirSource.startswith('~'):
        dirSource = os.path.expanduser(dirSource)

    global dirArchives
    dirArchives = config.get('Locations', 'Archives')
    if dirArchives.startswith('~'):
        dirArchives = os.path.expanduser(dirArchives)

    dbSection = config.get('Options', 'Database')
    
    global dbServer
    dbServer = config.get(dbSection, 'Server')
    global dbDatabase
    dbDatabase = config.get(dbSection, 'Database')
    global dbUser
    dbUser = config.get(dbSection, 'User')
    global dbPassword
    dbPassword = config.get(dbSection, 'Password')

    sTmp = config.get('Datasets', 'Exclude')
    if sTmp <> "":
        global excludeList
        excludeList = str.split(sTmp,',')

    #print "Source dir = '%s'" % (dirSource)
    #print "Archives dir = '%s'" % (dirArchives)
    #
    #print "dbServer = '%s'" % (dbServer)
    #print "dbDatabase = '%s'" % (dbDatabase)
    #print "dbUser = '%s'" % (dbUser)
    #print "dbPassword = '%s'" % (dbPassword)
    #
    #print "Excluding %s" % (excludeList)
    ##for exclude in excludeList:
    ##    print exclude


def listModules():

    for name, module in sorted(sys.modules.items()): 
        if hasattr(module, '__version__'): 
            print name, module.__version__ 


if __name__ == "__main__":
    #try:
    #    src_dir = sys.argv[1]
    #
    #except:
    #    src_dir = os.path.abspath(__file__)
    #    src_dir = os.path.dirname(src_dir)
    #    src_dir = os.path.join(src_dir, "Data")

    loadSettings(os.path.join(os.getcwd(), "ParseXML.ini"))


    #testCode()

    #    import profile
    #    profile.run('main("' + src_dir + '")', src_dir + '/profile.txt')
    #    reviewStats(src_dir + '/profile.txt')
    #    printEx(src_dir)

    main(dirSource, dirArchives)

