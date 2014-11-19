import datetime

def _Property(func):
    return property(**func())

class DataSet(object):

    def __init__(self):
        pass

    class Locations(object):

        def __init__(self):
            self.__dirSource = ''
            self.__dirDest = ''
    
        @_Property
        def dirSource():
            doc = "Source directory with trailing forward slash"
    
            def fget(self):
                return self.__dirSource
    
            def fset(self, value):
                self.__dirSource = value
                if not self.__dirSource.endswith("/"):
                    self.__dirSource = self.__dirSource + '/'
    
            def fdel(self):
                del self.__dirSource
    
            return locals()

        @_Property
        def dirDest():
            doc = "Destination directory with trailing forward slash"
    
            def fget(self):
                return self.__dirDest
    
            def fset(self, value):
                self.__dirDest = value
                if not self.__dirDest.endswith("/"):
                    self.__dirDest = self.__dirDest + '/'
    
            def fdel(self):
                del self.__dirDest
    
            return locals()

    class Info(object):

        def __init__(self):
            self.__zipFile = ''
            self.__fileName = ''
            self.__siteName =''
            self.__siteID = 0
            self.__collectorAbbr = ''
            self.__collectorSN = ''
            self.__collectorMacAddr = ''
            self.__collectorVersion = ''
            self.__collectorOSVersion = ''
            self.__collectorID = 0
            self.__dateFile = datetime.datetime.strptime('12:00:00', '%H:%M:%S')

        @_Property
        def zipFile():
            doc = "Source PKZIP file with extension"
    
            def fget(self):
                return self.__zipFile
    
            def fset(self, value):
                self.__zipFile = value
    
            def fdel(self):
                del self.__zipFile
    
            return locals()

        @_Property
        def fileName():
            doc = "Source PKZIP file without the extension"
    
            def fget(self):
                return self.__fileName
    
            def fset(self, value):
                self.__fileName = value
    
            def fdel(self):
                del self.__fileName
    
            return locals()

        @_Property
        def siteName():
            doc = "Abbreviated name of the site"
    
            def fget(self):
                return self.__siteName
    
            def fset(self, value):
                self.__siteName = value
    
            def fdel(self):
                del self.__siteName
    
            return locals()

        @_Property
        def siteID():
            doc = "Unique identifier (type int) for the site"
    
            def fget(self):
                return self.__siteID
    
            def fset(self, value):
                self.__siteID = value
    
            def fdel(self):
                del self.__siteID
    
            return locals()

        @_Property
        def collectorAbbr():
            doc = "Abbreviation identifying the collector model"
    
            def fget(self):
                return self.__collectorAbbr
    
            def fset(self, value):
                self.__collectorAbbr = value
    
            def fdel(self):
                del self.__collectorAbbr
    
            return locals()

        @_Property
        def collectorSN():
            doc = "Collector serial number"
    
            def fget(self):
                return self.__collectorSN
    
            def fset(self, value):
                self.__collectorSN = value
    
            def fdel(self):
                del self.__collectorSN
    
            return locals()

        @_Property
        def collectorMacAddr():
            doc = "Collector MAC address"
    
            def fget(self):
                return self.__collectorMacAddr
    
            def fset(self, value):
                self.__collectorMacAddr = value
    
            def fdel(self):
                del self.__collectorMacAddr
    
            return locals()

        @_Property
        def collectorVersion():
            doc = "Collector software version"
    
            def fget(self):
                return self.__collectorVersion
    
            def fset(self, value):
                self.__collectorVersion = value
    
            def fdel(self):
                del self.__collectorVersion
    
            return locals()

        @_Property
        def collectorOSVersion():
            doc = "Collector OS version"
    
            def fget(self):
                return self.__collectorVersion
    
            def fset(self, value):
                self.__collectorVersion = value
    
            def fdel(self):
                del self.__collectorVersion
    
            return locals()

        @_Property
        def collectorID():
            doc = "Unique identifier (type int) for the collector"
    
            def fget(self):
                return self.__collectorID
    
            def fset(self, value):
                self.__collectorID = value
    
            def fdel(self):
                del self.__collectorID
    
            return locals()

        @_Property
        def dateFile():
            doc = "Date and time the ZIP file was create (or modified)"
    
            def fget(self):
                return self.__dateFile
    
            def fset(self, value):
                self.__dateFile = value
    
            def fdel(self):
                del self.__dateFile
    
            return locals()

    class RawData(object):

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

        def __init__(self):
            self.__inverterModel = ''
            self.__inverterSN = ''
            self.__inverterID = 0
            self.__datasetAbbr =''
            self.__valFirst = 0.0
            self.__valLast = 0.0
            self.__valMin = 0.0
            self.__valMax = 0.0
            self.__valMean = 0.0
            self.__valBase = 0.0
            self.__period = 0
            self.__dateLogged = datetime.datetime.strptime('12:00:00', '%H:%M:%S')
            self.__dateFile = datetime.datetime.strptime('12:00:00', '%H:%M:%S')

        def resetValues(self):
            self.__init__()
            

        def isValid(self):
            if not self.__inverterModel == '' and \
            not self.__inverterSN == '' and \
            not self.__datasetAbbr == '' and \
            not self.__period == 0 and \
            not self.__dateLogged == '':
                return True;
            else:
                return False;


        @_Property
        def inverterModel():
            doc = "Inverter model number"
    
            def fget(self):
                return self.__inverterModel
    
            def fset(self, value):
                self.__inverterModel = value
    
            def fdel(self):
                del self.__inverterModel
    
            return locals()

        @_Property
        def inverterSN():
            doc = "Inverter serial number"
    
            def fget(self):
                return self.__inverterSN
    
            def fset(self, value):
                self.__inverterSN = value
    
            def fdel(self):
                del self.__inverterSN
    
            return locals()

        @_Property
        def inverterID():
            doc = "Unique identifier (type int) for the inverter"
    
            def fget(self):
                return self.__inverterID
    
            def fset(self, value):
                self.__inverterID = value
    
            def fdel(self):
                del self.__inverterID
    
            return locals()

        @_Property
        def datasetAbbr():
            doc = "Abbreviation of dataset description"
    
            def fget(self):
                return self.__datasetAbbr
    
            def fset(self, value):
                self.__datasetAbbr = value
    
            def fdel(self):
                del self.__datasetAbbr
    
            return locals()

        @_Property
        def valFirst():
            doc = "First value in data sample"
    
            def fget(self):
                return self.__valFirst
    
            def fset(self, value):
                self.__valFirst = value
    
            def fdel(self):
                del self.__valFirst
    
            return locals()

        @_Property
        def valLast():
            doc = "Last value in data sample"
    
            def fget(self):
                return self.__valLast
    
            def fset(self, value):
                self.__valLast = value
    
            def fdel(self):
                del self.__valLast
    
            return locals()

        @_Property
        def valMin():
            doc = "Minimum value in data sample"
    
            def fget(self):
                return self.__valMin
    
            def fset(self, value):
                self.__valMin = value
    
            def fdel(self):
                del self.__valMin
    
            return locals()

        @_Property
        def valMax():
            doc = "Maximum value in data sample"
    
            def fget(self):
                return self.__valMax
    
            def fset(self, value):
                self.__valMax = value
    
            def fdel(self):
                del self.__valMax
    
            return locals()

        @_Property
        def valMean():
            doc = "Meadian value in data sample"
    
            def fget(self):
                return self.__valMean
    
            def fset(self, value):
                self.__valMean = value
    
            def fdel(self):
                del self.__valMean
    
            return locals()

        @_Property
        def valBase():
            doc = "Base value in data sample"
    
            def fget(self):
                return self.__valBase
    
            def fset(self, value):
                self.__valBase = value
    
            def fdel(self):
                del self.__valBase
    
            return locals()

        @_Property
        def period():
            doc = "Period, in minutes, sample was collected"
    
            def fget(self):
                return self.__period
    
            def fset(self, value):
                self.__period = value
    
            def fdel(self):
                del self.__period
    
            return locals()

        @_Property
        def dateLogged():
            doc = "Date and time the sample was collected"
    
            def fget(self):
                return self.__dateLogged
    
            def fset(self, value):
                self.__dateLogged = value
    
            def fdel(self):
                del self.__dateLogged
    
            return locals()

        @_Property
        def dateFile():
            doc = "Date and time file was created"
    
            def fget(self):
                return self.__dateFile
    
            def fset(self, value):
                self.__dateFile = value
    
            def fdel(self):
                del self.__dateFile
    
            return locals()
