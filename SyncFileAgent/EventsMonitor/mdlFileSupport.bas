Attribute VB_Name = "FileSupport"
Option Explicit

Public Type FILE_INFO
    FileName As String
    FilePath As String
    'FileSize As Currency
    FileSize As Double
    FileTimeCreated As Date
    FileTimeLastModified As Date
    FileTimeLastAccessed As Date
    FileAttributes As Long
End Type
