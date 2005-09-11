VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private LocalSystemTime As SYSTEMTIME

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Private Function ConvertSYSTEMTIME(systime As SYSTEMTIME) As Date
    ConvertSYSTEMTIME = DateSerial(systime.wYear, systime.wMonth, systime.wDay) + _
        TimeSerial(systime.wHour, systime.wMinute, systime.wSecond)
End Function

Public Sub Class_Initialize()
    Me.LocalTime = Now
End Sub

Property Let LocalTime(dt As Date)
    LocalSystemTime.wYear = Year(dt)
    LocalSystemTime.wMonth = Month(dt)
    LocalSystemTime.wDay = Day(dt)
    LocalSystemTime.wHour = Hour(dt)
    LocalSystemTime.wMinute = Minute(dt)
    LocalSystemTime.wSecond = Second(dt)
End Property

Property Get LocalTime() As Date
    LocalTime = ConvertSYSTEMTIME(LocalSystemTime)
End Property

Property Get UTCTime() As Date
    Dim utc_fTime As FILETIME
    Dim loc_fTime As FILETIME
    Dim utc_sTime As SYSTEMTIME

    'Convert to UTC file time, and then system time
    SystemTimeToFileTime LocalSystemTime, loc_fTime
    LocalFileTimeToFileTime loc_fTime, utc_fTime
    FileTimeToSystemTime utc_fTime, utc_sTime
    UTCTime = ConvertSYSTEMTIME(utc_sTime)
End Property
