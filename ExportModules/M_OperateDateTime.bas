Attribute VB_Name = "M_OperateDateTime"
Option Explicit

#If VBA7 And Win64 Then
    Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

'------------------------------------------------------------------------------
' SYSTEMTIME構造体
'------------------------------------------------------------------------------
Type SYSTEMTIME
    sYear As Integer
    sMonth As Integer
    sDayOfWeek As Integer
    sDay As Integer
    sHour As Integer
    sMinute As Integer
    sSecond As Integer
    sMilliseconds As Integer
End Type

'------------------------------------------------------------------------------
' 現在の日付時刻を返す
'------------------------------------------------------------------------------
Public Function MyGetSystemTime() As String
    Dim sysTime As SYSTEMTIME
    Dim timeStr As String
    
    '// 現在日時取得
    Call GetLocalTime(sysTime)
    
    '// yyyy/mm/dd hh:mm:ss.fffに整形
    timeStr = Format(sysTime.sYear, "0000")
    timeStr = timeStr & "/"
    timeStr = timeStr & Format(sysTime.sMonth, "00")
    timeStr = timeStr & "/"
    timeStr = timeStr & Format(sysTime.sDay, "00")
    timeStr = timeStr & " "
    timeStr = timeStr & Format(sysTime.sHour, "00")
    timeStr = timeStr & ":"
    timeStr = timeStr & Format(sysTime.sMinute, "00")
    timeStr = timeStr & ":"
    timeStr = timeStr & Format(sysTime.sSecond, "00")
    timeStr = timeStr & "."
    timeStr = timeStr & Format(sysTime.sMilliseconds, "000")
    
    MyGetSystemTime = timeStr
End Function

