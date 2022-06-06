Attribute VB_Name = "M_OperateDateTime"
Option Explicit

#If VBA7 And Win64 Then
    Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

'------------------------------------------------------------------------------
' SYSTEMTIME�\����
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
' ���݂̓��t������Ԃ�
'------------------------------------------------------------------------------
Public Function MyGetSystemTime() As String
    Dim sysTime As SYSTEMTIME
    Dim timeStr As String
    
    '// ���ݓ����擾
    Call GetLocalTime(sysTime)
    
    '// yyyy/mm/dd hh:mm:ss.fff�ɐ��`
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
