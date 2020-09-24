Attribute VB_Name = "modPCM"
' Code module and project created by "Urthman"
'   http://www.jsent.biz/urthman/
'   http://www.mp3.com/urthman/

Option Explicit

Global Const DT1& = 350
Global Const DT2& = 440

Global Const BT1& = 480
Global Const BT2& = 620

Global Const RT1& = 440
Global Const RT2& = 480


Global WavFileName As String

Private Enum OFStyle
    fRead& = &H0
    fWrite& = &H1
    fReadWrite& = &H2
    fCreate& = &H1000
    fExist& = &H4000
End Enum

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName As String * 128
End Type

Private Declare Function OpenFile& Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As OFStyle)

Dim NullBuff As OFSTRUCT

Function FileExist(WhatFile$) As Boolean

    FileExist = (OpenFile(WhatFile, NullBuff, fExist) > 0)

End Function
Sub Main()

    WavFileName = App.Path
    If (Mid(WavFileName, Len(WavFileName), 1) = "\") Then
        WavFileName = (WavFileName & "PcmRiff.WAV")
    Else
        WavFileName = (WavFileName & "\PcmRiff.WAV")
    End If
    
    Load frmPcm
    
    With frmPcm
        .CheckRate
        .Show
        .SetFocus
    End With

End Sub
