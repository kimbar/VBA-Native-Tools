Attribute VB_Name = "TestUtil_General"
'@Folder "Tests"

Option Explicit
Option Private Module

Public Function RoundSigFig(ByVal val As Double, Optional sf As Long = 3) As String
    ' Rounds the output to `sf` significant figures
    Dim l10 As Double
    Dim neg As Double: neg = 1#
    If val = 0 Then
        RoundSigFig = "0"
        Exit Function
    ElseIf val < 0 Then
        val = -val
        neg = -1#
    End If
    l10 = (10 ^ Int(Log(val) / 2.30258509299405))
    RoundSigFig = CStr(CDbl(Left$(CStr(val / l10), sf + 1)) * l10 * neg)
End Function

Public Sub LFSR_16Bits(ByRef x As Long, ByVal times As Long)
    ' Linear-Feedback Shift Register pseudo-random data generation as seen in
    ' <https://commons.wikimedia.org/w/index.php?title=File:LFSR-F16.svg&oldid=462029164>
    Dim idx As Long
    For idx = 1 To times
        x = (x \ 2) Or ((((x And &H1) <> 0&) Xor ((x And &H4) <> 0&) Xor ((x And &H8) <> 0&) Xor ((x And &H20) <> 0&)) And &H8000&)
    Next
End Sub
