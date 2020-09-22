Attribute VB_Name = "Module1"
'Header Creator Commands (again i had to whip up.)
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Function cStr2Lng(lString As String) As Long
Dim lJ(0 To 3) As Long, lR As Long, lI As Byte
  If Len(lString) = 4 Then
    For lI = 0 To 3
      lJ(lI) = Asc(Mid(lString, lI + 1, 1))
    Next lI
    lR = lJ(0)
    lR = lR + (lJ(1) * 256)
    lR = lR + (lJ(2) * 65536)
    lR = lR + (lJ(3) * 16777216)
    cStr2Lng = lR
  End If
End Function
Function cLng2Str(lLong As Long) As String
Dim lS As String * 4
  Mid(lS, 1, 1) = Chr(lLong Mod 256)
  Mid(lS, 2, 1) = Chr(Int(lLong / 256) Mod 256)
  Mid(lS, 3, 1) = Chr(Int(lLong / 65536) Mod 256)
  Mid(lS, 4, 1) = Chr(Int(lLong / 16777216) Mod 256)
  cLng2Str = lS
End Function
