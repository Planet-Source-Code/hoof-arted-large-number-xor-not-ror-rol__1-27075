<div align="center">

## Large Number XOR / NOT / ROR / ROL


</div>

### Description

Use this code to do various byte manipulation on very large numbers. A function for RoleOver (ROL/ROR) (One function for both directions). Also includes a large number NOT and XOR. If you try doing a normal XOR, i.e.4294967295 XOR (Anyvalue), the system overflows. This function prevents that.
 
### More Info
 
Public Function largeXor(ByVal firstdouble As Double, seconddouble As Double) As Double

Public Function largenot(ByVal firstdouble As Double) As Double

Public Function RoleOver(ByVal bigdbl As Double, Dir As Integer, Count As Integer) As Double

DIR = Direction (1 = ROL 2 = ROR)

Count = Position Shift

Copy and place in Module

All functions return Double to work with other functions.

Make sure that results are worked with correctly to avoid further overflows.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hoof Arted](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hoof-arted.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hoof-arted-large-number-xor-not-ror-rol__1-27075/archive/master.zip)





### Source Code

```
'I have created these functions using some functions from other people. The LargeXor and NOT
'are my functions and so is the RoleOver.
'Visual Basic cannot do these things with large number so I wrote these functions.
'The RoleOver function works 100% correctly as do the others.
'
'Thanks for using my code. Thanks to the others who wrote the additional functions
'Hoof Arted
Public Function Hex2Bin(ByVal sHex As String) As String
  Dim i As Integer
  Dim j As Integer
  Dim nDec As Long
  Const HexChar As String = "0123456789ABCDEF"
  For i = 1 To Len(sHex)
    nDec = InStr(1, HexChar, Mid(sHex, i, 1)) - 1
    For j = 3 To 0 Step -1
      Hex2Bin = Hex2Bin & nDec \ 2 ^ j
      nDec = nDec Mod 2 ^ j
    Next j
  Next i
  'Remove the first unused 0
  i = InStr(1, Hex2Bin, "1")
  If i <> 0 Then Hex2Bin = Mid(Hex2Bin, i)
End Function
Public Function Bin2Hex(ByVal sBin As String) As String
  Dim i As Integer
  Dim nDec As Double
  sBin = String(4 - Len(sBin) Mod 4, "0") & sBin 'Add zero To complete Byte
  For i = 1 To Len(sBin)
    nDec = nDec + CInt(Mid(sBin, Len(sBin) - i + 1, 1)) * 2 ^ (i - 1)
  Next i
  Bin2Hex = nDec 'Hex(nDec)
  If Len(Bin2Hex) Mod 2 = 1 Then Bin2Hex = "0" & Bin2Hex
End Function
Public Function BigDecToHex(ByVal DecNum) As String
  ' This function is 100% accurate untill
  '   15,000,000,000,000,000 (1.5E+16)
  Dim NextHexDigit As Double
  Dim HexNum As String
  HexNum = ""
  While DecNum <> 0
  NextHexDigit = DecNum - (Int(DecNum / 16) * 16)
  If NextHexDigit < 10 Then
    HexNum = Chr(Asc(NextHexDigit)) & HexNum
  Else
    HexNum = Chr(Asc("A") + NextHexDigit - 10) & HexNum
  End If
  DecNum = Int(DecNum / 16)
Wend
If HexNum = "" Then HexNum = "0"
BigDecToHex = HexNum
End Function
Public Function RoleOver(ByVal bigdbl As Double, Dir As Integer, Count As Integer) As Double
Dim origionalbin As String
Dim Modbin As String
HEXSTR = BigDecToHex(bigdbl)
origionalbin = Hex2Bin(HEXSTR)
If Len(origionalbin) < 32 Then origionalbin = Mid$("00000000000000000000000000000000", 1, 32 - Len(origionalbin)) & origionalbin
If Dir = 2 Then Modbin = Right$(origionalbin, Count) + Left$(origionalbin, Len(origionalbin) - Count)
If Dir = 1 Then Modbin = Right$(origionalbin, Len(origionalbin) - Count) + Left$(origionalbin, Count)
RoleOver = Val(Bin2Hex(Modbin))
End Function
Public Function largeXor(ByVal firstdouble As Double, seconddouble As Double) As Double
Dim Position As Integer
Dim resultbin As String
resultbin = "00000000000000000000000000000000"
Firsthex = BigDecToHex(firstdouble)
secondhex = BigDecToHex(seconddouble)
firstbin = Hex2Bin(Firsthex)
If Len(firstbin) < 32 Then firstbin = Mid$("00000000000000000000000000000000", 1, 32 - Len(firstbin)) & firstbin
Secondbin = Hex2Bin(secondhex)
If Len(Secondbin) < 32 Then Secondbin = Mid$("00000000000000000000000000000000", 1, 32 - Len(Secondbin)) & Secondbin
' XOR logic
'0 0  0
'0 1  1
'1 0  1
'1 1  0
For Position = 1 To 32
If Mid$(firstbin, Position, 1) = "0" And Mid$(Secondbin, Position, 1) = "0" Then Mid$(resultbin, Position, 1) = "0"
If Mid$(firstbin, Position, 1) = "0" And Mid$(Secondbin, Position, 1) = "1" Then Mid$(resultbin, Position, 1) = "1"
If Mid$(firstbin, Position, 1) = "1" And Mid$(Secondbin, Position, 1) = "0" Then Mid$(resultbin, Position, 1) = "1"
If Mid$(firstbin, Position, 1) = "1" And Mid$(Secondbin, Position, 1) = "1" Then Mid$(resultbin, Position, 1) = "0"
Next Position
largeXor = Val(Bin2Hex(resultbin))
End Function
Public Function largenot(ByVal firstdouble As Double) As Double
Dim Position As Integer
Dim resultbin As String
resultbin = "00000000000000000000000000000000"
Firsthex = BigDecToHex(firstdouble)
firstbin = Hex2Bin(Firsthex)
If Len(firstbin) < 32 Then firstbin = Mid$("00000000000000000000000000000000", 1, 32 - Len(firstbin)) & firstbin
'NOT logic
'1 = 0
'0 = 1
For Position = 1 To 32
If Mid$(firstbin, Position, 1) = "1" Then Mid$(resultbin, Position, 1) = "0"
If Mid$(firstbin, Position, 1) = "0" Then Mid$(resultbin, Position, 1) = "1"
Next Position
largenot = Val(Bin2Hex(resultbin))
End Function
```

