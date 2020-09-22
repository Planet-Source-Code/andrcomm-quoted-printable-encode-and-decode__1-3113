<div align="center">

## Quoted\-Printable \-\- Encode and Decode


</div>

### Description

Very fast function to encode or decode Quoted-Printable.

VB6 only, but you can make it work with other versions, with a 3rd party replace function.
 
### More Info
 
Just pass it the string to be encoded, or to be decoded.

The encoded, or decoded string.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[AndrComm](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrcomm.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrcomm-quoted-printable-encode-and-decode__1-3113/archive/master.zip)





### Source Code

```
Public Function DecodeQP(ByRef StrToDecode As String) As String
Dim sTemp As String
Dim i As Integer
sTemp = StrToDecode
For i = 255 To 127 Step -1
  If InStr(1, sTemp, "=" & Hex$(i)) <> 0 Then sTemp = Replace(sTemp, "=" & Hex$(i), Chr$(i))
Next
  If InStr(1, sTemp, "=" & Hex$(61)) <> 0 Then sTemp = Replace(sTemp, "=" & Hex$(61), Chr$(255) & Chr$(254))
For i = 32 To 10 Step -1
  If InStr(1, sTemp, "=" & Hex$(i)) <> 0 Then sTemp = Replace(sTemp, "=" & Hex$(i), Chr$(i))
Next
For i = 9 To 0 Step -1
  If InStr(1, sTemp, "=" & "0" & Hex$(i)) <> 0 Then sTemp = Replace(sTemp, "=" & Hex$(i), Chr$(i))
Next
sTemp = Replace(sTemp, "=", "")
sTemp = Replace(sTemp, Chr$(255) & Chr$(254), "=")
DecodeQP = sTemp
End Function
Public Function EncodeQP(ByRef StrToEncode As String) As String
Dim sTemp As String
Dim i As Integer
sTemp = StrToEncode
For i = 255 To 127 Step -1
  If InStr(1, sTemp, Chr$(i)) <> 0 Then sTemp = Replace(sTemp, Chr$(i), "=" & Hex$(i))
Next
  If InStr(1, sTemp, Chr$(61)) <> 0 Then sTemp = Replace(sTemp, Chr$(61), "=" & Hex$(61))
For i = 32 To 10 Step -1
  If InStr(1, sTemp, Chr$(i)) <> 0 Then sTemp = Replace(sTemp, Chr$(i), "=" & Hex$(i))
Next
For i = 9 To 0 Step -1
  If InStr(1, sTemp, Chr$(i)) <> 0 Then sTemp = Replace(sTemp, Chr$(i), "=" & "0" & Hex$(i))
Next
EncodeQP = sTemp
End Function
```

