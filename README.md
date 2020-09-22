<div align="center">

## Format GetTickCount


</div>

### Description

Format the GetTickCount() API to days, hours, minutes, seconds and miliseconds.

Useful to measure time elapsed between two points.
 
### More Info
 
A long containing the tick count value, and an optional parameter to use different format types.

Ex.:

Msgbox FormatCount(GetTickCount(),1)

A string containing the formated output.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alexandre Moro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alexandre-moro.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alexandre-moro-format-gettickcount__1-3619/archive/master.zip)

### API Declarations

```
Declare Function GetTickCount Lib "Kernel32" () As Long
```


### Source Code

```
Function FormatCount(Count As Long, Optional FormatType As Byte = 0) As String
   Dim Days As Integer, Hours As Long, Minutes As Long, Seconds As Long, Miliseconds As Long
   Miliseconds = Count Mod 1000
   Count = Count \ 1000
   Days = Count \ (24& * 3600&)
   If Days > 0 Then Count = Count - (24& * 3600& * Days)
   Hours = Count \ 3600&
   If Hours > 0 Then Count = Count - (3600& * Hours)
   Minutes = Count \ 60
   Seconds = Count Mod 60
   Select Case FormatType
    Case 0
     FormatCount = Days & " dd, " & Hours & " h, " & _
      Minutes & " min, " & Seconds & " s, " & Miliseconds & _
      " ms"
    Case 1
      FormatCount = Days & " days, " & Hours & " hours, " & _
      Minutes & " minutes, " & Seconds & " seconds, " & Miliseconds & _
      " miliseconds"
    Case 2
      FormatCount = Days & ":" & Hours & ":" & _
      Minutes & ":" & Seconds & ":" & Miliseconds
   End Select
End Function
```

