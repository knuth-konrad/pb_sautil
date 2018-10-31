#Compile Exe
#Dim All

#Debug Error On
#Tools Off

#Include "saUtilCC.inc"
#Include "ShlObj.inc"


Function PBMain () As Long

   Trace New ".\saUtilTest.tra"
   Trace On

   BlankLine

   ' Function CreateTempFileName(szPath As AsciiZ, szPrefix As AsciiZ, szFileExtension As AsciiZ) As String
   Con.StdOut "-- CreateTempFileName"
   Con.StdOut CreateTempFileName(".\DATA", "sa", "")

   Local i, lCount As Long
   lCount = ArgC()
   Con.StdOut "Command$, ArgC: " & Command$ & ", " & Format$(lCount)

   For i = 1 To lCount
      Con.StdOut Format$(i, "00) ") & ArgV(i)
   Next i

   BlankLine


   ' Function FormatNumber(ByVal xNumber As Ext) Common As String
   Con.StdOut "-- FormatNumber"
   Con.StdOut FormatNumber(12345.678)

   BlankLine


   ' Function FormatNumberEx(ByVal xNumber As Ext, Optional lOmitDecimals As Long) Common As String
   Con.StdOut "-- FormatNumberEx"
   Con.StdOut FormatNumberEx(12345.678)
   Con.StdOut FormatNumberEx(12345.678, %True)

   BlankLine


   ' Function DateNowLocal(Optional ByVal lLocale As Long, Optional ByVal lDateFormat As Long) Common As String
   Con.StdOut "-- DateNowLocal"
   Con.StdOut DateNowLocal()
   Con.StdOut DateNowLocal(0, %DATE_SHORTDATE)

   BlankLine


   ' Function TimeNowLocal(Optional ByVal lLocale As Long) Common As String
   Con.StdOut "-- TimeNowLocal"
   Con.StdOut TimeNowLocal()

   BlankLine


   ' Function CreateNestedDirs(ByVal sPath As String) Common As Long
   Con.StdOut "-- CreateNestedDirs"
   Con.StdOut Format$(CreateNestedDirs("C:\DATA\PB\INCLUDE\INCKGK\saUtil\DATA\1\2"))
   Con.StdOut Format$(CreateNestedDirs(".\DATA\1\2\3"))

   BlankLine


   ' Various date & time methods
   Con.StdOut "-- Various date & time methods"
   Con.StdOut "- DateDMY: " & DateDMY()
   Con.StdOut "- DateTimeDMY: " & DateTimeDMY()
   Con.StdOut "- DateYMD: " & DateYMD()
   Con.StdOut "- DateTimeDMY: " & DateTimeYMD()
   Con.StdOut "- DateTimeYMDinUTC: " & DateTimeYMDinUTC()

   BlankLine


   ' Function FilesCount(ByVal sPath As String, Optional ByVal vntMask As Variant) Common As Dword
   Con.StdOut "-- FilesCount"
   Con.StdOut Format$(FilesCount("C:\Windows", "*.dll"))

   BlankLine


   ' Function baGetFileSize(ByRef sFileName As String) Common As Quad
   ' Function baGetAllFileSize(ByVal sPath As String, Optional ByVal vntFilePattern As Variant) Common As Quad
   Con.StdOut "-- baGetFileSize / baGetAllFileSize"
   Con.StdOut FormatNumberEx(baGetFileSize("C:\DATA\PB\INCLUDE\INCKGK\saUtil\DATA\sabb62.tmp"), %True)
   Con.StdOut FormatNumberEx(baGetAllFileSize("C:\Windows", "*.dll"), %True)

   BlankLine


'   Function Delete2RecycleBin(ByVal sFile As String, _
'      Optional ByVal lShowProgress As Long, Optional ByVal lConfirmation As Long, _
'      Optional ByVal lSimple As Long, Optional ByVal lSysErrors As Long, _
'      Optional ByVal hWndParent As Long) Common As Long
   Con.StdOut "-- Delete2RecycleBin"
   Con.StdOut Format$(Delete2RecycleBin("C:\DATA\PB\INCLUDE\INCKGK\saUtil\DATA\sabb62.tmp", _
      %True, _
      %True))

   BlankLine


   Trace Off
   Trace Close

End Function

Sub BlankLine()
   Con.StdOut ""
End Sub
