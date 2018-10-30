#Compile SLL "..\bin\saUtil.sll"
#Include Once "Win32Api.inc"
'------------------------------------------------------------------------------
'Purpose  : General purpose helper routines
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
#If Not %Def(%SAUTIL)
   %SAUTIL = 1      'To avoid dupe #Includes
#EndIf

%SA_READFILE_BLOCKSIZE = 4096&   ' Block size for method ReadFile
%SA_BLOCKSIZE = 8192             ' Block size for method FindInFile

' File access test
%famReadShared = 1
%famReadExclusive = 2
%famWriteShared = 3
%famWriteExclusive = 4
%famReadWriteShared = 5
%famReadWriteExclusive = 6
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
Union PBFileTime
   FT As FILETIME
   q As Quad
End Union

Union PBSystemTime
   ST As FILETIME
   q As Quad
End Union
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include Once "Windows.inc"
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
'==============================================================================

Function GetEXECompletePath() Common As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve fully qualified path to running EXE, including EXE file name
'
'Prereq.  : -
'Parameter: -
'Returns  : Fully qualified path of EXE
'Note     : -
'
'   Author: Knuth Konrad, 22.06.1999
'   Source: -
'  Changed: 23.01.2018
'           - Original code obsolete with new PB object "EXE". Function
'           now exists for backward compatibility
'------------------------------------------------------------------------------
   Function = Exe.Full$
End Function
'==============================================================================

Function GetEXEName() Common As String
'------------------------------------------------------------------------------
'Purpose  : Determine the name of running executable.
'
'Prereq.  : -
'Parameter: -
'Returns  : EXE file name incl. file extension, w/o path
'Note     : -
'
'   Author: Knuth Konrad, 22.06.1999
'   Source: -
'  Changed: 23.01.2018
'           - Original code obsolete with new PB object "EXE". Function
'           now exists for backward compatibility
'------------------------------------------------------------------------------
   Function = Exe.Namex$
End Function
'==============================================================================

Function GetEXEPath() Common As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve complete path to the running EXE
'
'Prereq.  : -
'Parameter: Fully qualified path incl. trailing backslash
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad, 22.06.1999
'   Source: -
'  Changed: 23.01.2018
'           - Original code obsolete with new PB object "EXE". Function
'           now exists for backward compatibility
'------------------------------------------------------------------------------
   Function = NormalizePath(Exe.Path$)
End Function
'==============================================================================

Function AppPath() Common As String
'------------------------------------------------------------------------------
'Purpose  : Returns the path to the executable, similar to VB's App.Path()
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local szPath As AsciiZ * %Max_Path

   GetModuleFileName GetModuleHandle(ByVal %NULL), szPath, %Max_Path
   If InStr(-1, szPath, "\") > 0 Then
      AppPath = Left$(szPath, InStr( -1, szPath, "\"))
   End If

End Function
'==============================================================================

Function NormalizePath(ByVal sPath As String, Optional ByVal vDelim As Variant) Common As String
'------------------------------------------------------------------------------
'Purpose  : Ensures that the passed directory ends with a directory separator
'
'Prereq.  : -
'Parameter: sPath    - Path/directory excluding file name
'           vDelim   - Directory/folder separator, defaults to "\"
'
'Returns  : Path including trailing delimiter
'Note     : -
'
'   Author: Bruce McKinney
'   Source: Hardcore Visual Basic 5
'  Changed: - 09.08.1999  Knuth Konrad
'           Optional path separator
'------------------------------------------------------------------------------
   Local sDelim As String

   If IsMissing(vDelim) Then
      sDelim = "\"
   Else
      sDelim = Variant$(vDelim)
   End If

   If Right$(sPath, Len(sDelim)) <> sDelim Then
      NormalizePath = sPath & sDelim
   Else
      NormalizePath = sPath
   End If

End Function
'==============================================================================

Function DenormalizePath(ByVal sPath As String, Optional ByVal vDelim As Variant) Common As String
'------------------------------------------------------------------------------
'Purpose  : Ensures that the passed directory does NOT end with a directory separator
'
'Prereq.  : -
'Parameter: sPath    - Path/directory excluding file name
'           vDelim   - Directory/folder separator, defaults to "\"
'
'Returns  : Path excluding trailing delimiter
'Note     : -
'
'   Author: Bruce McKinney
'   Source: Hardcore Visual Basic 5
'  Changed: - 09.08.1999  Knuth Konrad
'           Optional path separator
'------------------------------------------------------------------------------
   Local sDelim As String

   If IsMissing(vDelim) Then
      sDelim = "\"
   Else
      sDelim = Variant$(vDelim)
   End If

   If Right$(sPath, Len(sDelim)) = sDelim Then
      DenormalizePath = Left$(sPath, Len(sPath) - Len(sDelim))
   Else
      DenormalizePath = sPath
   End If

End Function
'==============================================================================

Function FileExist(szFile As AsciiZ) Common As Long
'------------------------------------------------------------------------------
'Purpose  : Checks if a file exists
'
'Prereq.  : -
'Parameter: szFile   - File to check, incl. full path
'Returns  : True/False  - File exists/doesn't exists
'Note     : -
'
'   Author: Knuth Konrad 17.08.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local udtWin32FFD As WIN32_FIND_DATA
   Local dwdRetval As Dword

   dwdRetval = FindFirstFile(szFile, udtWin32FFD)
   If dwdRetval = %INVALID_HANDLE_VALUE Then
      FileExist = %FALSE
      Exit Function
   Else
      dwdRetval = FindClose(dwdRetval)
      FileExist = %TRUE
   End If

End Function
'==============================================================================

Function DirExist(ByVal sDirectory As String) Common As Long
'------------------------------------------------------------------------------
'Purpose  : Checks if a diretory exists
'
'Prereq.  : -
'Parameter: Full path to directory
'Returns  : True/False
'Note     : This method returns True for folders with attributes Hidden, ReadOnly
'           and System
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local udtWin32FFD As WIN32_FIND_DATA
   Local lRetval As Long
   Local szTemp As AsciiZ * %Max_Path

   ' Ensure we're dealing with the folder name only
   szTemp = DenormalizePath(sDirectory)

   If Len(Trim$(szTemp)) < 1 Then
   ' "Empty"/"no" folders always exist
      DirExist = %TRUE
      Exit Function
   End If

   lRetval = FindFirstFile(szTemp, udtWin32FFD)
   If lRetval = %INVALID_HANDLE_VALUE Then
      DirExist = %False
      Exit Function
   Else
      If udtWin32FFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY Then
         DirExist = %True
      Else
         DirExist = %False
      End If
      lRetval = FindClose(lRetval)
   End If

End Function
'==============================================================================

Function GetFile(ByVal sFile As String) Common As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve a file in chunks from disk
'
'Prereq.  : -
'Parameter: sFile - (Full path to) File to retrieve
'Returns  : Fully qualified path of EXE
'Note     : -
'
'   Author: Knuth Konrad, 06.04.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local i, hFile As Long
   Local dwRemain, dwCurPos, dwBlockSize, dwFileSize As Dword
   Local sBuffer, sContent As String

   ' File not found -> outa here
   If Len(Dir$(sFile)) < 1 Then
      Exit Function
   End If

   dwBlockSize = %SA_READFILE_BLOCKSIZE
   sContent = ""

   hFile = FreeFile
   Open sFile For Binary As #hFile
   'OPEN szSource$ FOR BINARY ACCESS READ LOCK SHARED AS #hSource%

   dwFilesize = Lof(hFile)

   If Lof(hFile) > dwBlockSize Then
   ' Do this if the file size/remaining part is > dwBlockSize

      For i = 1 To Lof(hFile) \ dwBlockSize
         Get$ #hFile, dwBlockSize, sBuffer
         sContent = sContent & sBuffer
      Next i

      ' Get the remaining partial block

      Get$ #hFile, Lof(hFile) - ((Lof(hFile) \ dwBlockSize) * dwBlockSize), sBuffer
      sContent = sContent & sBuffer

   Else
   ' or do this if the file is less than lBlocksize&
      Get$ #hFile, Lof(hFile), sBuffer
      sContent = sContent & sBuffer
   End If

   Close #hFile
   GetFile = sContent

End Function
'==============================================================================

Function GetFilePartial(ByVal sFile As String, ByVal dwStartPos As Dword, ByVal sMatch As String, ByVal  _
   dwByte As Dword) Common As String
'------------------------------------------------------------------------------
'Purpose  : Retrieves parts of a file
'
'Prereq.  : -
'Parameter: sFile       - (Full path to) File to retrieve
'           dwStartPos  - Start retrieving here
'Returns  : Partial file contents
'Note     : -
'
'   Author: Knuth Konrad, 06.04.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local i, hFile As Long
   Local dwRemain, dwCurPos, dwBlockSize, dwFileSize As Dword
   Local sBuffer, sContent As String

   ' Safe guard
   If Len(Dir$(sFile)) < 1 Then
      Exit Function
   End If

   dwBlockSize = %SA_READFILE_BLOCKSIZE
   sContent = ""

   hFile = FreeFile
   Open sFile For Binary As #hFile

   If dwStartPos > 0 Then
      Seek #hFile, dwStartPos
   End If

   dwFilesize = Lof(hFile)

   If Lof(hFile) > dwBlockSize Then
   ' Do this if the file size/remaining part is > dwBlockSize

      For i = 1 To Lof(hFile) \ dwBlockSize
         Get$ #hFile, dwBlockSize, sBuffer
         sContent = sContent & sBuffer
      Next i

      ' Get the remaining partial block

      Get$ #hFile, Lof(hFile) - ((Lof(hFile) \ dwBlockSize) * dwBlockSize), sBuffer
      sContent = sContent & sBuffer

   Else
   ' or do this if the file is less than lBlocksize&
      Get$ #hFile, Lof(hFile), sBuffer
      sContent = sContent & sBuffer
   End If

   Close #hFile
   GetFilePartial = sContent

End Function
'==============================================================================

Function CreateTempFileName(szPath As AsciiZ, szPrefix As AsciiZ, _
   szFileExtension As AsciiZ) Common As String
'------------------------------------------------------------------------------
'Purpose  : Creates a temporary file name
'
'Prereq.  : -
'Parameter: szPath            - Folder in which the temp. file should be create.
'                               If empty, create in current folder.
'           szPrefix          - Create the temp. file name with this prefix
'           szFileExtension   - Create the temp. file with this extension.
'                               Pass the extension WITHOUT the leading ".", e.g. "myextension"
'Returns  : A non-existing file name for temp usage with szPath included,
'           e.g. .\data\sa123.tmp
'Note     : This method does NOT create a file. It simply generates a file name
'           in the folder passed and ensures that no file with a similar name
'           exists in the target directory
'
'   Author: Knuth Konrad, 26.08.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sFile As String
   Local szTemp As AsciiZ * %Max_Path
   Local lRetval As Long

   On Error GoTo CreateTempFileNameError

   If Len(szPath) = 0 Then szPath = ".\"

   szTemp = Space$(%Max_Path)
   ' Create a temp. file with Win32 API ...
   lRetval = GetTempFileName(szPath, szPrefix, 0&, szTemp)
   If lRetval <> 0 Then
      sFile = LCase$(Remove$(szTemp, Chr$(0)))
      '... since Windows actually creates a 0 byte file, delete it.
      Kill sFile
   Else
      ' Something went wrong ...
      Function = ""
      Exit Function
   End If

   ' Win API GetTempFileName always creates the file with the extension TMP
   ' -> check if a different extension should be used
   If Len(szFileExtension) > 0 And LCase$(szFileExtension) <> "tmp" Then

      Replace ".tmp" With "." & szFileExtension In sFile
      ' Now check for the existance of that file (name)
      Do While FileExist(ByCopy sFile) = %True
         sFile = CreateTempFileName(szPath, szPrefix, szFileExtension)
      Loop
   End If

   CreateTempFileName = sFile

   CreateTempFileNameExit:
   On Error GoTo 0
   Exit Function

   CreateTempFileNameError:
   ErrClear
   CreateTempFileName = ""
   Resume CreateTempFileNameExit

End Function
'==============================================================================

Function ShortenPath(ByVal sLongPath As String, ByVal lLenght As Long) Common As String
'------------------------------------------------------------------------------
'Purpose  : Shortens a folder name (for display purposes) and inserts the
'           common "..." into the name
'
'Prereq.  : -
'Parameter: sLongPath   - Folder name
'           lLenght     - max string length for the name
'Returns  : Returns a shortened version of the name or the cople name if it's < lLength
'Note     : -
'
'   Author: Knuth Konrad, 13.09.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sFilename, sTemp As String
   Local lRetval, lLen As Long

   ' Safe guard
   If Len(sLongPath) < = lLenght Then
      ShortenPath = sLongPath
      Exit Function
   End If

   ' If a file name is included, make sure that it alone
   ' doesn't exceed the max. length. If it does, return an empty string
   lRetval = InStr(-1, sLongPath, "\")
   If lRetval > 0 Then
      sFilename = Right$(sLongPath, Len(sLongPath) - lRetval)
      If Len(sFilename) > lLenght Then
         ShortenPath = ""
         Exit Function
      Else
         lLen = Len(sFilename)
         sTemp = Left$(sLongPath, lRetval)
      End If
   End If

   ShortenPath = Left$(sTemp,(Len(sTemp) - (lLen + 3)) / 2) & "..." & Right$(sTemp,(Len(sTemp) _
      - (lLen + 3)) / 2) & sFilename

End Function
'==============================================================================

Sub WriteINI(ByVal sFile As String, ByVal sSection As String, _
   ByVal sKey As String, ByVal sData As String) Common
'------------------------------------------------------------------------------
'Purpose  : Writing a value to a INI file (wrapper)
'
'Prereq.  : -
'Parameter: sFile    - INI file (fully qualified)
'           sSection - INI section
'           sKey     - INI key
'           sData    - Value to store
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad, 13.09.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local szFile As AsciiZ * 300
   Local szSection As AsciiZ * 300
   Local szKey As AsciiZ * 300
   Local szData As AsciiZ * 1000

   szFile = Trim$(sFile)
   szSection = Trim$(sSection)
   szKey = Trim$(sKey)
   szData = Trim$(sData)
   WritePrivateProfileString szSection, szKey, szData, szFile

End Sub
'==============================================================================

Function GetFromINIString(ByVal sSection As String, ByVal sKey As String, ByVal sDefault As String,  _
   ByVal sINIFile As String ) Common As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve a string value fom a INI
'
'Prereq.  : -
'Parameter: sSection - INI section
'           sKey     - INI key
'           sDefault - Default value.
'           sINIFile - INI file (fully qualified)
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad, 13.09.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local hFile As Long
   Local sTemp As String

   hFile = FreeFile
   Open sINIFile For Input Access Read Lock Shared As #hFile

   While Not Eof(hFile)
      Line Input #hFile, sTemp
      If Trim$(sTemp) = Trim$("[" & sSection & "]") Then
         While Not Eof(hFile)
            Line Input #hFile, sTemp
            If Left$(sTemp, Len(sKey) + 1) = sKey & "=" Then
               GetFromINIString = Remain$(sTemp, sKey & "=")
               Close #hFile
               Exit Function
            End If
         Wend
      End If
   Wend

   Close #hFile
   Function = sDefault

End Function
'==============================================================================

Function GetFromINIInt(ByVal sSection As String, ByVal sKey As String, ByVal lDefault As Long, _
   ByVal sINIFile As String ) Common As Long
'------------------------------------------------------------------------------
'Purpose  : Retrieve a integer value fom a INI
'
'Prereq.  : -
'Parameter: sSection - INI section
'           sKey     - INI key
'           lDefault - Default value.
'           sINIFile - INI file (fully qualified)
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad, 13.09.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local hFile As Long
   Local sTemp As String

   hFile = FreeFile
   Open sINIFile For Input Access Read Lock Shared As #hFile

   While Not Eof(hFile)
      Line Input #hFile, sTemp
      If Trim$(sTemp) = Trim$("[" & sSection & "]") Then
         While Not Eof(hFile)
            Line Input #hFile, sTemp
            If Left$(sTemp, Len(sKey) + 1) = sKey & "=" Then
               GetFromINIInt = Val(Remain$(sTemp, sKey & "="))
               Close #hFile
               Exit Function
            End If
         Wend
      End If
   Wend

   Close #hFile
   Function = lDefault

End Function
'==============================================================================

Sub SetHighPriority() Common
'------------------------------------------------------------------------------
'Purpose  : Set this process' priority to "High"
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : Wrapper for SetProcessPriority
'
'   Author: Knuth Konrad, 13.09.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local hProcess As Long
   Local lRetval As Long

   hProcess = GetCurrentProcess()
   lRetval = SetPriorityClass( hProcess, %High_Priority_Class )

End Sub
'==============================================================================

Function SetProcessPriority(ByVal dwPriority As Dword) Common As Long
'------------------------------------------------------------------------------
'Purpose  : Set this process' priority
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : Never sets priority to %REALTIME_PRIORITY_CLASS as this would
'           make the system unusable
'
'   Author: Knuth Konrad, 13.09.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local hProcess As Long
   Local lRetval As Long

   ' Constants for SetPriorityClass
   '%NORMAL_PRIORITY_CLASS             = &H00000020
   '%IDLE_PRIORITY_CLASS               = &H00000040
   '%HIGH_PRIORITY_CLASS               = &H00000080
   '%REALTIME_PRIORITY_CLASS           = &H00000100
   '%BELOW_NORMAL_PRIORITY_CLASS       = &H00004000
   '%ABOVE_NORMAL_PRIORITY_CLASS       = &H00008000

   hProcess = GetCurrentProcess()

   ' We don't allow RealTime ...
   Select Case dwPriority
   Case %Normal_Priority_Class, %Idle_Priority_Class, %High_Priority_Class, _
      %BELOW_NORMAL_PRIORITY_CLASS, %ABOVE_NORMAL_PRIORITY_CLASS

      lRetval = SetPriorityClass(hProcess, dwPriority)
      SetProcessPriority = %True

   Case Else
      SetProcessPriority = %False

   End Select

End Function
'==============================================================================

Sub PBDoEvents() Common
'------------------------------------------------------------------------------
'Purpose  : Emulate the VB Classic's DoEvents, i.e. allow other processes/threads
'           to process their messages
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad, 13.09.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Static Msg As tagMsg

   If PeekMessage( Msg, %NULL, 0, 0, %PM_REMOVE ) Then
      TranslateMessage Msg
      DispatchMessage Msg
   End If

End Sub
'==============================================================================

Function GetExtPos(ByVal sSpec As String) Common As Long
'------------------------------------------------------------------------------
'Purpose  : Determines the position of a file's extension (including the '.')
'           e.g. "12345.ext" -> 5
'
'Prereq.  : -
'Parameter: sSpec - File name
'Returns  : -
'Note     : -
'
'   Author: Bruce McKinney, Hardcore Visual Basic 5
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lLast, lExt As Long

   lLast = Len(sSpec)

   ' Parse backward to find extension or base
   For lExt = lLast + 1 To 1 Step - 1
      Select Case Mid$(sSpec, lExt, 1)
      Case "."
           ' First . from right is extension start
         Exit For
      Case "\"
           ' First \ from right is base start
         lExt = lLast + 1
         Exit For
      End Select
   Next

   ' Negative return indicates no extension, but this
   ' is base so callers don't have to reparse.
   Function = lExt

End Function
'==============================================================================

Function ExtractFileName(ByVal sFullname As String, Optional ByVal vntOmitExtension As Variant, _
   Optional ByVal vntDelim As Variant) Common As String
'------------------------------------------------------------------------------
'Purpose  : Extracts the file name (incl. extension) from a fully qualified path
'           e.g. "\\MyShare\Folder1\Folder2\12345.ext" -> 12345.ext (or 12345)
'
'Prereq.  : -
'Parameter: sFullname         - Full path
'           vntDelim          - Path/folder separator (defaults to "\")
'           vntOmitExtension  - True = return file name only
'Returns  : Extracted file name
'Note     : -
'
'   Author: Knuth Konrad, 21.06.2000
'   Source: -
'  Changed: 05.12.2017
'           - New parameters: vntDelim, vntOmitExtension
'------------------------------------------------------------------------------
   Local lSeperator, lOmitExtension As Long
   Local sDelim, sResult As String

   Trace On

   ' Default to "\" as the path seperator
   If IsMissing(vntDelim) Then
      sDelim = "\"
   Else
      sDelim = Variant$(vntDelim)
   End If
   Trace Print " - (saUtil:ExtractFileName), sDelim: " & sDelim

   ' Default to full file name (incl. extension)
   If IsMissing(vntOmitExtension) Then
      lOmitExtension = 0
   Else
      lOmitExtension = Variant#(vntOmitExtension)
   End If
   Trace Print " - (saUtil:ExtractFileName), lOmitExtension: " & Format$(lOmitExtension)

   ' If there's no seperator, the whole "thing" is the file name
   lSeperator = InStr(-1, sFullname, sDelim)

   If lSeperator > 0 Then
      sResult = Mid$(sFullname, lSeperator + 1)
   Else
      sResult = sFullname
   End If
   Trace Print " - (saUtil:ExtractFileName), sResult: " & sResult

   If IsTrue(lOmitExtension) Then
      ' Get rid off the file extension
      Local sTemp As String
      sTemp = ExtractExtensionName(sResult)
      sResult = Left$(sResult, Len(sResult) - Len(sTemp))
   End If

   ExtractFileName = sResult

End Function
'==============================================================================

Function ExtractPath(ByVal sFullname As String, Optional ByVal vntDelim As Variant) Common As String
'------------------------------------------------------------------------------
'Purpose  : Extracts the path (incl. drive letter) from a fully qualified path
'           *without'* the trailng path seperator
'           e.g. "\\MyShare\Folder1\Folder2\12345.ext" -> \\MyShare\Folder1\Folder2
'
'Prereq.  : -
'Parameter: sFullname         - Full path
'           vntDelim          - Path/folder separator (defaults to "\")
'Returns  : Extracted path
'Note     : -
'
'   Author: Knuth Konrad, 29.03.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lSeperator As Long
   Local sDelim As String

   If IsMissing(vntDelim) Then
      sDelim = "\"
   Else
      sDelim = Variant$(vntDelim)
   End If

   lSeperator = InStr(-1, sFullname, sDelim)

   If lSeperator > 0 Then
      ExtractPath = Left$(sFullname, lSeperator - 1)
   Else
      ExtractPath = sFullname
   End If

End Function
'==============================================================================

Function ExtractExtensionName(ByVal sFullname As String) Common As String
'------------------------------------------------------------------------------
'Purpose  : Extracts the file extension (incl. ".") from from a fully qualified path
'           e.g. "\\MyShare\Folder1\Folder2\12345.ext" -> .ext
'
'Prereq.  : -
'Parameter: sFullname         - Full path
'Returns  : Extracted file extension
'Note     : -
'
'   Author: Knuth Konrad, 21.06.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lDot As Long

   lDot = InStr(-1, sFullname, ".")

   If lDot > 0 Then
      ExtractExtensionName = Right$(sFullname, Len(sFullname) - (lDot - 1))
   Else
      ExtractExtensionName = "."
   End If

End Function
'==============================================================================

Function ArgC(Optional ByVal vntCmd As Variant) Common As Long
'------------------------------------------------------------------------------
'Purpose  : Returns (counts) the number of command line parameters
'
'Prereq.  : -
'Parameter: vntCmd   - Command line
'Returns  : Number of parameters
'Note     : -
'
'   Author: Knuth Konrad, 21.06.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local arg As Long
   Local f   As String
   Local q   As Long
   Local cmd As String

   If IsMissing(vntCmd) Then
      cmd = Command$
   Else
      cmd = Variant$(vntCmd)
   End If

   Do While Len(cmd)
      Incr arg
      f = Left$(cmd, 1)
      If Asc(f) = 34 Then
         q = InStr(Mid$(cmd ,2), $Dq)
         If q Then
            f = Left$(cmd, q+1)
         Else
            f = cmd
         End If
      Else
         f = f + Extract$(Mid$(cmd,2), Any $Dq+" /")
      End If
      cmd = LTrim$(Mid$(cmd, Len(f)+1))
   Loop

   Function = arg

End Function
'==============================================================================

Function ArgV(ByVal lWhich As Long) Common As String
'------------------------------------------------------------------------------
'Purpose  : Returns the <lWhich) parameter (key and value),
'           e.g. for "/a=123 /bcd=xyz", ArgV(2) returns "/bcd=xyz"
'
'Prereq.  : -
'Parameter: lWhich   - Which key/value parameter pair to return
'Returns  : key/value pair
'Note     : -
'
'   Author: Knuth Konrad, 21.06.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local arg As Long
   Local f   As String
   Local q   As Long
   Local cmd As String

   cmd = Command$

   Do While Len(cmd)
      Incr arg
      f = Left$(cmd, 1)
      If Asc(f) = 34 Then
         q = InStr(Mid$(cmd, 2), $Dq)
         If q Then
            f = Left$(cmd, q+1)
         Else
            f = cmd
         End If
      Else
         f = f + Extract$(Mid$(cmd, 2), Any $Dq+" /")
      End If
      cmd = LTrim$(Mid$(cmd, Len(f) +1))

      If arg = lWhich Then
         Exit Do
      Else
         f = ""
      End If
   Loop

   Function = f

End Function
'==============================================================================

Function StrIncr (ByVal sString As String) Common As String
'------------------------------------------------------------------------------
'Purpose  : Increments an alphanumerical string by "1",
'           e.g. 101Z -> 102A
'
'Prereq.  : -
'Parameter: sString  - String to increment
'Returns  : String incremented by "1"
'Note     : -
'
'   Author: Dave Navarro, Jr., Last Revision: July 15, 1994
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lValue, x As Long
   Local sChar As String

   ' Numers and '9's only -> just add 1
   If Tally(sString, "9") = Len(sString) Then
      lValue = Val(sString)
      Incr lValue
      Function = Trim$(Str$(lValue))
      Exit Function
   End If

   For x = Len(sString) To 1 Step - 1

      sChar = Mid$(sString, x, 1)

      If (sChar >= "0") And (sChar <= "8") Then
         sChar = Chr$(Asc(sChar) + 1)
         Mid$(sString, x, 1) = sChar
         Exit For
      ElseIf sChar = "9" Then
         sChar = "0"
         Mid$(sString, x, 1) = sChar
      ElseIf (sChar >= "A") And (sChar <= "Y") Then
         sChar = Chr$(Asc(sChar) + 1)
         Mid$(sString, x, 1) = sChar
         Exit For
      ElseIf sChar = "Z" Then
         sChar = "A"
         Mid$(sString, x, 1) = sChar
      ElseIf (sChar >= "a") And (sChar <= "y") Then
         sChar = Chr$(Asc(sChar) + 1)
         Mid$(sString, x, 1) = sChar
         Exit For
      ElseIf sChar = "z" Then
         sChar = "a"
         Mid$(sString, x, 1) = sChar
      End If

   Next x

   Function = sString

End Function
'==============================================================================

Function BackupFile(ByVal sFileSource As String, ByRef sFileDest As String, _
   Optional ByVal bolCopyOnly As Long, Optional vntIncrementTarget As Variant) Common As Long
'------------------------------------------------------------------------------
'Purpose  : Creates a backup of a file by copying or moving it from folder <b> to
'           folder <b>
'
'Prereq.  : -
'Parameter: sFileSource          - Fully qualified source file
'           sFileDest            - Fully qualified destination file name
'           bolCopyOnly          - Copy (1) or move (0 = Default)?
'           vntIncrementTarget   - If destination file already exists, create a new
'                                  new copy using the pattern <file>.<nnnn>.<ext>
'                                  Default = True (1)
'Returns  : %True / %False (Success / Failure)
'Note     : If the destination file already exists and vntIncrementTarget = False, then
'           the existing (destination) file will be overwritten
'
'   Author: Knuth Konrad 14.11.2012
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local i, bolIncrementTarget, lRetry As Long
   Local sDestPath, sDestFile, sDestExt, sTempFile As String

   ' Safe guard. Does source exist?
   If IsFalse(FileExist(ByCopy sFileSource)) Then
      BackupFile = %TRUE
      Exit Function
   End If

   If Not IsMissing(vntIncrementTarget) Then
      bolIncrementTarget = Variant#(vntIncrementTarget)
   Else
      bolIncrementTarget = %True
   End If

   ' Does destination exist?
   If FileExist(ByCopy sFileDest) Then
   ' Yes -> Split the destination into its parts (path, file name, extension)
      sDestPath = PathName$(Path, sFileDest)
      sDestFile = PathName$(Name, sFileDest)
      sDestExt = PathName$(Extn, sFileDest)
      If IsTrue(bolIncrementTarget) Then
         ' Create a copy of the destination with a different name. We only try up to 9999
         Do
            sTempFile = sDestPath & sDestFile & "." & Format$(i, "0000") & sDestExt
            i = i + 1
         Loop Until IsFalse(FileExist(ByCopy sTempFile)) Or (i > 9999)
      Else
         sTempFile = sFileDest
      End If
   Else
      sTempFile = sFileDest
   End If

   sFileDest = sTempFile


   ' Do the actual copy/move operation.
   ' Tries this 3 times, with a 500ms delay between operations to
   ' mitigate potential locking conflicts
   lRetry = 0
   Do While lRetry < 3

      Try
         If IsTrue(bolCopyOnly) Then
            FileCopy sFileSource, sTempFile
         Else
            Name sFileSource As sTempFile
         End If
         BackupFile = %True
      Catch
         BackupFile = %False
      End Try

      If Err Then
         Sleep 500
         ErrClear
         Incr lRetry
      Else
         Exit Loop
      End If

   Loop

End Function
'==============================================================================

Function RenameFileExt(ByVal sFullFileName As String, Optional ByVal vntFileExtChar As Variant, _
   Optional ByVal bolAppendOnly As Long) Common As Long
'------------------------------------------------------------------------------
'Purpose  : Rename a file by replacing the file extension's last character with vntFileExtChar
'           folder <b>
'
'Prereq.  : -
'Parameter: sFullFilename  - File to be renamed
'           sFileExtChar   - Replacing character
'           bolAppendOnly  - True = append sFileExtChar to file extension
'                            False = replace last character (default)
'Returns  : %True / %False (Success / Failure)
'Note     : If a file with the newly created file name already exists in that
'           location, try to create a different file name by adding an incremental number
'           (up to 999).
'           If all those filesa also exist, simply overwrite the first (initially created)
'           file.
'
'   Author: Knuth Konrad 14.11.2012
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sPath, sFile, sExt, sNewName As String
   Local sFileExtChar As String
   Local lFileIndex As Long

   ' Safe guard
   If IsFalse(FileExist(ByCopy sFullFileName)) Then
      RenameFileExt = %True
      Exit Function
   End If

   If IsMissing(vntFileExtChar) Then
      sFileExtChar = "_"
   Else
      sFileExtChar = Variant$(vntFileExtChar)
   End If

   ' Split the file name into its parts
   sPath = PathName$(Path, sFullFileName)
   sFile = PathName$(Name, sFullFileName)
   sExt = PathName$(Extn, sFullFileName)

   ' Nur anhängen?Just append?
   If IsTrue(bolAppendOnly) Then
      ' Yes
      sExt = sExt & sFileExtChar
   Else
      ' No
      ' Replace last character of extension
      If Len(sExt) > 1 Then
         Mid$(sExt, Len(sExt), 1) = sFileExtChar
      Else
         sExt = sExt & sFileExtChar
      End If

   End If

   ' Add new extension to file name
   If Len(sFile) > 1 Then
      sFile = Left$(sFile, GetExtPos(sFile) - 1)
   End If

   sNewName = sPath & sFile & sExt

   ' If a similar file (name) already exists, try to create a copy
   ' by adding an incremental part to the file name, but only do this
   ' up to 1000 times (= 1000 different file names)
   Do While IsTrue(FileExist(ByCopy sNewName)) And lFileIndex < 1000

      sNewName = sPath & sFile & "." & Format$(lFileIndex, "000") & sExt
      Incr lFileIndex

   Loop

   Try
      If lFileIndex < 1000 Then
      ' There's an non-existing one we could used
         Name sFullFileName As sNewName
      Else
      ' More than 999 files with similar name exist -> Kill the file with the originally created
      ' new file name and replace it with our copy
         Name sFullFileName As sPath & sFile & sExt
      End If
      RenameFileExt = %True
   Catch
      RenameFileExt = %False
   End Try

End Function
'==============================================================================

Function FormatNumber(ByVal xNumber As Ext) As String
'------------------------------------------------------------------------------
'Purpose  : Formats a number with proper user locale settings
'
'Prereq.  : -
'Parameter: xNumber  - number to format
'Returns  : Formatted numer string
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lpzInputValue  As AsciiZ * 22  '18 digits, leading zero, optional leading minus, and decimal point.
   Local lpzOutputValue As AsciiZ * 40  'additional room provided for commas, etc.

   lpzInputValue = LTrim$(Str$(xNumber, 18))

   Call GetNumberFormat(%LOCALE_USER_DEFAULT, ByVal 0, lpzInputValue, ByVal 0, lpzOutputValue, ByVal 40)

   Function = lpzOutputValue

End Function
'==============================================================================

Function FormatNumberEx(ByVal xNumber As Ext, Optional lOmitDecimals As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Format a number according to the current user's local
'
'Prereq.  : -
'Parameter: -
'Returns  : xNumber  - Number to format
'Note     : -
'
'   Author: Knuth Konrad, 25.10.2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lpzInputValue  As WStringZ * 44  '18 digits, leading zero, optional leading minus, and decimal point.
   Local lpzOutputValue As WStringZ * 80  'additional room provided for commas, etc.
   Local lRet As Long

   lpzInputValue = LTrim$(Str$(xNumber, 18))

'   stdout "lpzInputValue: " & lpzInputValue
'   lRet = GetNumberFormatEx(byval %LOCALE_NAME_USER_DEFAULT, ByVal 0, lpzInputValue, ByVal 0, lpzOutputValue, sizeof(lpzOutputValue))
'   StdOut "GetNumberFormatEx: " & Format$(lRet)
'   stdout "lpzOutputValue: " & lpzOutputValue
'   StdOut "Left$(lpzOutputValue, lRet): " & left$(lpzOutputValue, lRet)
'
'   local i as long
'   for i = 1 to lRet
'      stdout format$(asc(lpzOutputValue, i)) & " ";
'   next

   Call GetNumberFormatEx(ByVal %LOCALE_NAME_USER_DEFAULT, ByVal 0, lpzInputValue, ByVal 0, lpzOutputValue, SizeOf(lpzOutputValue))

   If IsTrue(lOmitDecimals) Then
      ' Drop the decimals and decimal separator
      Local szDecimal As AsciiZ * 2
      Local lRetVal As Long

      lRetVal = GetLocaleInfo(%LOCALE_SYSTEM_DEFAULT, %LOCALE_SDECIMAL, szDecimal, SizeOf(szDecimal))
      Function = Extract$(lpzOutputValue, szDecimal)

   Else

      Function = lpzOutputValue

   End If

End Function
'==============================================================================

Function FormatLoc(ByVal fextValue As Ext, ByVal sMask As String) As String
'------------------------------------------------------------------------------
'Does     : Format a number and taking Locale settings for decimal and
'           thousander separator into account
'
'Called From: -
'Requires : -
'Parameter: -
'Returns  : Formatted String
'Note     : -
'
'   Author: Knuth Konrad
'  created: 18.07.2001
'  changed: -
'------------------------------------------------------------------------------
   Local szThousand As AsciiZ * 2
   Local szDecimal As AsciiZ * 2
   Local lRetVal As Long
   Local sResult As String

   lRetVal = GetLocaleInfo(%LOCALE_SYSTEM_DEFAULT, %LOCALE_SDECIMAL, szDecimal, SizeOf(szDecimal))
   lRetVal = GetLocaleInfo(%LOCALE_SYSTEM_DEFAULT, %LOCALE_STHOUSAND, szThousand, SizeOf(szThousand))

   If (szDecimal = ".") And (szThousand = ",") Then
   'The Locale settings match the PB settings -> return the FORMAT$-string
      FormatLoc = Format$(fextValue, sMask)
   Else
   'The Locale settings differ -> Swap around
      sResult = Format$(fextValue, sMask)
      Replace Any ",." With ".," In sResult
      FormatLoc = sResult
   End If

End Function
'----------------------------------------------------------------------------

Function DateNowLocal() As String

   Local lRetval As Long, udtST As SYSTEMTIME
   Local szDate As AsciiZ * %Max_Path

   GetLocalTime udtST

   szDate = ""
   lRetval = GetDateFormat(%LOCALE_SYSTEM_DEFAULT, _
      %DATE_LONGDATE, _
      udtST, _
      "", _
      szDate, _
      SizeOf(szDate))
   DateNowLocal = Left$(szDate, lRetval)

End Function
'===========================================================================

Function TimeNowLocal() As String

   Local lRetval As Long, udtST As SYSTEMTIME
   Local szTime As AsciiZ * %Max_Path

   GetLocalTime udtST

   szTime = ""
   lRetval = GetTimeFormat(%LOCALE_SYSTEM_DEFAULT, _
      0, _
      udtST, _
      "", _
      szTime, _
      SizeOf(szTime))
   TimeNowLocal = Left$(szTime, lRetval)

End Function
'===========================================================================

Function CreateDir(ByVal sDir As String) As Long

   On Error Resume Next

   Local sPath As String, sLabel As String

   sLabel = sDir
   sPath = Left$(sLabel, 3)
   sLabel = Right$(sLabel, Len(sLabel) - 3)
   If Right$(sLabel, 1) <> "\" Then sLabel = sLabel & "\"
   sPath = sPath & Mid$(sLabel, 1, InStr(sLabel, "\") - 1)
   sLabel = Right$(sLabel, Len(sLabel) - InStr(sLabel, "\"))

   While Right$(sPath, 1) <> "\"
      MkDir sPath
      If sLabel <> "" Then
         sPath = sPath & "\" & Mid$(sLabel, 1, InStr(sLabel, "\") - 1)
      Else
         sPath = sPath & "\"
      End If
      sLabel = Right$(sLabel, Len(sLabel) - InStr(sLabel, "\"))
   Wend

End Function
'===========================================================================
'
'Function AppendStr2( ByVal stPos   As Long, _
'                     ByRef sBuffer As String, _
'                     ByVal Addon   As Long, _
'                     ByVal lenAdd  As Long _
'                     ) Export As Long
'
'' Quelle: http://www.powerbasic.com/support/pbforums/showthread.php?t=35836
''//
''//  Fast function for concatenating many strings
''//
'
'
'    #Register None
'
'    Local pBuffer As Long
'
'    ' If the buffer is not large enough to handle the adding
'    ' of this string then we need to expand the buffer.
'    If stPos + lenAdd + 1 > Len(sBuffer) Then
'       sBuffer = sBuffer & String$( Max&(lenAdd, 100 * 1024), 0)   ' increase 100K minimum
'    End If
'
'    ' Copy the new string to the end of the buffer
'    pBuffer = StrPtr(sBuffer)
'
'    ! cld               ; Read forwards
'
'    ! mov edi, pBuffer  ; Put buffer address In edi
'    ! Add edi, stPos    ; Add starting offset To it
'
'    ! mov esi, Addon    ; Put String address In esi
'    ! mov ecx, lenAdd   ; length In ecx As counter
'
'    ! rep movsb         ; Copy ecx count Of bytes From esi To edi
'
'    ! mov edx, stPos
'    ! Add edx, lenAdd   ; Add stPos And lenAdd For Return value
'
'    ! mov Function, edx
'
'End Function
'===========================================================================

Function FindInFile( ByVal sFile As String, ByVal sPattern As String, ByVal dwStartPos As Dword ) As  _
   Dword

   Register hFile As Long, i As Long, j As Long
   Local dwBlockSize As Dword, dwBlocks As Dword, dwCurPos As Dword, dwFound As Dword, dwRemain As Dword
   Local sTemp As String
   Dim asChunk(0) As String

   'Datei nicht vorhanden -> raus hier
   If Len( Dir$(sFile)) < 1 Then Exit Function

   hFile = FreeFile
   Open sFile For Binary As #hFile

   'Suchausdruck ist größer als die gesamte Datei -> raus hier
   If Len(sPattern) > Lof(hFile) Then
      Close #hFile
      Exit Function
   End If

   'Positionszeiger auf die Startposition setzen
   If dwStartPos < 1 Then dwStartPos = 1
   Seek #hFile, dwStartPos

   'Chunk-Größe ermitteln
   If Lof(hFile) - dwStartPos < %SA_BLOCKSIZE Then
      dwBlockSize = Lof(hFile) - dwStartPos
   Else
      dwBlockSize = %SA_BLOCKSIZE
   End If

   'Wieviele Chunks müssen wir lesen damit wir die ganze Datei haben?
   dwBlocks = (Lof(hFile) - dwStartPos) \ dwBlockSize
   'Und wieviel Rest bleibt dann da noch
   dwRemain = (Lof(hFile) - dwStartPos) - (dwBlocks * dwBlockSize)

   'Wir merken uns die Postion in der Datei
   dwCurPos = dwStartPos - 1

   'Datei auslesen...
   For i = 1 To dwBlocks
      If UBound(asChunk) * dwBlockSize < Len(sPattern) Then
         ReDim Preserve asChunk(0 To UBound(asChunk) + 1)
      End If
      'Block holen...
      Get$ #hFile, dwBlockSize, asChunk(UBound(asChunk))
      'Wir merken uns die Postion in der Datei
      dwCurPos = dwCurPos + dwBlockSize
      'Wenn nun die Anzahl Bytes die wir gelesen haben >= der Länge des Suchstrings ist,
      'überprüfen ob der Suchstring darin enthalten ist.
      If UBound( asChunk ) * dwBlockSize > = Len( sPattern ) Then
         'Temporären String zusammenbauen, den wir mit dem Suchstring vergleichen können
         sTemp = ""
         For j = 1 To UBound( asChunk )
            sTemp = sTemp & asChunk( j )
         Next j
         dwFound = InStr( sTemp, sPattern )
         ' Wir haben unseren Suchstring gefunden -> Position in Datei berechnen und raus hier
         If IsTrue( dwFound ) Then
            FindInFile = dwCurPos - Len( sTemp ) + dwFound
            Close #hFile
            Exit Function
         'Nichts gefunden -> 1. Array-Element rausschmeißen und den Rest nachrücken
         Else
            Array Delete asChunk( 1 )
         End If
      End If
   Next i

   'Den Rest der Datei noch dazuholen
   If IsTrue( dwRemain ) Then
      If UBound( asChunk ) * dwBlockSize < Len( sPattern ) Then
         ReDim Preserve asChunk( 0 To UBound( asChunk ) + 1 )
      End If
      'Block holen...
      Get$ #hFile, dwRemain, asChunk( UBound( asChunk ))
      'Wir merken uns die Postion in der Datei
      dwCurPos = dwCurPos + dwRemain
      'Wenn nun die Anzahl Bytes die wir gelesen haben >= der Länge des Suchstrings ist,
      'überprüfen ob der Suchstring darin enthalten ist.
      If UBound( asChunk ) * dwBlockSize > = Len( sPattern ) Then
         'Temporären String zusammenbauen, den wir mit dem Suchstring vergleichen können
         sTemp = ""
         For j = 1 To UBound( asChunk )
            sTemp = sTemp & asChunk( j )
         Next j
         dwFound = InStr( sTemp, sPattern )
         'Wir haben unseren Suchstring gefunden -> Position in Datei berechnen
         If IsTrue( dwFound ) Then
            FindInFile = dwCurPos - Len( sTemp ) + dwFound
         End If
      End If
   End If

   Close #hFile

End Function
'===========================================================================

Function DateDMY(Optional ByVal vntDelim As Variant) As String

   Local sDelim As String
   Local o As IPowerTime

   Let o = Class "PowerTime"
   o.Now

   If IsMissing(vntDelim) Then
      sDelim = "-"
   Else
      sDelim = Variant$(vntDelim)
   End If

   Function = Format$(o.Day, "00") & sDelim & _
      Format$(o.Month, "00") & sDelim & _
      Format$(o.Year, "0000")

End Function
'==============================================================================

Function DateTimeDMY(Optional ByVal vntDelim As Variant) As String

   Local sDelim As String
   Local o As IPowerTime

   Let o = Class "PowerTime"
   o.Now

   If IsMissing(vntDelim) Then
      sDelim = "-"
   Else
      sDelim = Variant$(vntDelim)
   End If

   Function = Format$(o.Day, "00") & sDelim & _
      Format$(o.Month, "00") & sDelim & _
      Format$(o.Year, "0000") & "T" & _
      Format$(o.Hour, "00") & ":" & _
      Format$(o.Minute, "00") & ":" & _
      Format$(o.Second, "00") & "." & _
      Format$(o.MSecond)

End Function
'==============================================================================

Function DateYMD(Optional ByVal vntDelim As Variant) As String

   Local sDelim As String
   Local o As IPowerTime

   Let o = Class "PowerTime"
   o.Now

   If IsMissing(vntDelim) Then
      sDelim = "-"
   Else
      sDelim = Variant$(vntDelim)
   End If

   Function = Format$(o.Year, "0000") & sDelim & _
      Format$(o.Month, "00") & sDelim & _
      Format$(o.Day, "00")

End Function
'==============================================================================

Function DateTimeYMD(Optional ByVal vntDelim As Variant) As String

   Local sDelim As String
   Local o As IPowerTime

   Let o = Class "PowerTime"
   o.Now

   If IsMissing(vntDelim) Then
      sDelim = "-"
   Else
      sDelim = Variant$(vntDelim)
   End If

   Function = Format$(o.Year, "0000") & sDelim & _
      Format$(o.Month, "00") & sDelim & _
      Format$(o.Day, "00") & "T" & _
      Format$(o.Hour, "00") & ":" & _
      Format$(o.Minute, "00") & ":" & _
      Format$(o.Second, "00") & "." & _
      Format$(o.MSecond)

End Function
'==============================================================================

Function DateTimeYMDinUTC(Optional ByVal vntDelim As Variant) As String

   Local sDelim As String
   Local o As IPowerTime

   Let o = Class "PowerTime"
   o.NowUTC

   If IsMissing(vntDelim) Then
      sDelim = "-"
   Else
      sDelim = Variant$(vntDelim)
   End If

   Function = Format$(o.Year, "0000") & sDelim & _
      Format$(o.Month, "00") & sDelim & _
      Format$(o.Day, "00") & "T" & _
      Format$(o.Hour, "00") & _
      Format$(o.Minute, "00") & _
      Format$(o.Second, "00") & "." & _
      Format$(o.MSecond, "0000")

End Function
'==============================================================================

Function FormatDate(szFormatMask As AsciiZ, Optional ByVal vntDate As Variant) As String
'------------------------------------------------------------------------------
'Purpose  : Returns a formatted date string
'
'Prereq.  : -
'Parameter: sFormatMask - Desired format, equivalent to the formats allowed by
'                         Windwos API GetDateFormat:
' d      - Day of month as digits with no leading zero for single-digit days.
' dd     - Day of month as digits with leading zero for single-digit days.
' ddd    - Day of week as a three-letter abbreviation. The function uses the LOCALE_SABBREVDAYNAME value associated with the specified locale.
' dddd   - Day of week as its full name. The function uses the LOCALE_SDAYNAME value associated with the specified locale.
' M      - Month as digits with no leading zero for single-digit months.
' MM     - Month as digits with leading zero for single-digit months.
' MMM    - Month as a three-letter abbreviation. The function uses the LOCALE_SABBREVMONTHNAME value associated with the specified locale.
' MMMM   - Month as its full name. The function uses the LOCALE_SMONTHNAME value associated with the specified locale.
' y      - Year as last two digits, but with no leading zero for years less than 10.
' yy     - Year as last two digits, but with leading zero for years less than 10.
' yyyy   - Year represented by full four digits.
' gg     - Period/era string. The function uses the CAL_SERASTRING value associated with the specified locale. This element is ignored if the date
'           to be formatted does not have an associated era or period string.
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 02.05.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local szResult As AsciiZ * %Max_Path, lRetVal As Long
   Local oDate As IPowerTime, udtST As SYSTEMTIME, uniFT As PBFileTime

   If IsMissing(vntDate) Then
      Let oDate = Class "PowerTime"
      oDate.Now
   ElseIf VariantVT(vntDate) <> %VT_Date Then
      Let oDate = Class "PowerTime"
      oDate.Now
   Else
      Let oDate = vntDate
   End If

   szResult = ""
   uniFT.q = oDate.FileTime
   Call FileTimeToSystemTime(uniFT.FT, udtST)

   lRetVal = GetDateFormat(ByVal 0&, ByVal 0&, udtST, szFormatMask, szResult, SizeOf(szResult))
   If IsTrue(lRetVal) Then
      FormatDate = Left$(szResult, lRetVal)
   End If

End Function
'==============================================================================

Function FormatTime(szFormatMask As AsciiZ, Optional ByVal vntTime As Variant) As String
'------------------------------------------------------------------------------
'Purpose  : Returns a formatted time string
'
'Prereq.  : -
'Parameter: sFormatMask - Desired format, equivalent to the formats allowed by
'                         Windwos API GetTimeFormat:
' h   - Hours with no leading zero for single-digit hours; 12-hour clock.
' hh  - Hours with leading zero for single-digit hours; 12-hour clock.
' H   - Hours with no leading zero for single-digit hours; 24-hour clock.
' HH  - Hours with leading zero for single-digit hours; 24-hour clock.
' m   - Minutes with no leading zero for single-digit minutes.
' mm  - Minutes with leading zero for single-digit minutes.
' s   - Seconds with no leading zero for single-digit seconds.
' ss  - Seconds with leading zero for single-digit seconds.
' t   - One character time-marker string, such as A or P.
' tt  - Multicharacter time-marker string, such as AM or PM.
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 02.05.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local szResult As AsciiZ * %Max_Path, lRetVal As Long
   Local oDate As IPowerTime, udtST As SYSTEMTIME, uniFT As PBFileTime

   If IsMissing(vntTime) Then
      Let oDate = Class "PowerTime"
      oDate.Now
   ElseIf VariantVT(vntTime) <> %VT_Date Then
      Let oDate = Class "PowerTime"
      oDate.Now
   Else
      Let oDate = vntTime
   End If

   szResult = ""
   uniFT.q = oDate.FileTime
   Call FileTimeToSystemTime(uniFT.FT, udtST)

   lRetVal = GetTimeFormat(ByVal 0&, ByVal 0&, udtST, szFormatMask, szResult, SizeOf(szResult))
   If IsTrue(lRetVal) Then
      FormatTime = Left$(szResult, lRetVal)
   'Else
   '   StdOut FuncName$ & ": " & Format$(lRetval)
   End If

End Function
'==============================================================================

Function FormatDateTime(szDateMask As AsciiZ, szTimeMask As AsciiZ, Optional ByVal vntDate As Variant, Optional ByVal vntTime As Variant) As String
'------------------------------------------------------------------------------
'Purpose  : Returns a formatted date/time string
'
'Prereq.  : -
'Parameter: szXxxxMask   - Desired format, equivalent to the formats allowed by
'                         Windwos API GetDateFormat, GetTimeFormat
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 02.05.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sResultDate, sResultTime As String

   If Not IsMissing(vntDate) Then
      sResultDate = FormatDate(szDateMask, vntDate)
   Else
      sResultDate = FormatDate(szDateMask)
   End If

   If Not IsMissing(vntTime) Then
      sResultTime = FormatTime(szTimeMask, vntTime)
   Else
      sResultTime = FormatTime(szTimeMask)
   End If

   If Len(sResultTime) > 0 Then
      FormatDateTime = sResultDate & " " & sResultTime
   Else
      FormatDateTime = sResultDate & sResultTime
   End If

End Function
'==============================================================================

Function PBNow() As IPowerTime

   Local o As IPowerTime

   Let o = Class "PowerTime"
   o.Now

   Let PBNow = o

End Function
'==============================================================================

' Create a formatted system-message
'
Function WinErrMsg (ByVal dError As Dword) As String

    Local pBuffer   As WStringZ Ptr
    Local ncbBuffer As Dword

    ncbBuffer = FormatMessageW( _
                    %FORMAT_MESSAGE_ALLOCATE_BUFFER _
                 Or %FORMAT_MESSAGE_FROM_SYSTEM _
                 Or %FORMAT_MESSAGE_IGNORE_INSERTS, _
                    ByVal %NULL, _
                    dError, _
                    ByVal MAKELANGID(%LANG_NEUTRAL, %SUBLANG_DEFAULT), _
                    ByVal VarPtr(pBuffer), _
                    0, _
                    ByVal %NULL)

    If ncbBuffer Then
        Function = Peek$$(pBuffer, ncbBuffer)
        LocalFree pBuffer
    Else
        Function = "Unknown error code: &H" + Hex$(dError, 8)
    End If

End Function
'==============================================================================

Function FilesCount(ByVal sPath As String, Optional ByVal vntMask As Variant) As Dword
'------------------------------------------------------------------------------
'Funktion : Zählt die Anzahl der Dateien in einem Verzeichnis
'
'Vorauss. : -
'Parameter: sPath -  Verzeichnis das die Dateien beinhaltet
'           sMask -  Dateimaske für Suche
'Rückgabe : Anzahl Dateien
'Notiz    : -
'
'    Autor: Knuth Konrad 19.08.2004
' geändert: -
'------------------------------------------------------------------------------
   Local dwCount As Dword
   Local sSearch, sMask As String

   Local sTemp As String

   sPath = NormalizePath(sPath)

   If IsMissing(vntMask) Then
      sMask = "*.*"
   Else
      sMask = Variant$(vntMask)
   End If

   sSearch = sPath & sMask

   sTemp = Dir$(sSearch)
   If Len(Trim$(sTemp)) < 1 Then
      FilesCount = 0
      Exit Function
   End If

   Do While Len(Trim$(sTemp)) > 0
      dwCount = dwCount + 1
      sTemp = Dir$
   Loop

   FilesCount = dwCount

End Function
'==============================================================================

Function FileAccess(ByVal sFile As String, ByVal lAccessMode As Long) As Long
'------------------------------------------------------------------------------
'Purpose  : Test a file for access modes
'
'Prereq.  : -
'Parameter: sFile       - File in question incl. full path
'           lAccessMode - Access mode:
'           %famReadShared = 1
'           %famReadExclusive = 2
'           %famWriteShared = 3
'           %famWriteExclusive = 4
'           %famReadWriteShared = 5
'           %famReadWriteExclusive = 6
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 17.07.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local hFile As Long

   If Not FileExist(ByCopy sFile) Then
   'Datei gar nicht vorhanden->immer Fehler
      FileAccess = %False
      Exit Function
   End If

   hFile = FreeFile

   Try
      Select Case lAccessMode
      Case %famReadShared
         Open sFile For Input Access Read Shared As #hFile
      Case %famReadExclusive
         Open sFile For Input Access Read Lock Read Write As #hFile
      Case %famWriteShared
         Open sFile For Binary Access Write Lock Shared As #hFile
      Case %famWriteExclusive
         Open sFile For Binary Access Write Lock Read Write As #hFile
      Case %famReadWriteShared
         Open sFile For Append Access Read Write Shared As #hFile
      Case %famReadWriteExclusive
         Open sFile For Append Access Read Write Lock Read Write As #hFile
      Case Else
      End Select

      FileAccess = %True

   Catch
      FileAccess = %False

   End Try

   Close #hFile

End Function
'==============================================================================

Function baGetFileSize(ByRef sFileName As String) As Quad
'------------------------------------------------------------------------------
'Purpose  : Return a file's size
'
'Prereq.  : -
'Parameter: sFileName   - File name incl. full path
'Returns  : File size in bytes
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local W32FD             As WIN32_FIND_DATA
   Local hFile             As Dword

   ' Safeguard
   If Len(sFileName) = 0 Then
      Exit Function
   End If

   hFile = FindFirstFile(ByVal StrPtr(sFileName), W32FD)
   If hFile <> %INVALID_HANDLE_VALUE Then
      Function = W32FD.nFileSizeHigh * &H0100000000 + W32FD.nFileSizeLow
      FindClose hFile
   End If

End Function
'---------------------------------------------------------------------------

Function Delete2RecycleBin(ByVal sFile As String, _
   Optional ByVal lShowProgress As Long, Optional ByVal lConfirmation As Long, _
   Optional ByVal lSimple As Long, Optional ByVal lSysErrors As Long, _
   Optional ByVal hWndParent As Long) As Long
'------------------------------------------------------------------------------
'Purpose  : Delete a file to the recycle bin
'
'Prereq.  : -
'Parameter: sFileName   - File name incl. full path
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local l As Long
   Local udt As SHFILEOPSTRUCT

   ' File name(s) must be terminated with two NUL
   sFile = sFile & $Nul & $Nul
   udt.pFrom = StrPtr(sFile)

   ' In order to delete to the rec bin, this MUST be true
   udt.fFlags = %FOF_ALLOWUNDO

   If IsFalse(lShowProgress) Then
      udt.fFlags = udt.fFlags Or %FOF_SILENT
   End If

   If IsFalse(lConfirmation) Then
      udt.fFlags = udt.fFlags Or %FOF_NOCONFIRMATION
   End If

   If IsTrue(lSimple) Then
      udt.fFlags = udt.fFlags Or %FOF_SIMPLEPROGRESS
   End If

   If IsFalse(lSysErrors) Then
      udt.fFlags = udt.fFlags Or %FOF_NOERRORUI
   End If

   udt.wFunc = %FO_DELETE
   udt.hWnd = hWndParent

   Delete2RecycleBin = SHFileOperation(udt)

End Function
'==============================================================================

' 'API Konstanten für ShellAndWaitApi
'Private Const NORMAL_PRIORITY_CLASS As Long = &H20
'Private Const INFINITE As Long = -1
'Private Const STARTF_USESHOWWINDOW  As Long = &H1

Function ShellAndWaitApi(ByVal szExec As AsciiZ * %Max_Path) As Long
'------------------------------------------------------------------------------
'Funktion : Startet externes Programm und wartet auf Beendigung
'
'Vorauss. : -
'Parameter: -
'Rückgabe : -
'Notiz    : -
'
'    Autor: MS
'   Quelle: http://support.microsoft.com/kb/129797
' geändert: -
'------------------------------------------------------------------------------
   Local udtProc As PROCESS_INFORMATION
   Local udtStart As STARTUPINFO
   Local lRet As Long

   ' Window styles
   '%SW_HIDE             = 0
   '%SW_SHOWNORMAL       = 1
   '%SW_NORMAL           = 1
   '%SW_SHOWMINIMIZED    = 2
   '%SW_SHOWMAXIMIZED    = 3
   '%SW_MAXIMIZE         = 3
   '%SW_SHOWNOACTIVATE   = 4
   '%SW_SHOW             = 5
   '%SW_MINIMIZE         = 6
   '%SW_SHOWMINNOACTIVE  = 7
   '%SW_SHOWNA           = 8
   '%SW_RESTORE          = 9
   '%SW_SHOWDEFAULT      = 10
   '%SW_FORCEMINIMIZE    = 11
   '%SW_MAX              = 11


   ' Initialize the STARTUPINFO structure:
   udtStart.dwFlags = %STARTF_USESHOWWINDOW
   udtStart.wShowWindow = %SW_ShowMinNoActive
   udtStart.cb = Len(udtStart)

   lRet = CreateProcessA($Nul, szExec, ByVal 0&, ByVal 0&, 1&, _
      %Normal_Priority_Class, 0&, $Nul, udtStart, udtProc)

   ' Wait for the shelled application to finish:
   lRet = WaitForSingleObject(udtProc.hProcess, %INFINITE)
   Call GetExitCodeProcess(udtProc.hProcess, lRet)
   Call CloseHandle(udtProc.hThread)
   Call CloseHandle(udtProc.hProcess)

   ShellAndWaitApi = lRet

End Function
'==============================================================================

Function CreateTimeStamp(Optional ByVal vntDate As Variant, Optional ByVal vntDelim As Variant, _
   Optional ByVal vntDateOnly As Variant, Optional ByVal vntFormat As Variant) As String
'------------------------------------------------------------------------------
'Purpose  : Creates a (date)(time) stamp string of format YYYYMMDDHHNNSS
'
'Prereq.  : -
'Parameter: vntDate     - Create time stamp from this date, defaults to Now()
'           vntDelim    - Date part delimiter, defaults to "-"
'           vntDateOnly - True: Omit the time part of dtmDate
'           vntFormat   - user-defined format, must be a format
'                         supported by Win32 API GetDateFormat, GetTimeFormat
'Returns  : -
'Note     : If both vntDelim *and* vntFormat are passed, vntDelim takes
'           precedence
'
'   Author: Knuth Konrad 25.01.2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sResult As String

   Local dtmDate As IPowerTime
   Local sDelim, sFormat As String
   Local bolDateOnly As Long

   Trace On
   Trace Print FuncName$

   ' Defaults
   If Not IsMissing(vntDate) Then
      Let dtmDate = vntDate
   Else
      Let dtmDate = PBNow()
   End If

   If Not IsMissing(vntDelim) Then
      sDelim = Variant$(vntDelim)
   Else
      sDelim = "-"
   End If

   If Not IsMissing(vntDateOnly) Then
      bolDateOnly = Variant#(vntDateOnly)
   Else
      bolDateOnly = %True
   End If

   If Not IsMissing(vntFormat) Then
      sFormat = Variant$(vntFormat)
   Else
      sFormat = "yyyymmddhhnnss"
   End If

   If sDelim = "-" Then
      ' When a delimitr was passed, it takes
      ' precedence over a possible custom format mask

      sResult = Format$(dtmDate.Year, "0000") & sDelim & _
         Format$(dtmDate.Month, "00") & sDelim & _
         Format$(dtmDate.Day, "00")

      If IsTrue(bolDateOnly) Then
         CreateTimeStamp = sResult
      Else
         CreateTimeStamp = sResult & sDelim & _
         Format$(dtmDate.Hour, "00") & _
         Format$(dtmDate.Minute, "00") & _
         Format$(dtmDate.Second, "00")
      End If

   Else
      ' A custom format was passed

      If IsTrue(bolDateOnly) Then
         CreateTimeStamp =  FormatDate(ByCopy sFormat, dtmDate)
      Else
         CreateTimeStamp =  FormatDate(ByCopy sFormat, dtmDate) & FormatTime(ByCopy sFormat, dtmDate)
      End If

   End If

End Function
'==============================================================================

Function GetWindowsComputerName() As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve the (local) machine's name
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 26.01.2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lpszBuff As AsciiZ * %MAX_COMPUTERNAME_LENGTH
   Local lLen As Long, lRet As Long

   ' Get the Computer Name
   lpszBuff = Space$(%MAX_COMPUTERNAME_LENGTH)
   lLen = Len(lpszBuff)
   lRet = GetComputerName(lpszBuff, lLen)
   If lRet > 0 Then
      Function = Left$(lpszBuff, lLen)
   End If

End Function
'==============================================================================

Function GetWindowsUserName() As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve the logged on (Windows) user's name
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : Returns the user name of the process owning user
'
'   Author: Knuth Konrad 26.01.2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lpszBuff As AsciiZ * 256
   Local lLen As Long, lRet As Long

   'Get the Login User Name
   lpszBuff = Space$(256)
   lLen = Len(lpszBuff)
   lRet = GetUserName(lpszBuff, lLen)

   If lRet > 0 Then
      Function = Left$(lpszBuff, lLen - 1)
   End If

End Function
'==============================================================================

Sub GetFileVersion(ByVal sFile As String, ByRef lVerMajor As Long, ByRef lVerMinor As Long, _
   ByRef lVerRevision As Long, ByRef lVerBuild As Long)
'------------------------------------------------------------------------------
'Purpose  : Retrieve the version number of an executable.
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     :
'
'   Author: Michael Matthias (MCM)
'   Source: https://forum.powerbasic.com/forum/user-to-user-discussions/programming/54705-solved-getfileversioninfo-fails
'  Changed: 26.01.2018
'           - Converted to Sub. Return numerical version information via ByRef
'           parameters
'------------------------------------------------------------------------------
   Local lResSize As Long
   Local ffi As VS_FIXEDFILEINFO Ptr
   Local lRet As Long
   Local sBuffer As String

   ' Defaults
   lVerMajor = 0
   lVerMinor = 0
   lVerRevision = 0
   lVerBuild = 0

   lResSize = GetFileVersionInfoSize (ByCopy sFile, lRet)
   If lResSize= 0 Then
      Exit Sub
   End If

   sBuffer = Space$(lResSize)
   lRet =  GetFileVersionInfo(ByCopy sFile, %NULL, lResSize, ByVal StrPtr(sBuffer))

   ' ** Read the VS_FIXEDFILEINFO info
   VerQueryValue ByVal StrPtr(sBuffer), "\", ffi, SizeOf(@ffi)
   lVerMajor = @ffi.dwProductVersionMs  \ &h10000
   lVerMinor = @ffi.dwProductVersionMs Mod &h10000
   lVerRevision = @ffi.dwProductVersionLS Mod &h10000  ' this is for MY software which uses VERSION_MAJOR, VERSION_MINOR, 0, VERSION_BUILD under FILEVERSION
   lVerBuild = @ffi.dwProductVersionLS \  &h10000

End Sub
'==============================================================================

Function GetFileVersionString(ByVal sFile As String) As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve the version number of an executable as a string of format
'           <Major>.<Minor>.<Revision>.<Build>
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     :
'
'   Author: Knuth Konrad 26.01.2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lMajor, lMinor, lRevision, lBuild As Long

   GetFileVersion sFile, lMajor, lMinor, lRevision, lBuild
   Function = Format$(lMajor) & "." & _
      Format$(lMinor) & "." & _
      Format$(lRevision) & "." & _
      Format$(lBuild)

End Function
'==============================================================================

Function FullPathAndUNC(ByVal sPath As String) As String
'------------------------------------------------------------------------------
'Purpose  : Resolves/expands a path from a relative path to an absolute path
'           and UNC path, if the drive is mapped
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 30.01.2017
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   ' Determine if it's a relative or absolute path, i.e. .\MyFolder or C:\MyFolder
   Local szPathFull As AsciiZ * %Max_Path, sPathFull As String, lResult As Long
   sPathFull = sPath
   lResult = GetFullPathName(ByCopy sPath, %Max_Path, szPathFull, ByVal 0)
   If lResult <> 0 Then
      sPathFull = Left$(szPathFull, lResult)
   End If

   ' Now that we've got that sorted, resolve the UNC path, if any
   Local dwError As Dword
   FullPathAndUNC = UNCPathFromDriveLetter(sPathFull, dwError, 0)

End Function
'------------------------------------------------------------------------------

Function UNCPathFromDriveLetter(ByVal sPath As String, ByRef dwError As Dword, _
   Optional ByVal lDriveOnly As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Returns a fully qualified UNC path location from a (mapped network)
'           drive letter/share
'
'Prereq.  : -
'Parameter: sPath       - Path to resolve
'           dwError     - ByRef(!), Returns the error code from the Win32 API, if any
'           lDriveOnly  - If True, return only the drive letter
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 17.07.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   ' 32-bit declarations:
   Local sTemp As String
   Local szDrive As AsciiZ * 3, szRemoteName As AsciiZ * 1024
   Local lSize, lStatus As Long

   ' The size used for the string buffer. Adjust this if you
   ' need a larger buffer.
   Local lBUFFER_SIZE As Long
   lBUFFER_SIZE = 1024

   If Len(sPath) > 2 Then
      sTemp = Mid$(sPath, 3)
      szDrive = Left$(sPath, 2)
   Else
      szDrive = sPath
   End If

   ' Return the UNC path (\\Server\Share).
   lStatus = WNetGetConnectionA(szDrive, szRemoteName, lBUFFER_SIZE)

   ' Verify that the WNetGetConnection() succeeded. WNetGetConnection()
   ' returns 0 (NO_ERROR) if it successfully retrieves the UNC path.
   If lStatus = %NO_ERROR Then

      If IsTrue(lDriveOnly) Then

         ' Display the UNC path.
         UNCPathFromDriveLetter = Trim$(szRemoteName, Any $Nul & $WhiteSpace)

      Else

         UNCPathFromDriveLetter = Trim$(szRemoteName, Any $Nul & $WhiteSpace) & sTemp

      End If

   Else

      ' Return the original filename/path unaltered
      UNCPathFromDriveLetter = sPath

   End If

   dwError = lStatus

End Function
'------------------------------------------------------------------------------

Function GetIPAddressv4ForHost(Optional ByVal vntHost As Variant) As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve the IPv4 address for a host name.
'
'Prereq.  : -
'Parameter: vntHost     - Host to look up, if empty retieve IP of local machine
'Returns  : Formatted IP address string
'Note     : -
'
'   Author: Knuth Konrad 09.08.2018
'   Source: PB/Win help file, Keyword "HOST ADDR"
'  Changed: -
'------------------------------------------------------------------------------
   Local wsHost As WString, sResult As String
   Local lIP As Long
   Local p As Byte Ptr

   If IsMissing(vntHost) Then
      wsHost = ""
   Else
      wsHost = Variant$$(vntHost)
   End If

   Host Addr wsHost To lIP
   p = VarPtr(lIP)

   GetIPAddressv4ForHost = Using$("#_.#_.#_.#", @p, @p[1], @p[2], @p[3])

End Function
'------------------------------------------------------------------------------

' *** Possibly unused methods --- >


#If 0
Function IsOSNT() As Boolean
'------------------------------------------------------------------------------
'Funktion : Ermittelt ob das OS Windows NT oder Win9x ist
'
'Vorauss. : -
'Parameter: -
'Rückgabe : True  -  OS ist NT
'           False -  OS ist Win95 oder Win98
'Notiz    : -
'
'Autor    : Knuth Konrad
'erstellt : 02.11.1999
'geändert :
'------------------------------------------------------------------------------

'Zur Nutzung des OSVERSIONINFO-Struktur ist zunächst ihrem Parameter dwOSVersionInfoSize
'die Größe der Struktur zu übergeben, was wir wie immer mit der VB-Anweisung Len
'bewerkstelligen. Eine Variable dieses Typs kann dann an GetVersionEx übergeben werden, um
'die interessierenden Informationen zu ermitteln:

Dim OS As OSVERSIONINFO

OS.dwOSVersionInfoSize = Len(OS)
GetVersionEx OS

'Zur Ermittlung des Betriebssystem reicht es aus, die Parameter dwPlatformId, dwMajorVersion
'und dwMinorVersion auszuwerten: Der Parameter dwMajorVersion trägt für Windows NT 4,
'Windows 95 und Windows 98 immer den Wert 4.
'
'Unter Windows NT ist dwPlatformId immer gleich der Konstanten VER_PLATFORM_WIN32_NT,
'während diese unter Windows 95 und Windows 98 den Wert von
'VER_PLATFORM_WIN32_WINDOWS annimmt. In letzterem Fall kann zwischen den beiden durch
'Auswertung des Parameters dwMinorVersion unterschieden werden: Unter Windows 95 ist dieser
'gleich 0, unter Windows 98 hingegen beträgt sein Wert 10. Somit ergibt sich:

With OS
   If .dwMajorVersion = 4 Then
   'Windows NT, Windows 95 oder Windows 98
      If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
      'Windows NT
         IsOSNT = True
         Exit Function
      End If

      If .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
         IsOSNT = False
         Exit Function
      'Windows 98 oder Windows 95
'          If .dwMinorVersion = 0 Then
'             ' das niederwertige word von dwBuildnumber prüfen
'             If (.dwBuildNumber And &HFFFF&) > 1000 Then
'                MsgBox "Windows 95 >= OS R2"
'             Else
'                MsgBox "Windows 95 < OS R2"
'             End If
'          ElseIf .dwMinorVersion = 10 Then
'             MsgBox "Windows 98"
'          End If
      End If
   End If
End With

End Function
#EndIf


Function ShortenPathText(ByVal sPath As String, ByVal lMaxLen As Long) As String
'------------------------------------------------------------------------------
'Funktion : Kürzt eine Pfadangabe auf lMaxLen Zeichen
'
'Vorauss. : -
'Parameter: sPath    -  zu kürzende Pfadangabe
'           lMaxLen  -  maximal Länge des Pfades
'Rückgabe : -
'
'Autor    : Doberenz & Kowalski
'erstellt : 26.11.1999
'geändert : Knuth Konrad
'           ungarische Notation und Stringfunktion statt Variant (Mid, Left...)
'Notiz    : Quelle: Visual Basic 6 Kochbuch, Hanser Verlag
'------------------------------------------------------------------------------
   Local i, lLen, lDiff As Long
   Local sTemp As String

   lLen = Len(sPath)

   ShortenPathText = sPath

   If Len(sPath) < = lMaxLen Then
      Exit Function
   End If

   For i = (lLen - lMaxLen + 6) To lLen
      If Mid$(sPath, i, 1) = "\" Then Exit For
   Next i

   If InStr(sPath, "\") < 1 Then
   ' Ist wohl nur eine Datei, ohne Pfadangabe -> die "Mitte" des Namens kürzen

      sTemp = sPath

      If lLen > lMaxLen Then
         lDiff = lLen - lMaxLen
      Else
         lDiff = 0
      End If

      lDiff = lDiff \ 2

      If lDiff > 2 Then
         sTemp = Left$(sPath, lLen \ 2) & "..." & Right$(sPath, lLen \ 2)
      End If

   Else
      If i < lLen Then
         sTemp = Left$(sPath, 3) & "..." & Right$(sPath, lLen - (i - 1))
      Else
         sTemp = Left$(sPath, 3) & "..." & Right$(sPath, lMaxLen - 6)
      End If
   End If

   ShortenPathText = sTemp

End Function
'==============================================================================
