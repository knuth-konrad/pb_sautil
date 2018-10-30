#Compile Exe
#Dim All

#Link "saUtil.sll"


Function PBMain () As Long

   ' Function CreateTempFileName(szPath As AsciiZ, szPrefix As AsciiZ, szFileExtension As AsciiZ) As String
   Con.StdOut CreateTempFileName(".\DATA", "sa", "")

   Local i, lCount As Long
   lCount = ArgC()
   Con.StdOut "Command$, ArgC: " & Command$ & ", " & Format$(lCount)

   For i = 1 To lCount
      Con.StdOut Format$(i, "00) ") & ArgV(i)
   Next i


End Function
