'------------------------------------------------------------------------------
'Purpose  : Renames any file as yyyymmdd.<Ext>
'
'Prereq.  : -
'Note     : /f:<file> /d:<directory> /o /p:<Prefix>
'
'   Author: Knuth Konrad 15.07.2002
'   Source: -
'  Changed: 19.04.2017
'           - Code reformatting
'           15.05.2017
'           - Application manifest added
'           - Replace source code include with SLL
'------------------------------------------------------------------------------

'----------------------------------------------------------------------------
'*** PROGRAMM/COMPILE OPTIONEN ***
#Compile Exe ".\RENToday.exe"
#Break On
#Option Version5
#Dim All

#Link "baCmdLine.sll"

#Debug Error On
#Tools Off

DefLng A-Z

%VERSION_MAJOR = 1
%VERSION_MINOR = 5
%VERSION_REVISION = 2

' Version Resource information
#Include ".\RENTodayRes.inc"
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
#Include Once "Win32API.inc"
'#Include "IbaCmdLine.inc"
#Include "sautilcc.inc"
'------------------------------------------------------------------------------
'*** Variabels ***
'------------------------------------------------------------------------------
' Flags retrieved from command line
Global glOverwrite As Long
Global gsPrefix As String
'==============================================================================

Function PBMain()
'------------------------------------------------------------------------------
'Purpose  : Programm startup method
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
   Local sTemp, sCmd As String
   Local sNew, sOrg As String
   Local i, j As Long
   Local lFound, lFileAction, lDirAction As Long

   Trace New ".\RenToday.tra"
   Trace On

   ' Application intro
   Print ""
   ConHeadline "RENToday", %VERSION_MAJOR, %VERSION_MINOR, %VERSION_REVISION
   ConCopyright "2002-2017", $COMPANY_NAME
   Print ""

   If Len(Trim$(Command$)) < 1 Or InStr(Command$, "/?") > 0  Then
      ShowHelp
      Exit Function
   End If

   ' *** Parse the parameters passed
   ' Valid CLI parameters are:
   ' /f= (file)
   ' /d= (directory)
   ' /o= (overwrite)
   ' /p= (prefix)
   sCmd = Command$

   Local o As IBACmdLine
   Local vnt As Variant

   Let o = Class "cBACmdLine"

   If IsFalse(o.Init(sCmd)) Then
      Print "Couldn't parse parameters: " & sCmd
      Print "Type RENToday /? for help"
      Let o = Nothing
      Exit Function
   End If

   i = o.ValuesCount
   If (i < 1) Or (i > 3) Then
      Print "Invalid number of parameters."
      Print ""
      ShowHelp
      Exit Function
   End If

   ' Parse CLI parameters

   ' Single file only
   If IsTrue(o.HasParam("f")) Then
      sTemp = Variant$(o.GetValueByName("f"))
      lFound = %TRUE : lFileAction = %TRUE
      sNew = sTemp
      sOrg = sNew

   ' All file in a directory
   ElseIf IsTrue(o.HasParam("d")) Then
      sTemp = Variant$(o.GetValueByName("d"))
      lFound = %TRUE : lDirAction = %TRUE
      sNew = sTemp
      sOrg = sNew

   End If

   ' Overwrite existing files?
   If IsTrue(o.HasParam("o")) Then
      glOverwrite = %TRUE
   End If

   ' Prefix?
   If IsTrue(o.HasParam("p")) Then
      sTemp = Variant$(o.GetValueByName("p"))
      gsPrefix = Remove$(sTemp, $Dq)
   End If

   ' *** Sanity check(s)
   If IsFalse(lFileAction) And IsFalse(lDirAction) Then
      Print "No file or directory specified."
      Print ""
      ShowHelp
      Exit Function
   End If


   ' Echo the CLI parameters
   If IsTrue(lFileAction) Then
      Con.StdOut "Source file     : " & sNew
   ElseIf IsTrue(lDirAction) Then
      Con.StdOut "Source directory: " & sNew
   End If
   Con.StdOut "Overwrite       : " & Choose$(glOverwrite, "True" Else "False")
   Con.StdOut "Prefix          : " & Switch$(Len(gsPrefix) < 1, "<none>", Len(gsPrefix) >= 1, gsPrefix)
   Con.StdOut ""

   ' *** We're in business. Rename a single file or files in a folder?

   Local dwFileCount As Dword
   If IsTrue(lFileAction) Then
   ' A file

      ' Safeguard
      sNew = Unwrap$(sNew, $Dq, $Dq)
      If IsFalse(FileExist(ByCopy sNew)) Then
         Con.StdOut "File " & sNew & " not found."
         Exit Function
      End If

      Call RenFile(sNew, sOrg)
      dwFileCount = 1

   ElseIf IsTrue(lDirAction) Then
   ' A folder

      dwFileCount = RenDirectory(sNew)

   End If

   Con.StdOut ""
   Con.StdOut "File(s) renamed: " & Format$(dwFileCount)

   PBMain = dwFileCount

End Function
'------------------------------------------------------------------------------

Sub ShowHelp()
'------------------------------------------------------------------------------
'Purpose  : Show help screen
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

   Print ""
   Print "RENToday Usage:"
   Print "---------------"
   Print "RENToday /f=<Filename>|/d=<directory with file spec> [/p=<Prefix>] [/o]"
   Print ""
   Print "i.e."
   Print ""
   Print "RENToday /f=d:\data\myfile.txt"
   Print "  will rename the single file d:\data\myfile.txt to 20020228_134228_623.txt, assuming today's date is"
   Print "  February 28th, 2002 and the time is 13:42:28 (and 623 milliseconds)."
   Print "- or -"
   Print ""
   Print "RENToday /d=d:\data\*.txt"
   Print "  will rename each file with the file extension .txt in d:\data\ to 20020228_134228_623.txt, assuming today's date is"
   Print "  February 28th, 2002 and the time is 13:42:28 (and 623 milliseconds) when the first renaming occurs."
   Print "  Of course, the time will be updated for each following renaming process. RENToday implements a 3 millisecond delay"
   Print "  internally to ensure that file names are unique."
   Print ""
   Print "Please note: switch /f takes precedence over switch /d."
   Print ""
   Print "RENToday /f=d:\data\myfile.txt /p=MyPrefix_"
   Print "  will rename d:\data\myfile.txt to MyPrefix_20020228.txt, assuming the above example date & time."
   Print ""
   Print "Option /p=<prefix> adds <prefix> in front of the file's name."
   Print "Option /o will overwrite an existing file."

End Sub
'---------------------------------------------------------------------------

Function RenFile(ByVal sNew As String, sOrg As String) As Long
'------------------------------------------------------------------------------
'Purpose  : Renames a single file.
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad
'   Source: -
'  Changed: 09.02.2017
'           - File extension was wrongly echoed to command line
'------------------------------------------------------------------------------
   Local sOld, sExt As String
   Local sMsg, sTemp As String
   Local oPTime As IPowerTime

   ErrClear

   Let oPTime = Class "PowerTime"
   sTemp = Format$(oPTime.Hour, "00") & "_" & Format$(oPTime.Minute, "00")  & "_" & Format$(oPTime.Second, "00") & "_" & Format$(oPTime.MSecond, "000")

   sExt = Remove$(Mid$(sNew, GetExtPos(sNew)), $Dq)

   oPTime.Now
   sTemp = gsPrefix & Right$(Date$,4) & Left$(Date$,2) & Mid$(Date$,4,2) & _
      "_" & Format$(oPTime.Hour, "00") & Format$(oPTime.Minute, "00") & Format$(oPTime.Second, "00") & "_" & Format$(oPTime.MSecond, "000")

   sOld = ExtractFileName(sNew)

   Replace sOld With sTemp In sNew

   sMsg = "- Scanning for file "
   Con.StdOut sMsg & ShortenPathText(sOrg, Con.Screen.Col-(1 + Len(sMsg)))

   sTemp = Trim$(Remove$(sOrg, $Dq))
   sNew = sNew & sExt

   If IsTrue(FileExist(ByCopy sNew)) Then
      If IsTrue(glOverwrite) Then
         Kill sNew
         Con.StdOut "  Renaming " & sOrg
         Con.StdOut "   -> " & sNew
         Name sTemp As sNew
      Else
         Con.Color 4, -1
         Con.StdOut "  Skipping ... " & sNew & " already exists."
         Con.Color 7, -1
      End If
   Else
      Con.StdOut "  Renaming " & sOrg
      Con.StdOut "   -> " & sNew
      Name sTemp As sNew
   End If

End Function
'---------------------------------------------------------------------------

Function RenDirectory(ByVal sNew As String) As Long
'------------------------------------------------------------------------------
'Purpose  : Rename all files in a directory
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 15.07.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Local sTemp, sMsg, sPath As String
Local lCount As Long
Local udtDirData As DirData
Local szSourceFile As WStringZ * %Max_Path

Local hSearch As Dword                ' Search handle
Local udtWFD As WIN32_FIND_DATAW      ' FindFirstFile structure

sPath = ExtractPath(sNew)
' Add closing "
sPath = Unwrap$(sPath, $Dq, $Dq)

sMsg = "- Scanning folder "
Con.StdOut sMsg & ShortenPathText(sPath, Con.Screen.Col-(1+Len(sMsg)))

'sSourceFile = NormalizePath(sPath) & sFilePattern
'con.stdout "FileSpec: " & sNew & ", " & format$(len(sNew))
szSourceFile = Unwrap$(Trim$(sNew), $Dq, $Dq)

hSearch = FindFirstFileW(szSourceFile, udtWFD)       ' Get search handle, if success
If hSearch <> %INVALID_HANDLE_VALUE Then    ' Loop through directory for files

   Do

      If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY Then ' If not directory bit is set (files only here...)

         sTemp = Remove$(udtWFD.cFileName, Any Chr$(0))

         sMsg = "  Renaming "
         Con.StdOut  sMsg & ShortenPathText(NormalizePath(sPath) & sTemp, Con.Screen.Col-(1+Len(sMsg)))
         Con.StdOut "   -> " & NewFileName(NormalizePath(sPath) & sTemp)
         Incr lCount

         Try
            If IsTrue(FileExist(ByCopy NewFileName(NormalizePath(sPath) & sTemp))) Then
               If IsTrue(glOverwrite) Then
                  Kill NewFileName(NormalizePath(sPath) & sTemp)
                  Name NormalizePath(sPath) & sTemp As NewFileName(NormalizePath(sPath) & sTemp)
               Else
                  Con.Color 4, -1
                  Con.StdOut "   Skipping ... " & NewFileName(NormalizePath(sPath) & sTemp) & " already exists."
                  Con.Color 7, -1
               End If
            Else
               Name NormalizePath(sPath) & sTemp As NewFileName(NormalizePath(sPath) & sTemp)
            End If

         Catch

            Con.Color 12, -1
            sMsg = "  ERROR: can't rename "
            Con.StdOut  sMsg & ShortenPathText(sTemp, Con.Screen.Col-(1+Len(sMsg)))
            Con.Color 7, -1
            Decr lCount
         End Try

      End If   '// If (udtWFD.dwFileAttributes And %FILE_ATTRIBUTE_DIRECTORY) <> %FILE_ATTRIBUTE_DIRECTORY

      Sleep 3  ' Ensure unique file names, as we're using millisends in the file name

   Loop While FindNextFileW(hSearch, udtWFD)

   FindClose hSearch

End If   '// If hSearch <> %INVALID_HANDLE_VALUE

RenDirectory = lCount

End Function
'---------------------------------------------------------------------------

Function NewFileName(ByVal sFile As String) As String
'------------------------------------------------------------------------------
'Purpose  : Programm startup method
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 15.07.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sOld, sOrg, sExt, sTemp As String
   Local oPTime As IPowerTime

   Let oPTime = Class "PowerTime"
   sTemp = Format$(oPTime.Hour, "00") & "_" & Format$(oPTime.Minute, "00")  & "_" & Format$(oPTime.Second, "00") & "_" & Format$(oPTime.MSecond, "000")

   sExt = Mid$(sFile, GetExtPos(sFile))

   oPTime.Now
   sTemp = gsPrefix & Right$(Date$,4) & Left$(Date$,2) & Mid$(Date$,4,2) & _
      "_" & Format$(oPTime.Hour, "00") & Format$(oPTime.Minute, "00") & Format$(oPTime.Second, "00") & "_" & Format$(oPTime.MSecond, "000")
   sOld = ExtractFileName(sFile)

   Replace sOld With sTemp In sFile

   NewFileName = sFile & sExt

End Function
'---------------------------------------------------------------------------

Function ErrString(ByVal lErr As Long, Optional ByVal vntPrefix As Variant) As String
'------------------------------------------------------------------------------
'Purpose  : Returns an formatted error string from an (PB) error number
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 12.02.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local sPrefix As String

   If Not IsMissing(vntPrefix) Then
      sPrefix = Variant$(vntPrefix)
   End If

   ErrString = sPrefix & Format$(lErr) & " - " & Error$(lErr)

End Function
'------------------------------------------------------------------------------
