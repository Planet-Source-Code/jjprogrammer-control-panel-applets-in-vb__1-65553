Attribute VB_Name = "INIReadWrite"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function ReadIni(ByVal mFile As String, ByVal iSection As String, ByVal iKeyName As String, Optional iDefault As String) As String
    
    Dim Ret As String, NC As Long
    
    'Create the buffer
    Ret = String(500, Chr$(0))
    
    'Retrieve the string
    NC = GetPrivateProfileString(iSection, iKeyName, iDefault, Ret, Len(Ret), mFile)
    
    'NC is the number of characters copied to the buffer
    If NC <> 0 Then
        Ret = Left$(Ret, NC)
    Else
        'Make sure to cut it down to number of char's returned
        Ret = ""
    End If
    
    'Turn the funky vbcrlf string into VBCRLFs
    Ret = Replace(Ret, "%%&&Chr(13)&&%%", vbCrLf)
    
    'Return the setting
    ReadIni = Ret
End Function

Public Sub WriteIni(ByVal mFile As String, ByVal iSection As String, ByVal iKeyName As String, iValue As String)
    If iValue > vbNullString Then
        'Make sure to change it to a String
        iValue = CStr(iValue)
        'Turn all vbcrlf's into that funky string
        iValue = Replace(iValue, vbCrLf, "%%&&Chr(13)&&%%")
    End If
    WritePrivateProfileString iSection, iKeyName, iValue, mFile
End Sub

Public Function Read_Sections(ByVal mFile As String) As String
    
    Dim Ret As String, NC As Long
    
    'Create the buffer
    Ret = String(500, 0)
    
    'Retrieve the string, return '[-na-]' if there is none
    NC = GetPrivateProfileString(vbNullString, vbNullString, vbNullString, Ret, Len(Ret), mFile)
    
    'NC is the number of characters returned
    If NC <> 0 Then
        Ret = Left$(Ret, NC - 1)
    End If
    
    'Return the sections
    Read_Sections = Ret
End Function

Public Function ReadKeys(ByVal mFile As String, iSection As String) As String
    Dim Ret As String, NC As Long
    
    'Create the buffer
    Ret = String(500, 0)
    
    'Retrieve the string, return '[-na-]' if there is none
    NC = GetPrivateProfileString(iSection, vbNullString, vbNullString, Ret, Len(Ret), mFile)
    
    'NC is the number of characters copied to the buffer
    If NC <> 0 Then
        Ret = Left$(Ret, NC - 1)
    End If
    'Return the sections
    ReadKeys = Ret
End Function

Public Function DeleteSection(ByVal mFile As String, iSection As String)
    WriteIni mFile, iSection, vbNullString, vbNullString
End Function

Function DeleteKey(ByVal mFile As String, iSection As String, iKeyName As String)
    If iKeyName > "" Then WriteIni mFile, iSection, iKeyName, vbNullString
End Function

 Function FileExists(ByVal FileName As String) As Boolean
Dim fso
Set fso = CreateObject("scripting.filesystemobject")
FileExists = fso.FileExists(FileName)
End Function
Sub EditKey(ByVal mFile As String, iSection As String, iKeyName As String, NewKeyName As String)
Dim keyvalue As String
   keyvalue = ReadIni(mFile, iSection, iKeyName)
   DeleteKey mFile, iSection, iKeyName
   WriteIni mFile, iSection, NewKeyName, keyvalue
End Sub
Sub EditSection(SectionName As String, NewSection As String)
Dim fno As Integer, fn1 As Integer, fn2 As Integer
Dim fname As String, apppath$, inifile$, searchtext$
apppath$ = App.Path
If Right$(apppath$, 1) <> "\" Then apppath$ = apppath$ & "\"
inifile$ = apppath$ & "RValues.ini"
fn1 = FreeFile
fname = inifile$

'load the file into string
   Open fname For Input As #fn1
       searchtext$ = Input$(LOF(fn1) - 1, fn1)
   Close #fn1

'edit file
SearchAndReplace searchtext$, "[" & SectionName & "]", "[" & NewSection & "]"
  

'write new file from string
 fn2 = FreeFile
 Open fname For Output As #fn2
       Print #fn2, (searchtext$)
   Close #fn2
  
End Sub
Sub SearchAndReplace(OriginalString As String, SearchVal As String, ReplaceVal As String)
Dim iLoc As Long
Dim SearchLen As Integer

iLoc = InStr(1, OriginalString, SearchVal, vbBinaryCompare)
SearchLen = Len(SearchVal)
OriginalString = Left$(OriginalString, iLoc - 1) + ReplaceVal + Right$(OriginalString, Len(OriginalString) - (iLoc - 1 + SearchLen))
End Sub

