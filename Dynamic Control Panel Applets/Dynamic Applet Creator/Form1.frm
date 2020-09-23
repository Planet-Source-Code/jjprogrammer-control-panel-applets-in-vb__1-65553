VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Dynamic Control Panel Applet Creator/Editor"
   ClientHeight    =   2265
   ClientLeft      =   1980
   ClientTop       =   2220
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   10680
   Begin VB.CommandButton Command2 
      Caption         =   "Add  Applet"
      Height          =   495
      Left            =   1250
      TabIndex        =   4
      Top             =   135
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtEdit 
      Height          =   500
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin ControlPanelApplets.ComboFileDir combo1 
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      FontCharset     =   0
      ForeColor       =   -2147483630
   End
   Begin MSFlexGridLib.MSFlexGrid lstApplets 
      Height          =   1525
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   2699
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   500
      BackColor       =   16777088
      ForeColor       =   16711680
      BackColorFixed  =   16761024
      BackColorSel    =   16744576
      BackColorBkg    =   16777152
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Menu MnuMainMenu 
      Caption         =   "MainMenu"
      Visible         =   0   'False
      Begin VB.Menu MnuRemove 
         Caption         =   "Remove Applet"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type AppletInfo
  mAppletName As String
  mAppletIcon As String
  mAppletInfotip As String
End Type

Dim MyApplet As AppletInfo
Dim bDoNotEdit   As Boolean
Dim bOnFixedPart As Boolean
Dim inifile As String, NumberOfApplets As Integer



Private Sub combo1_LostFocus()
Call combo1_Validate(False)
End Sub

Private Sub combo1_Validate(Cancel As Boolean)
pSetCellValue
End Sub



Private Sub Command1_Click()
Dim x As Integer
WriteIni inifile, "General", "NumberOfApplets", Str(NumberOfApplets)
For x = 0 To NumberOfApplets - 1
    If Not lstApplets.TextMatrix(x + 1, 0) > "" Then Exit Sub
    WriteIni inifile, "Applet" & x + 1, "AppletName", lstApplets.TextMatrix(x + 1, 0)
    WriteIni inifile, "Applet" & x + 1, "Info", lstApplets.TextMatrix(x + 1, 1)
    WriteIni inifile, "Applet" & x + 1, "IconFile", lstApplets.TextMatrix(x + 1, 3)
    WriteIni inifile, "Applet" & x + 1, "ExeFile", lstApplets.TextMatrix(x + 1, 4)
Next x
End Sub

Private Sub Command2_Click()
If NumberOfApplets > 0 Then lstApplets.Rows = lstApplets.Rows + 1
    NumberOfApplets = lstApplets.Rows - 1
    lstApplets.Row = NumberOfApplets
    lstApplets.Col = 0
    lstApplets.Text = "Applet " & NumberOfApplets
    lstApplets.Col = 1
    lstApplets.Text = ReadIni(inifile, "Applets", "Info" & NumberOfApplets, "No name available")
    lstApplets.Col = 2
    On Error Resume Next
    Set lstApplets.CellPicture = LoadPicture(ReadIni(inifile, "Applets", "Icon" & NumberOfApplets))  ' Set the cellpicture
    lstApplets.Col = 3
    lstApplets.Text = ReadIni(inifile, "Applets", "Icon" & NumberOfApplets, "No name available")
    lstApplets.Col = 4
    lstApplets.Text = ReadIni(inifile, "Applets", "FileEnumberofappletse" & NumberOfApplets, "No name available")
End Sub

Private Sub Form_Load()
Dim messagestring As String, systemdirectory As String, x As Integer, Y As Integer
lstApplets.Clear
lstApplets.FormatString() = "Applet Name                " & vbTab & "Infotip                                         " & vbTab & " Icon " & vbTab & "Icon  Path                                            " & vbTab & "Application  Path                                       "
txtEdit = ""
bDoNotEdit = False
systemdirectory = Environ("windir") & IIf(Len(Environ("OS")), "\SYSTEM32", "\SYSTEM") & "\"
inifile = systemdirectory & "ControlPanelApplet.ini"
messagestring = ReadIni(inifile, "General", "NumberOfApplets", "No applets available")
NumberOfApplets = Val(messagestring)
If NumberOfApplets = 0 Then lstApplets.AddItem (" ")
For x = 0 To NumberOfApplets - 1
    lstApplets.AddItem ReadIni(inifile, "Applet" & x + 1, "AppletName", "No name available")
    lstApplets.Row = x + 1
    lstApplets.Col = 1
    lstApplets.Text = ReadIni(inifile, "Applet" & x + 1, "Info", "No name available")
    lstApplets.Col = 2
    On Error Resume Next
    Set lstApplets.CellPicture = LoadPicture(ReadIni(inifile, "Applet" & x + 1, "IconFile"))  ' Set the cellpicture
    lstApplets.Col = 3
    lstApplets.Text = ReadIni(inifile, "Applet" & x + 1, "IconFile", "No name available")
    lstApplets.Col = 4
    lstApplets.Text = ReadIni(inifile, "Applet" & x + 1, "ExeFile", "No name available")
Next x
combo1.IncludeFiles = True
End Sub

Private Sub lstApplets_GotFocus()
If bDoNotEdit Then Exit Sub
' Copy the textbox's value to the grid
' and hide the textbox.
'
Call pSetCellValue

End Sub

Private Sub lstApplets_KeyPress(KeyAscii As Integer)
If lstApplets.MouseCol <> 2 Then
    Call EditGrid(lstApplets, txtEdit, KeyAscii)
Else
    Call EditGrid(lstApplets, combo1, KeyAscii)
End If
End Sub


Private Sub lstApplets_LeaveCell()
'bDoNotEdit = True
End Sub


Private Sub lstApplets_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim l      As Long
Dim lWidth As Long
If Button = 2 Then
        PopupMenu MnuMainMenu
        Exit Sub
End If

With lstApplets
    For l = 0 To .Cols - 1
        If .ColIsVisible(l) Then
            lWidth = lWidth + .ColWidth(l)
        End If
    Next
    '
    ' See if we are on the fixed part of the grid.
    '
    bOnFixedPart = (x < 0) Or _
                   (x > lWidth) Or _
                   (Y < .RowHeight(0)) Or _
                   (Y > .Rows * .RowHeight(0))
End With
End Sub


Private Sub lstApplets_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim lXPoint As Long
Dim lYPoint As Long
Dim lIndex As Long

        With lstApplets
            If ActiveControl.Name = "lstApplets" Or ActiveControl.Name = "txtEdit" Or ActiveControl.Name = "combo1" Then
            Else
            .SetFocus
            End If
            ' show tip
            On Error Resume Next 'in case there is no info in .TextMatrix, which causes an error
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        End With

End Sub

Private Sub lstApplets_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Display the textbox if the user clicked
    ' on a non-fixed row or column.
    '
If bOnFixedPart Then
    txtEdit.Visible = False
    combo1.Visible = False
    Exit Sub
End If
If lstApplets.MouseCol < 2 Then
    Call EditGrid(lstApplets, txtEdit, 32)
ElseIf lstApplets.MouseCol > 2 Then
    Call EditGrid(lstApplets, combo1, 32)
End If




End Sub

Private Sub lstApplets_RowColChange()
'bDoNotEdit = True
Select Case lstApplets.Col
Case 3
    combo1.FileSpec = "ICO"
Case 4
    combo1.FileSpec = "EXE"
End Select
End Sub

Private Sub MnuRemove_Click()
Dim x As Long, Y As Integer
txtEdit.Visible = False
combo1.Visible = False
x = MsgBox("This will delete " & lstApplets.TextMatrix(lstApplets.MouseRow, 0) & ", continue?", vbYesNo, "Delete Applet")
If x = vbNo Then Exit Sub
DeleteSection inifile, "Applet" & lstApplets.MouseRow
If Not lstApplets.Rows >= 1 Then
   For Y = 0 To 4
    If Y = 2 Then
     Set lstApplets.CellPicture = LoadPicture()
    Else
     lstApplets.TextMatrix(lstApplets.Row, Y) = " "
    End If
   Next Y
Else
    lstApplets.RemoveItem lstApplets.MouseRow
End If
NumberOfApplets = NumberOfApplets - 1
WriteIni inifile, "General", "NumberOfApplets", Str(NumberOfApplets)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'
' See what key was pressed in the textbox.
 Dim X2 As Long, lngUpper As Long, lngLower As Long
With lstApplets
    Select Case KeyCode
        Case 13   'ENTER
               If Not Trim(Me.lstApplets.TextMatrix(Me.lstApplets.Row, 0)) > "" Then Exit Sub
                    If Me.lstApplets.MouseRow > 0 And Me.lstApplets.RowSel = Me.lstApplets.Row Then
                            If .Col = 0 Then
                                .SetFocus
                                Exit Sub
                            End If
                        With Me.lstApplets
                             .TextMatrix(.Row, .Col) = txtEdit

                            MyApplet.mAppletIcon = .TextMatrix(.Row, 1)
                            MyApplet.mAppletInfotip = .TextMatrix(.Row, 2)
                            MyApplet.mAppletName = .TextMatrix(.Row, 0)
                        End With
                    Else  'selected multiple rows
                        With Me.lstApplets
                            If (.Rows = 1) Or (.Row = 0) Then Exit Sub
                            If .Row < .RowSel Then
                                lngLower = .Row
                                lngUpper = .RowSel
                            Else
                                lngLower = .RowSel
                                lngUpper = .Row
                            End If
                            Select Case .Col
                                Case 0 'file name -- not used for multi edit

                                Case 1, 2, 4 'artist,album,genre
                                    For X2 = lngLower To lngUpper
                                            If lngUpper < lngLower Then Exit For
                                             .TextMatrix(X2, .Col) = txtEdit
                                             MyApplet.mAppletIcon = .TextMatrix(.Row, 1)
                                             MyApplet.mAppletInfotip = .TextMatrix(.Row, 2)
                                             MyApplet.mAppletName = .TextMatrix(.Row, 0)
                                            lngUpper = lngUpper - 1
                                     Next X2

                                Case 3 'title -- not used for multi edit
                            End Select
                        End With
                    End If
            txtEdit.Visible = False
            .SetFocus
        Case 27   'ESC
             txtEdit.Visible = False
            .SetFocus
        Case 37   'Left arrow
            .SetFocus
            DoEvents
            If .Col > .FixedCols Then
                bDoNotEdit = True
                .Col = .Col - 1
                bDoNotEdit = False
            End If
        Case 38   'Up arrow
            .SetFocus
            DoEvents
            If .Row > .FixedRows Then
                bDoNotEdit = True
                .Row = .Row - 1
                bDoNotEdit = False
            End If
        Case 39   'Right arrow
            .SetFocus
            DoEvents
            If .Col < .FixedCols Then
                bDoNotEdit = True
                .Col = .Col + 1
                bDoNotEdit = False
            End If
        Case 40   'Down arrow
            .SetFocus
            DoEvents
            If .Row < .Rows - 1 Then
                bDoNotEdit = True
                .Row = .Row + 1
                bDoNotEdit = False
            End If
    End Select
End With
End Sub
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
' Delete carriage returns and Esc to get rid of beep
Select Case KeyAscii
    Case Asc(vbCr), 27
        KeyAscii = 0
End Select
End Sub
Private Sub pSetCellValue()
'
' NOTE:
'       This code should be called anytime
'       the grid loses focus and the grid's
'       contents may change.  Otherwise, the
'       cell's new value may be lost and the
'       textbox may not line up correctly.
'
If bOnFixedPart Or lstApplets.MouseCol = 2 Then Exit Sub
If bDoNotEdit Then Exit Sub
If txtEdit.Visible Then
   If txtEdit > "" Then lstApplets.Text = txtEdit.Text
    txtEdit.Visible = False
End If
If combo1.Visible Then
        If UCase(Right(combo1.StartingFolder, 3)) = combo1.FileSpec Then
            lstApplets.Text = combo1.StartingFolder
            lstApplets.Col = 2
            If UCase(Right(combo1.StartingFolder, 3)) = "ICO" Then
               Set lstApplets.CellPicture = LoadPicture(combo1.StartingFolder)  ' reset the cellpicture
            End If
        End If
    combo1.Visible = False
End If

End Sub
Private Sub EditGrid(EditList As MSFlexGrid, EditControl As Control, KeyAscii As Integer)

    With EditControl
        Select Case KeyAscii
            Case 0 To 32
                '
                ' Edit the current text.
                If TypeOf EditControl Is TextBox Then
                    .Text = lstApplets
                    .SelStart = 0
                    .SelLength = 1000
                End If
    '        Case 8, 46, 48 To 57
    '            '
    '            ' Replace the current text but only
    '            ' if the user entered a number.
    '            '
    '            .Text = Chr(KeyAscii)
    '            .SelStart = 1
    '        Case Else
    '            '
    '            ' If an alpha character was entered,
    '            ' use a zero instead.
    '            '
    '            .Text = "0"
        End Select
    End With
    
    ' Show the EditControl at the right place.
    '
    With EditList
        If .CellWidth < 0 Then Exit Sub
        If TypeOf EditControl Is TextBox Then
            EditControl.Move .Left + .CellLeft, .top + .CellTop, .CellWidth, .CellHeight
        ElseIf TypeOf EditControl Is ComboFileDir Then
            EditControl.Move .Left + .CellLeft, .top + .CellTop, .CellWidth
            EditControl.StartingFolder = .Text
        End If
        '
        ' NOTE:
        '   Depending on the style of the Grid Lines that you set, you
        '   may need to adjust the textbox position slightly. For example
        '   if you use raised grid lines use the following:
        '
        'editcontrol.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
    End With
    
    EditControl.Visible = True
    EditControl.SetFocus
End Sub

Private Sub txtEdit_LostFocus()
txtEdit.Visible = False
End Sub
