VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   1515
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   6585
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      ScrollBars      =   1  'Horizontal
      TabIndex        =   7
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Call VB .cpl"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Dynamic idInfo"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Dynamic idName"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "idInfo"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "idName"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "No. of applets"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Control Panel Messages Constants:
Private Const CPL_INIT = 1
Private Const CPL_GETCOUNT = 2
Private Const CPL_INQUIRE = 3
Private Const CPL_SELECT = 4 'obsolete
Private Const CPL_DBLCLK = 5
Private Const CPL_STOP = 6
Private Const CPL_EXIT = 7
Private Const CPL_NEWINQUIRE = 8
Private Type CPLINFO
    idIcon          As Long       'icon resource id in dialog
    idName          As Long       'Name string resource in dialog
    idInfo          As Long       'Info string resource in dialog
    lData           As Long       'dialog handle
End Type
Private Type NEWCPLINFO
    dwSize As Long                'size of NEWCLPLINFO structure
    dwFlags As Long
    dwHelpContext As Long         'not used
    lData As Long                 'long pointer to app instance
    hIcon As Long                 'long pointer to app icon
    szInfo As String * 224        'info:
                                      'szName as String * 32
                                      'szInfo As String * 64
                                      'szHelpFile As String * 128
    
End Type

Private ci      As CPLINFO
Private ci2      As NEWCPLINFO

Private Declare Function CPlApplet_MMSys Lib "ControlPanelVb.cpl" Alias "CPlApplet" (ByVal hwndCPl As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long



Private Sub Command1_Click()
Dim x As Long, tmpstring As String, str2 As Integer, textstring As String
If CPlApplet_MMSys(hWnd, CPL_INIT, 0, 0) <> 0 Then
        x = CPlApplet_MMSys(hWnd, CPL_GETCOUNT, 0, VarPtr(ci))
        Text3 = x
        CPlApplet_MMSys hWnd, CPL_INQUIRE, 0, VarPtr(ci)
        If ci.idName > 0 Then 'string resource id from fixed dialog
            Text1 = "Fixed dialog resource (Name) #" & ci.idName
        Else
            Text1 = "none - dynamic"
        End If
        If ci.idInfo > 0 Then 'string resource id from fixed dialog
            Text2 = "Fixed dialog resource (Info) #" & ci.idInfo
        Else
            Text2 = "none - dynamic"
        End If
        CPlApplet_MMSys hWnd, CPL_NEWINQUIRE, 0, VarPtr(ci2)
        tmpstring = StrConv(ci2.szInfo, vbUnicode)
        If Left$(tmpstring, 2) > Chr$(0) & Chr$(0) Then 'dynamic string
            str2 = InStr(1, tmpstring, Chr$(0) & Chr$(0))
            Text4 = strip(Left$(tmpstring, str2), Chr$(0))
            Text5 = strip(Mid$(tmpstring, str2, Len(tmpstring)), Chr$(0))
        End If
        CPlApplet_MMSys hWnd, CPL_DBLCLK, 0, ci.lData
        CPlApplet_MMSys hWnd, CPL_STOP, 0, ci.lData
        CPlApplet_MMSys hWnd, CPL_EXIT, 0, 0
    End If
End Sub
'Strips embedded character$ from teststring$
Private Function strip$(ByVal teststring$, ByVal charstring$, Optional MatchCase As Boolean = False, Optional StripEachChar As Boolean = False)
Dim temp$, CharStringLen As Integer, i As Integer, x As Integer, CompareMode As Integer
temp$ = teststring$
If Not MatchCase Then CompareMode = 1 Else CompareMode = 0

If StripEachChar Then CharStringLen = Len(charstring$) Else CharStringLen = 1
For x = 1 To CharStringLen
start:
If StripEachChar Then
    i% = InStr(1, temp$, Mid$(charstring$, x, 1), CompareMode)
Else
    i% = InStr(1, temp$, charstring$, CompareMode)
End If
If i% Then
    temp$ = Left$(temp$, i% - 1) & Mid$(temp$, i% + 1, Len(temp$))
    GoTo start
End If

Next x
strip$ = temp$
End Function
                                                                                                                                                                                                                                                               
                                                                                                                                                                                                                                                               
                                                                                                                                                                                                                                                               
                                                                                                                                                                    
