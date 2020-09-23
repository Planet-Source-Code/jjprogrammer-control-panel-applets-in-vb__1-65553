Attribute VB_Name = "CtrlPanelModule1"
''''''''''''''''''''''''''''''''''''''''''''''''''
''    This project is an adaption of the        ''
''     DLL PROJECT ©2004 DanSoft Australia      ''
'' http://www.Planet-Source-Code.com/vb/scripts ''
'' /ShowCode.asp?txtCodeId=54190&lngWId=1       ''
''                                              ''
''    Your dlls MUST HAVE a DLLMain and Main    ''
''  proc, otherwise it won't compile properly!  ''
''''''''''''''''''''''''''''''''''''''''''''''''''
'Control Panel Messages Constants:
Private Const CPL_INIT = 1
Private Const CPL_GETCOUNT = 2
Private Const CPL_INQUIRE = 3
Private Const CPL_SELECT = 4
Private Const CPL_DBLCLK = 5
Private Const CPL_STOP = 6
Private Const CPL_EXIT = 7
Private Const CPL_NEWINQUIRE = 8
Private Const CPL_DYNAMIC_RES = vbNull
'The structure sent to Control Panel from CPlApplet in response to the CPL_INQUIRE message
'When a dialog box is provided, the stucture contains resource information and an application-defined value for a dialog box.
'To dynamically set icon or display strings, you can specify the CPL_DYNAMIC_RES value for the idIcon,
'idName, or idInfo members rather than specifying a valid resource identifier. This causes the Control Panel to send
'the CPL_NEWINQUIRE message each time it needs the icon and display strings. The Applet is also loaded each time the
'CPL_NEWINQUIRE message is sent, adding some overhead, but this makes it possible without a dialog box.
'Important note: after changing values and recompiling the .cpl, you must restart the control panel without the .cpl present
'or else log off and back on. This will reinitialize the control panel and flush the cache it keeps.

Private Type CPLINFO            'C++ :   typedef struct tagCPLINFO {
    idIcon As Long              '           int idIcon;           - Resource identifier of the icon that represents the dialog box.
    idName As Long              '           int idName;           - Resource identifier of the string containing the dialog box name(displayed below the icon).
    idInfo As Long              '           int idInfo;           - Resource identifier of the string containing the dialog box description (tooltip and statusbar text).
    lData  As Long              '           LONG_PTR lpData;      - Pointer to data defined by the application. When the Control Panel sends the
                                '                                   CPL_DBLCLK and CPL_STOP messages, it passes this value back to your application.
End Type                        '         } CPLINFO;

'The structure sent to Control Panel from CPlApplet in response to the CPL_NEWINQUIRE message
Private Type NEWCPLINFO         'C++ :  typedef struct tagNEWCPLINFO {
    dwSize As Long              '           DWORD dwSize;         - Length of the structure, in bytes.
                                '                                   A DWORD is 32 bytes - same as VB long.
    dwFlags As Long             '           DWORD dwFlags;        - This member is ignored.
    dwHelpContext As Long       '           DWORD dwHelpContext;  - This member is ignored.
    lData As Long               '           LONG_PTR lpData;      - Pointer to data defined by the application. When the Control Panel sends the
                                '                                   CPL_DBLCLK and CPL_STOP messages, it passes this value back to your application.
    hIcon As Long               '           HICON hIcon;          - Identifier of the icon that represents the dialog box.
    szName(1 To 64) As Byte     '           TCHAR szName[32];     - Null-terminated string containing the dialog box name(displayed below the icon).
    szInfo(1 To 128) As Byte    '           CHAR szInfo[64];      - Null-terminated string containing the dialog box description (tooltip and statusbar text).
    szHelpFile As String * 128  '           TCHAR szHelpFile[128];- This member is ignored.
End Type                        '         } CPLNEWINFO;

Private ci  As CPLINFO
Private ci2 As NEWCPLINFO
Private hModule As Long

Function DLLMain(ByVal A As Long, ByVal B As Long, ByVal c As Long) As Long
    hModule = VarPtr(A)
    DLLMain = 1
End Function

Sub Main()
    'This is a dummy, so the IDE doesn't complain
    'there is no Sub Main.
End Sub

'since VB doesn't take VOID parameters we'll use the NEWCPLINFO structure for both CPL_INQUIRE and CPL_NEWINQUIRE messages
'and use a little trickery to set values
Public Function CPlApplet(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam1 As Long, lParam2 As NEWCPLINFO) As Long
  Dim x As Integer, numApplets As Long, sHelp As String, sIcon As Long, sName As String, sInfo As String
 sName = "My Applet"
 sInfo = "Control Panel Applet in VB!"
 Select Case uMsg
    Case CPL_INIT
        CPlApplet = 1
    Case CPL_GETCOUNT
        numApplets = 1
        CPlApplet = numApplets
    Case CPL_INQUIRE
    'in response to CPL_INQUIRE message, Contol Panel expects lParam2
    '  to be actually the CPLINFO structure:              not the CLP_NEWINFO structure:
                    ' idIcon as long                           dwSize As Long
                    ' idName as long                           dwFlags As Long
                    ' idInfo as long                           dwHelpContext As Long
                    ' lData  as long                           lData As Long
                    '                                           ...etc
                                
        lParam2.dwSize = CPL_DYNAMIC_RES 'so this is actually the idIcon variable in CPLINFO (second long value in the structure)
                                'which when set to null takes on the default icon for the .cpl.
                                'To change the icon just change the default icon in project properties--make.
                                '(hint: change Form1.icon)
                                
        lParam2.lData = hModule 'Our hInstance in DLLMain. Since the lData variable is the fourth long in both stuctures,
                                'there's no problem here.
        'idName and idInfo are set to dynamic by default (which is what we want) so we don't need to set them here.
                                
        CPlApplet = 0
    Case CPL_NEWINQUIRE
        lParam2.dwSize = Len(lParam2)
        lParam2.dwFlags = 0
        lParam2.dwHelpContext = 0
        'This is how we get the VB string into a null-terminated string.(VB strings stored internally can contain nulls and won't work correctly)
        For x = 1 To 64
            If x > Len(sName) Or x = 64 Then
                lParam2.szName(x) = 0
            Else
                lParam2.szName(x) = Asc(Mid$(sName, x, 1))
            End If
        Next x
        For x = 1 To 128
            If x > Len(sInfo) Or x = 128 Then
                lParam2.szInfo(x) = 0
            Else
                lParam2.szInfo(x) = Asc(Mid$(sInfo, x, 1))
            End If
        Next x
        CPlApplet = 1
    Case CPL_SELECT
        CPlApplet = 0
    Case CPL_DBLCLK
            Shell "notepad.exe", vbNormalFocus
        CPlApplet = 1
    Case CPL_STOP
        CPlApplet = 0
    Case CPL_EXIT
        CPlApplet = 0
End Select

End Function
