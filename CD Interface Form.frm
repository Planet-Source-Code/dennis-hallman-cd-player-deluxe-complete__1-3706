VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CDIface 
   BorderStyle     =   0  'None
   Caption         =   "CD Deluxe"
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CD Interface Form.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdgIface 
      Left            =   4845
      Top             =   1050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Skin"
      Filter          =   "*.bmp"
   End
   Begin VB.PictureBox picControlMin 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox picControlExit 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox picTitleBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   0
      Width           =   4515
   End
   Begin VB.PictureBox PicSourceImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   -360
      Picture         =   "CD Interface Form.frx":1CFA
      ScaleHeight     =   3360
      ScaleWidth      =   5940
      TabIndex        =   0
      Top             =   2280
      Width           =   5940
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4875
      Top             =   510
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   -15
      ScaleHeight     =   1875
      ScaleWidth      =   4530
      TabIndex        =   4
      Top             =   300
      Width           =   4530
      Begin VB.PictureBox picLoadSkin 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3645
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   15
         Top             =   495
         Width           =   330
      End
      Begin VB.ComboBox cboTrack 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   315
         Left            =   3810
         TabIndex        =   14
         Top             =   885
         Width           =   540
      End
      Begin VB.PictureBox Picture9 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3270
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   13
         Top             =   495
         Width           =   330
      End
      Begin VB.PictureBox Picture8 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4020
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   12
         Top             =   495
         Width           =   330
      End
      Begin VB.PictureBox Picture7 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2895
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   11
         Top             =   495
         Width           =   330
      End
      Begin VB.PictureBox Picture6 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   10
         Top             =   495
         Width           =   330
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4020
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   9
         Top             =   105
         Width           =   330
      End
      Begin VB.PictureBox Picture4 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3645
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   8
         Top             =   105
         Width           =   330
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3270
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   7
         Top             =   105
         Width           =   330
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2895
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   6
         Top             =   105
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         ScaleHeight     =   322.667
         ScaleMode       =   0  'User
         ScaleWidth      =   336.286
         TabIndex        =   5
         Top             =   105
         Width           =   330
      End
      Begin VB.Label TimeWindow 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[00] 00:00"
         BeginProperty Font 
            Name            =   "Digital SF"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   435
         Left            =   420
         TabIndex        =   19
         Top             =   315
         Width           =   1605
      End
      Begin VB.Label TotalPlay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tracks: 00 CD Time: 00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   210
         Left            =   180
         TabIndex        =   18
         Top             =   1350
         Width           =   2070
      End
      Begin VB.Label TrackTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Track Time: 00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   210
         Left            =   2565
         TabIndex        =   17
         Top             =   1350
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Track No:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   270
         Left            =   2880
         TabIndex        =   16
         Top             =   915
         Width           =   915
      End
   End
End
Attribute VB_Name = "CDIface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FFSpeed As Long    'Seconds to seek for ff/rew
Dim CDPlaying As Boolean        'true if CD is currently playing
Dim CDLoaded As Boolean         'true if CD is the the player
Dim NumTracks As Integer        'number of Tracks on audio CD
Dim TrackLength() As String     'array containing length of each Track
Dim Track As Integer            'current Track
Dim Min As Integer              'current Minute on Track
Dim Sec As Integer              'current Second on Track
Dim Cmd As String               'string to hold mci command strings
Dim TotalTrackTime As String    'For Display.
Dim TotalTrackPlay As String    'For Display.
'For Registry Settings.
Public DenColor As Long, DenSkin As String
'For Moving Form.
Dim MoveFrom As Boolean, LastPoint As POINTAPI

'Load the form..
Private Sub Form_Load()
    'Get Saved Form Position
    LoadWindowPos Me
    'Get Saved Colours & Skin
    DenSettings False
    'For Background             '
    Me.Width = 4480
    Me.Height = 2175
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    Call LoadIface
    
    ' If we're already running, then quit
    If (App.PrevInstance = True) Then
        End
    End If
    ' Initialize variables
    Timer1.Enabled = False
    FFSpeed = 5
    CDLoaded = False
    ' If the cd is being used, then quit
    If (Send("open cdaudio alias cd wait shareable", True) = False) Then
        Send "Close all", False
        End
    End If
    Send "set cd time format tmsf wait", True
    Timer1.Enabled = True
End Sub

'Setup the Background Graphics..
Private Sub LoadIface()
    Call BitBlt(picTitleBar.hDC, 0, 0, 300, 20, PicSourceImage.hDC, 0, 0, SRCCOPY)
    picTitleBar.Refresh
    Call BitBlt(picMain.hDC, 0, 0, 300, 125, PicSourceImage.hDC, 0, 20, SRCCOPY)
    picMain.Refresh
    Call BitBlt(picControlMin.hDC, 0, 0, 13, 13, PicSourceImage.hDC, 301, 139, SRCCOPY)
    picControlMin.Refresh
    Call BitBlt(picControlExit.hDC, 0, 0, 13, 13, PicSourceImage.hDC, 316, 139, SRCCOPY) '422, 145, SRCCOPY)
    picControlExit.Refresh
    Call BitBlt(picLoadSkin.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 324, 70, SRCCOPY)
    picLoadSkin.Refresh

    'The Buttons                '
    Call BitBlt(Picture1.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 1, SRCCOPY)
    Picture1.Refresh
    Call BitBlt(Picture2.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 47, SRCCOPY)
    Picture2.Refresh
    Call BitBlt(Picture3.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 24, SRCCOPY)
    Picture3.Refresh
    Call BitBlt(Picture4.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 70, SRCCOPY)
    Picture4.Refresh
    Call BitBlt(Picture5.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 324, 1, SRCCOPY)
    Picture5.Refresh
    Call BitBlt(Picture6.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 93, SRCCOPY)
    Picture6.Refresh
    Call BitBlt(Picture7.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 116, SRCCOPY)
    Picture7.Refresh
    Call BitBlt(Picture8.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 324, 47, SRCCOPY)
    Picture8.Refresh
    Call BitBlt(Picture9.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 324, 24, SRCCOPY)
    Picture9.Refresh
End Sub

'Unload the Program..
'Private Sub Form_Unload(Cancel As Integer)
'    'Close all MCI devices opened by this program
'    Send "close all", False
'End Sub

'Send a MCI command string..
Private Function Send(Cmd As String, fShowError As Boolean) As Boolean
    Static rc As Long
    Static errStr As String * 200

    rc = mciSendString(Cmd, 0, 0, hwnd)
    If (fShowError And rc <> 0) Then
        mciGetErrorString rc, errStr, Len(errStr)
        MsgBox errStr
    End If
    Send = (rc = 0)
End Function

'Start the CD Playing..
Private Sub picture1_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'play_Click()
    Call BitBlt(Picture1.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 347, 1, SRCCOPY)
    Picture1.Refresh
End Sub

Private Sub picture1_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single) 'play_Click()
    Call BitBlt(Picture1.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 1, SRCCOPY)
    Picture1.Refresh
    If X > 0 And X < Picture1.Width And Y > 0 And Y < Picture1.Height Then
        If (CDLoaded) Then
            Send "play cd", True
            CDPlaying = True
        End If
    End If
End Sub

'Stop the CD playing..
Private Sub picture5_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture5.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 370, 1, SRCCOPY)
    Picture5.Refresh
End Sub

Private Sub picture5_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(Picture5.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 324, 1, SRCCOPY)
    Picture5.Refresh
    If X > 0 And X < Picture5.Width And Y > 0 And Y < Picture5.Height Then
        If CDPlaying = True Then
            Send "stop cd wait", True
            Cmd = "seek cd to " & Track
            Send Cmd, True
            CDPlaying = False
            Update
            
        End If
    End If
End Sub
'Pause the CD..
Private Sub picture4_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture4.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 347, 70, SRCCOPY)
    Picture4.Refresh
End Sub

Private Sub picture4_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture4.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 70, SRCCOPY)
    Picture4.Refresh
    If X > 0 And X < Picture4.Width And Y > 0 And Y < Picture4.Height Then
        If CDPlaying = True Then
            Send "pause cd", True
            CDPlaying = False
            Update
        End If
    End If
End Sub

'Goto Next Track..
Private Sub picture3_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture3.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 347, 24, SRCCOPY)
    Picture3.Refresh
End Sub

Private Sub picture3_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture3.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 24, SRCCOPY)
    Picture3.Refresh
    If X > 0 And X < Picture3.Width And Y > 0 And Y < Picture3.Height Then
        If (Track < NumTracks) Then
            If (CDPlaying) Then
                Cmd = "play cd from " & Track + 1
                Send Cmd, True
            Else
                If (CDLoaded) Then
                    Cmd = "seek cd to " & Track + 1
                    Send Cmd, True
                End If
            End If
        Else
            If (CDLoaded) Then
                Send "seek cd to 1", True
            End If
        End If
        Update
    End If
End Sub

'Goto previous Track..
Private Sub picture2_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture2.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 347, 47, SRCCOPY)
    Picture2.Refresh
End Sub

Private Sub picture2_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture2.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 47, SRCCOPY)
    Picture2.Refresh
    If X > 0 And X < Picture2.Width And Y > 0 And Y < Picture2.Height Then
        Dim from As String
        If (Min = 0 And Sec = 0) Then
            If (Track > 1) Then
                from = CStr(Track - 1)
            Else
                from = CStr(NumTracks)
            End If
        Else
            from = CStr(Track)
        End If
        If (CDPlaying) Then
            Cmd = "play cd from " & from
            Send Cmd, True
        Else
            If (CDLoaded) Then
                Cmd = "seek cd to " & from
                Send Cmd, True
            End If
        End If
        Update
    End If
End Sub

'Fast forward..
Private Sub picture7_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture7.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 347, 116, SRCCOPY)
    Picture7.Refresh
End Sub

Private Sub picture7_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture7.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 116, SRCCOPY)
    Picture7.Refresh
    If X > 0 And X < Picture7.Width And Y > 0 And Y < Picture7.Height Then
        If (CDPlaying) Then
            Dim s As String * 40
            Send "set cd time format milliSeconds", True
            mciSendString "status cd position wait", s, Len(s), 0
            Cmd = "play cd from " & CStr(CLng(s) + FFSpeed * 1000)
            mciSendString Cmd, 0, 0, 0
            Send "set cd time format tmsf", True
        Else
            If (CDLoaded) Then
                If (CDPlaying) Then
                    Cmd = "seek cd to " & CStr(CLng(s) + FFSpeed * 1000)
                End If
            End If
        End If
        Update
    End If
End Sub

'Rewind the CD..
Private Sub Picture6_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture6.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 347, 93, SRCCOPY)
    Picture6.Refresh
End Sub

Private Sub Picture6_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture6.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 301, 93, SRCCOPY)
    Picture6.Refresh
    If X > 0 And X < Picture6.Width And Y > 0 And Y < Picture6.Height Then
        If (CDPlaying) Then
            Dim s As String * 40
            Send "set cd time format milliSeconds", True
            mciSendString "status cd position wait", s, Len(s), 0
            Cmd = "play cd from " & CStr(CLng(s) - FFSpeed * 1000)
            mciSendString Cmd, 0, 0, 0
            Send "set cd time format tmsf", True
        Else
            If (CDLoaded) Then
                If (CDPlaying) Then
                    Cmd = "seek cd to " & CStr(CLng(s) - FFSpeed * 1000)
                End If
            End If
        End If
        Update
    End If
End Sub

'Eject the CD..
Private Sub picture9_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture9.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 370, 24, SRCCOPY)
    Picture9.Refresh
End Sub

Private Sub picture9_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture9.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 324, 24, SRCCOPY)
    Picture9.Refresh
    If X > 0 And X < Picture9.Width And Y > 0 And Y < Picture9.Height Then
        Send "stop cd wait", True
        Send "set cd door open", True
        Update
    End If
End Sub

'Show About Form..
Private Sub picture8_mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture8.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 370, 47, SRCCOPY)
    Picture8.Refresh
End Sub

Private Sub picture8_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single) 'stopbtn_Click()
    Call BitBlt(Picture8.hDC, 0, 0, 22, 22, PicSourceImage.hDC, 324, 47, SRCCOPY)
    Picture8.Refresh
    If X > 0 And X < Picture8.Width And Y > 0 And Y < Picture8.Height Then
        CDAbout.Show
    End If
End Sub

'Exit/Close Button..
Private Sub picControlExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picControlExit.hDC, 0, 0, 13, 13, PicSourceImage.hDC, 316, 153, SRCCOPY)
    picControlExit.Refresh
End Sub

Private Sub picControlExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picControlExit.hDC, 0, 0, 13, 13, PicSourceImage.hDC, 316, 139, SRCCOPY)
    picControlExit.Refresh
    If X > 0 And X < picControlExit.Width And Y > 0 And Y < picControlExit.Height Then
        If CDPlaying = True Then
            Send "stop cd wait", True
            Cmd = "seek cd to " & Track
            Send Cmd, True
            CDPlaying = False
            Update
            Send "Close all", True
            SaveWindowPos Me
            Unload Me
        Else
            Send "Close all", True
            SaveWindowPos Me
            Unload Me
        End If
    End If
End Sub

'Minimise Button..
Private Sub picControlMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picControlMin.hDC, 0, 0, 13, 13, PicSourceImage.hDC, 301, 153, SRCCOPY)
    picControlMin.Refresh
End Sub

Private Sub picControlMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picControlMin.hDC, 0, 0, 13, 13, PicSourceImage.hDC, 301, 139, SRCCOPY)
    picControlMin.Refresh
    If X > 0 And X < picControlMin.Width And Y > 0 And Y < picControlMin.Height Then
        Me.WindowState = 1
    End If
End Sub

'Moving The Form..
Private Sub picTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    MoveFrom = True
End Sub

Private Sub picTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not MoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.X - LastPoint.X) * iTPPX&
    iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub picTitleBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveFrom = False
End Sub

'Set the fast-forward speed..
Private Sub FFSpeed_Click()
    Dim s As String
    s = InputBox("Enter the new speed in Seconds", "Fast Forward Speed", CStr(FFSpeed))
    If IsNumeric(s) Then
        FFSpeed = CLng(s)
    End If
End Sub

'Timer Update..
Private Sub Timer1_Timer()
    Update
End Sub

'Update the display and state variables..
Private Sub Update()
    Static s As String * 30

    ' Check if CD is in the player
    mciSendString "status cd media present", s, Len(s), 0
    If (CBool(s)) Then
        ' Enable all the controls, get CD information
        If (CDLoaded = False) Then
            mciSendString "status cd number of Tracks wait", s, Len(s), 0
            NumTracks = CInt(Mid$(s, 1, 2))
        
            ' If CD only has 1 Track, then it's probably a data CD
            If (NumTracks = 1) Then
                Exit Sub
            End If
        
            mciSendString "status cd length wait", s, Len(s), 0
            TotalTrackPlay = "Tracks: " & NumTracks & "  CD Time: " & Left(s, 5)
            TotalPlay.Caption = TotalTrackPlay
            
            ReDim TrackLength(1 To NumTracks)
            Dim i As Integer
            For i = 1 To NumTracks
                Cmd = "status cd length Track " & i
                mciSendString Cmd, s, Len(s), 0
                TrackLength(i) = s
            Next
            
            '####################################
            ' Fill list of Track Nos.
            Dim it As Integer
            cboTrack.Clear
            For it = 1 To NumTracks
                cboTrack.AddItem it
            Next it
            cboTrack.Text = cboTrack.List(0)
            '####################################

            Send "seek cd to 1", True
            CDLoaded = True
        End If

        ' Update the Track time display
        mciSendString "status cd position", s, Len(s), 0
        Track = CInt(Mid$(s, 1, 2))
        Min = CInt(Mid$(s, 4, 2))
        Sec = CInt(Mid$(s, 7, 2))
        TimeWindow.Caption = "[" & Format(Track, "00") & "] " & Format(Min, "00") _
            & ":" & Format(Sec, "00")
        TotalTrackTime = "Track Time: " & Left(TrackLength(Track), 5)
        TrackTime.Caption = TotalTrackTime
        cboTrack.Text = cboTrack.List(Track - 1)
        
        ' Check if CD is playing
        mciSendString "status cd mode", s, Len(s), 0
        CDPlaying = (Mid$(s, 1, 7) = "playing")
    Else
        'eject.Enabled = False
        ' Disable all the controls, clear the display
        If (CDLoaded = True) Then
            CDLoaded = False
            CDPlaying = False
            TotalPlay.Caption = ""
            TrackTime.Caption = ""
            TimeWindow.Caption = ""
        End If
    End If
End Sub

Private Sub cboTrack_click()
    If (CDLoaded) Then
        'Set cboTrack value first
        cboTrack.ListIndex = Val(cboTrack.Text) - 1
        If (Track <= NumTracks) Then
            If (CDPlaying) Then
                Cmd = "play cd from " & Val(cboTrack.Text)
                Send Cmd, True
            Else
                Cmd = "seek cd to " & Val(cboTrack.Text)
                Send Cmd, True
                Send "play cd", True
                CDPlaying = True
            End If
        End If
    Else
        Send "seek cd to 1", True
    End If
    Update
End Sub


Private Sub picLoadSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picLoadSkin.hDC, 0, 0, 70, 23, PicSourceImage.hDC, 370, 70, SRCCOPY)
    picLoadSkin.Refresh
End Sub

Private Sub picLoadSkin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long
    If GetCapture() <> picLoadSkin.hwnd Then
        Ret = SetCapture(picLoadSkin.hwnd)
        Call BitBlt(picLoadSkin.hDC, 0, 0, 70, 23, PicSourceImage.hDC, 324, 70, SRCCOPY)
        picLoadSkin.Refresh
        Me.MousePointer = 99
    End If
    If X > 0 And X < picLoadSkin.Width And Y > 0 And Y < picLoadSkin.Height Then
        CurrentX = X
        CurrentY = Y
    Else
        If GetCapture() = picLoadSkin.hwnd Then
            Ret = ReleaseCapture()
            Call BitBlt(picLoadSkin.hDC, 0, 0, 70, 23, PicSourceImage.hDC, 324, 70, SRCCOPY)
            picLoadSkin.Refresh
            Me.MousePointer = 0
        End If
    End If
End Sub

Private Sub picLoadSkin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Error_Event:
    Call BitBlt(picLoadSkin.hDC, 0, 0, 70, 23, PicSourceImage.hDC, 324, 70, SRCCOPY)
    picLoadSkin.Refresh
    If X > 0 And X < picLoadSkin.Width And Y > 0 And Y < picLoadSkin.Height Then
        cdgIface.InitDir = App.Path
        cdgIface.Filter = "Skin Files (*.bmp)|*.bmp"
        cdgIface.CancelError = True
        cdgIface.ShowOpen
        If cdgIface.filename <> "" Then
            DenSkin = cdgIface.filename
        End If
        PicSourceImage.Picture = LoadPicture(DenSkin)
        Call LoadIface
        DenSettings True
    End If
Error_Event:
    Exit Sub
End Sub

Private Sub TimeWindow_Click()
    With cdgIface
        .CancelError = True
        On Error GoTo ColorErrHandler
        .ShowColor
        TimeWindow.ForeColor = .Color
        TotalPlay.ForeColor = .Color
        TrackTime.ForeColor = .Color
        cboTrack.ForeColor = .Color
        Label1.ForeColor = .Color
        DenColor = .Color
    End With
    DenSettings (True)
ColorErrHandler:
End Sub

Public Sub DenSettings(DenSet As Boolean)
    '>> DenSkin = "Path & Filename" from Load Dialog
    If (DenSet) Then
        'Save Program Color Setting. (This Works).
        SaveColSet Me, DenColor
        If DenSkin <> "" Then
            'Save Program Skin Setting. (This Works).
            SaveSkinSet Me, DenSkin
        End If
    Else
        'Load Program Color Setting. (This Works).
        LoadColSet Me
        'Load Program Skin Setting. (This Works).
        LoadSkinSet Me
    End If
End Sub

