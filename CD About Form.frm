VERSION 5.00
Begin VB.Form CDAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1845
   ClientLeft      =   3675
   ClientTop       =   3195
   ClientWidth     =   4155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "CD About Form.frx":0000
   ScaleHeight     =   1845
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton AboutExit 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Default         =   -1  'True
      DownPicture     =   "CD About Form.frx":2560
      Height          =   465
      Left            =   3030
      Picture         =   "CD About Form.frx":29A2
      TabIndex        =   0
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label Comments 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks"
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   60
      TabIndex        =   5
      Top             =   1560
      Width           =   2730
   End
   Begin VB.Label Company 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dennis Hallman"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   1275
      Width           =   2715
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   1065
      Width           =   2730
   End
   Begin VB.Label Version 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version Number"
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   525
      Width           =   4035
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Program Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   4020
   End
End
Attribute VB_Name = "CDAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Load the Form..
Private Sub Form_Load()
    CDAbout.Left = CDIface.Left + (CDIface.Width - CDAbout.Width) / 2
    CDAbout.Top = CDIface.Top + (CDIface.Height - CDAbout.Height) / 2
    CDAbout.Title = App.Title
    CDAbout.Version = "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbLf & App.LegalCopyright
    CDAbout.Company = App.CompanyName
    CDAbout.Comments = App.Comments
End Sub

'Unload the Form..
Private Sub AboutExit_Click()
    Unload Me
End Sub

