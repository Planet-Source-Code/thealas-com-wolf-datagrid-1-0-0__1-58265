VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FEF3E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Wolf DataBase Controls"
   ClientHeight    =   3870
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5490
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   366
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3780
      TabIndex        =   3
      Top             =   3015
      Width           =   1590
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":000C
      Height          =   1050
      Left            =   2835
      TabIndex        =   5
      Top             =   1575
      Width           =   2355
   End
   Begin VB.Label lblVer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2745
      TabIndex        =   4
      Top             =   3465
      Width           =   2670
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(C) 2004-2005"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2790
      TabIndex        =   2
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright TheAlas Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2790
      TabIndex        =   1
      Top             =   450
      Width           =   2070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wolf DataBase Controls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2790
      TabIndex        =   0
      Top             =   135
      Width           =   2025
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3720
      Left            =   90
      Picture         =   "frmAbout.frx":007E
      Top             =   45
      Width           =   2565
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   lblVer.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

'Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   ShellExecute GetDesktopWindow(), "open", "http://" & Label4.Caption, "", "http", 5
'End Sub

