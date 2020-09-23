VERSION 5.00
Begin VB.Form frmContents 
   Caption         =   "Contents"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPocetni.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   7050
   Begin VB.PictureBox picKomande 
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
      Height          =   5415
      Left            =   135
      ScaleHeight     =   5415
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   90
      Width           =   6585
      Begin VB.CommandButton Command5 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   4680
         Width           =   1410
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2490
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmPocetni.frx":0442
         Top             =   2790
         Width           =   4830
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Employees"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   4185
         Width           =   1410
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Articles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   3690
         Width           =   1410
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Clients"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   3195
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Orders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   2700
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Demo application"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   405
         Left            =   630
         TabIndex        =   7
         Top             =   1350
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Small Corporation"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   26.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   750
         Left            =   585
         TabIndex        =   6
         Top             =   585
         Width           =   5160
      End
   End
End
Attribute VB_Name = "frmContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   frmDetail.SQLFile = "roba_def"
   frmDetail.Show
   frmDetail.SetMode "Orders"
End Sub

Private Sub Command2_Click()
   frmDetail.SQLFile = "klijenti_def"
   frmDetail.Show
   frmDetail.SetMode "Clients"
End Sub

Private Sub Command3_Click()
   frmDetail.SQLFile = "artikli_def"
   frmDetail.Show
   frmDetail.SetMode "Articles"
End Sub

Private Sub Command5_Click()
   Unload frmMain
End Sub

Private Sub Form_Load()
   Me.Width = 7000
   Me.Height = 6000
End Sub

Private Sub Form_Resize()
On Error Resume Next
   With picKomande
      .Move Me.ScaleWidth / 2 - .Width / 2, Me.ScaleHeight / 2 - .Height / 2
   End With
End Sub
