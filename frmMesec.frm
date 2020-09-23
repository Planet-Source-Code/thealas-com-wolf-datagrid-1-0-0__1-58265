VERSION 5.00
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date"
   ClientHeight    =   3300
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3690
   Icon            =   "frmMesec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   246
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtGodina 
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   1305
      Width           =   3345
   End
   Begin VB.CheckBox chOtprema 
      Caption         =   "Filter shipment date, not payment date"
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   2340
      Width           =   3300
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   2295
      TabIndex        =   2
      Top             =   2790
      Width           =   1320
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Accept"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   900
      TabIndex        =   1
      Top             =   2790
      Width           =   1320
   End
   Begin VB.ComboBox cboMesec 
      Height          =   315
      ItemData        =   "frmMesec.frx":000C
      Left            =   180
      List            =   "frmMesec.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   3345
   End
   Begin VB.Label Label2 
      Caption         =   "For Year:"
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   1035
      Width           =   3930
   End
   Begin VB.Label Label1 
      Caption         =   "Month"
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   135
      Width           =   3930
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public sCol As String
Public dt1 As String, dt2 As String
Public sql_dt1 As String, sql_dt2 As String

Private Sub chOtprema_Click()
   If chOtprema.Value = 1 Then sCol = "Shipment Date" Else sCol = "Payment Date"
End Sub

Private Sub cmdOk_Click()
   
   sql_dt1 = DateSerial(txtGodina.Text, cboMesec.ListIndex + 1, 1)
   sql_dt2 = DateAdd("m", 1, sql_dt1)

   dt1 = sql_dt1
   dt2 = sql_dt2

   dt1 = Year(dt1) & ", " & Month(dt1) & ", " & Day(dt1)
   dt2 = Year(dt2) & ", " & Month(dt2) & ", " & Day(dt2)

   sql_dt1 = Month(sql_dt1) & "/" & Day(sql_dt1) & "/" & Year(sql_dt1)
   sql_dt2 = Month(sql_dt2) & "/" & Day(sql_dt2) & "/" & Year(sql_dt2)

   Unload Me
End Sub

Private Sub cmdCancel_Click()
   dt1 = ""
   dt2 = ""
   Unload Me
End Sub



Private Sub Form_Load()
   With cboMesec
      Dim i&, s$
      For i = 1 To 12
         Dim dt As Date: dt = DateSerial(Now, i, 1)
         s = FormatDateTime(dt, vbLongDate)
         s = Format(s, "mmmm")
         s = UCase(Left(s, 1)) & Right(s, Len(s) - 1)
         .AddItem s
      Next i
   End With
   
   txtGodina.Text = Year(Now)
   cboMesec.ListIndex = Month(Now) - 1
   sCol = "Payment Date"
End Sub


