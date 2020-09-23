VERSION 5.00
Begin VB.Form frmSql 
   Caption         =   "SQL Syntax"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSql.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   5055
   Begin VB.CommandButton cmdSacuvaj 
      Height          =   420
      Left            =   2610
      Picture         =   "frmSql.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Save file"
      Top             =   2790
      Width           =   465
   End
   Begin VB.ComboBox cboFiles 
      Height          =   330
      ItemData        =   "frmSql.frx":09CC
      Left            =   45
      List            =   "frmSql.frx":09CE
      TabIndex        =   2
      Text            =   "[SQL Files List]"
      Top             =   2835
      Width           =   2535
   End
   Begin VB.CommandButton cmdIzvrsi 
      Caption         =   "&Execute"
      Height          =   420
      Left            =   3105
      TabIndex        =   1
      ToolTipText     =   "Execute command to currently active recordset"
      Top             =   2790
      Width           =   1545
   End
   Begin VB.TextBox txtSql 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4650
   End
End
Attribute VB_Name = "frmSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFiles_Click()
   txtSql = OpenSqlFile(cboFiles.Text)
End Sub

Public Sub cmdIzvrsi_Click()
On Error GoTo e
   frmDetail.EnableEdit True
   frmDetail.DataGrid.Enabled = True
   
   adoActiveAdo.RecordSource = txtSql
   adoActiveAdo.CursorType = adOpenKeyset
   adoActiveAdo.Refresh
   adoActiveAdo.Recordset.Requery
   Set adoActiveGrid.DataSource = adoActiveAdo.Recordset
Exit Sub
e:
   MsgBox "Error in SQL syntax!", vbCritical, "Greska"
   frmDetail.EnableEdit False
   frmDetail.DataGrid.Enabled = False
End Sub

Private Sub cmdSacuvaj_Click()
   Dim ind&
   SelectComboItem cboFiles, cboFiles.Text
   ind = cboFiles.ListIndex
   
   SaveSqlFile cboFiles.Text, txtSql.Text

   cboFiles.Clear
   
   Dim file$: file = Dir(App.Path & "\*.sql")
   While file <> ""
      cboFiles.AddItem Left(file, Len(file) - 4)
      file = Dir
   Wend
   
   cboFiles.ListIndex = ind
End Sub

Private Sub Form_Load()
   Dim file$
   
   Me.Width = 6000: Me.Height = 4500
   
   file = Dir(App.Path & "\*.sql")
   While file <> ""
      cboFiles.AddItem Left(file, Len(file) - 4)
      file = Dir
   Wend
End Sub

Private Sub Form_Resize()
On Error Resume Next
   With txtSql
      .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmdIzvrsi.Height
   End With
   
   With cmdIzvrsi
      .Move Me.ScaleWidth - .Width, Me.ScaleHeight - .Height
   End With
   
   With cmdSacuvaj
      .Move cmdIzvrsi.Left - .Width - 45, cmdIzvrsi.Top
   End With
   
   With cboFiles
      .Move 45, Me.ScaleHeight - .Height - 45, cmdSacuvaj.Left - .Left - 45
   End With
   
End Sub
