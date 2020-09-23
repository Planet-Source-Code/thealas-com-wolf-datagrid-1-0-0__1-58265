VERSION 5.00
Object = "*\AWolfDataBaseSystem.vbp"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDemoSimple 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simplest Possible Demo"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDemoSimple.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   90
      Top             =   1665
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DataSource"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin WolfDataBaseControls.DataGrid DataGrid 
      Height          =   2940
      Left            =   90
      TabIndex        =   0
      Top             =   2070
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   5186
      CurrentRecordColor=   0
      AllowDelete     =   -1  'True
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NullString      =   ""
      IndicatorWidth  =   18
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   $"frmDemoSimple.frx":000C
      Height          =   1410
      Left            =   315
      TabIndex        =   1
      Top             =   90
      Width           =   5010
   End
End
Attribute VB_Name = "frmDemoSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Adodc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
   Adodc.Caption = Adodc.Recordset.Fields("Name").Value
End Sub


Private Sub Form_Load()
   With Adodc
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\SmallCorporation.mdb" & ";Persist Security Info=False"
      .RecordSource = "SELECT * FROM tblClients"
      .Refresh
      Set DataGrid.DataSource = .Recordset
   End With
End Sub
