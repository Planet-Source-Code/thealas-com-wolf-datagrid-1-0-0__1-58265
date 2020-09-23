VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "*\AWolfDataBaseSystem.vbp"
Begin VB.Form frmDetail 
   Caption         =   "Orders"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNarucenaRoba.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   WindowState     =   2  'Maximized
   Begin WolfDataBaseControls.DataGrid DataGrid 
      Height          =   2040
      Left            =   90
      TabIndex        =   55
      Top             =   2880
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   3598
      HideKeyColumns  =   -1  'True
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
   Begin VB.PictureBox picCommands 
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   7695
      ScaleHeight     =   1905
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   45
      Width           =   1815
      Begin VB.CommandButton cmdIzmeni 
         Caption         =   "Finish Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   42
         ToolTipText     =   "Changes the data of selected record"
         Top             =   0
         Width           =   1725
      End
      Begin VB.CommandButton cmdDodaj 
         Caption         =   "Add New"
         Height          =   375
         Left            =   0
         TabIndex        =   0
         ToolTipText     =   "Entered data is written to a newly added record"
         Top             =   405
         Width           =   1725
      End
      Begin VB.CommandButton cmdOcisti 
         Caption         =   "Clear Fields"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Cleans all fields"
         Top             =   810
         Width           =   1725
      End
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   0
      Top             =   2250
      Width           =   7635
      _ExtentX        =   13467
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
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox picRoba 
      BorderStyle     =   0  'None
      Height          =   2085
      Left            =   45
      ScaleHeight     =   2085
      ScaleWidth      =   7620
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   7620
      Begin VB.TextBox txtKP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1050
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "Commercial Packages"
         Top             =   855
         Width           =   915
      End
      Begin VB.TextBox txtTP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1050
         TabIndex        =   13
         Text            =   "0"
         ToolTipText     =   "Transport Packages"
         Top             =   1215
         Width           =   915
      End
      Begin VB.TextBox txtRabat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1050
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "Discount Percentage"
         Top             =   1575
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1995
         Picture         =   "frmNarucenaRoba.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Calculate CP"
         Top             =   1215
         Width           =   285
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4560
         Picture         =   "frmNarucenaRoba.frx":049E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Calculate automaticaly"
         Top             =   1215
         Width           =   285
      End
      Begin VB.TextBox txtPotrazuje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   1575
         Width           =   915
      End
      Begin VB.TextBox txtDuguje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1215
         Width           =   915
      End
      Begin VB.TextBox txtNaplaceno 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   855
         Width           =   915
      End
      Begin VB.ComboBox cboArtikal 
         Height          =   345
         Left            =   2520
         TabIndex        =   6
         Text            =   "cboArtikal"
         Top             =   270
         Width           =   2445
      End
      Begin VB.ComboBox cboKlijent 
         Height          =   345
         ItemData        =   "frmNarucenaRoba.frx":04FA
         Left            =   0
         List            =   "frmNarucenaRoba.frx":04FC
         TabIndex        =   5
         Top             =   270
         Width           =   2400
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4545
         Picture         =   "frmNarucenaRoba.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Set balance to zero, full payment"
         Top             =   855
         Width           =   285
      End
      Begin Crystal.CrystalReport crOtpremnica 
         Left            =   180
         Top             =   900
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
         ReportFileName  =   "D:\SpiderDataBaseSystem\otpremnica.rpt"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtOtprema 
         Height          =   330
         Left            =   5175
         TabIndex        =   15
         Top             =   675
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49938432
         CurrentDate     =   38315
      End
      Begin MSComCtl2.DTPicker dtNaplata 
         Height          =   330
         Left            =   5175
         TabIndex        =   16
         Top             =   1395
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   582
         _Version        =   393216
         Format          =   49938432
         CurrentDate     =   38315
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "CP:"
         Height          =   225
         Left            =   15
         TabIndex        =   26
         ToolTipText     =   "Komercijalnih pakovanja"
         Top             =   855
         Width           =   990
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "TP:"
         Height          =   225
         Left            =   15
         TabIndex        =   25
         ToolTipText     =   "Transportnih pakovanja"
         Top             =   1215
         Width           =   990
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Rebate:"
         Height          =   225
         Left            =   15
         TabIndex        =   24
         ToolTipText     =   "Rabat za ovog kupca u procentima"
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Payment Date:"
         Height          =   225
         Left            =   5175
         TabIndex        =   23
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Shipping Date:"
         Height          =   225
         Left            =   5175
         TabIndex        =   22
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Demand:"
         Height          =   225
         Left            =   2520
         TabIndex        =   21
         Top             =   1575
         Width           =   1050
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Owe:"
         Height          =   225
         Left            =   2505
         TabIndex        =   20
         Top             =   1215
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Payed:"
         Height          =   225
         Left            =   2505
         TabIndex        =   19
         Top             =   855
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Article:"
         Height          =   225
         Left            =   2565
         TabIndex        =   18
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Client:"
         Height          =   225
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   525
      End
      Begin VB.Shape Shape1 
         Height          =   1185
         Left            =   2475
         Top             =   765
         Width           =   2490
      End
      Begin VB.Shape Shape2 
         Height          =   1185
         Left            =   0
         Top             =   765
         Width           =   2400
      End
      Begin VB.Shape Shape3 
         Height          =   1680
         Left            =   5040
         Top             =   270
         Width           =   2445
      End
   End
   Begin VB.PictureBox picKlijenti 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   7665
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   7665
      Begin TabDlg.SSTab SSTab1 
         Height          =   2130
         Left            =   45
         TabIndex        =   28
         Top             =   45
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   3757
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Osnovne informacije"
         TabPicture(0)   =   "frmNarucenaRoba.frx":055A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label14"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label13"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label12"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label11"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtRabat2"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cboKomercijalista"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtAdresa"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtImeKlijenta"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Komentar/Dodatno"
         TabPicture(1)   =   "frmNarucenaRoba.frx":0576
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtKomentar"
         Tab(1).ControlCount=   1
         Begin VB.TextBox txtImeKlijenta 
            BackColor       =   &H00C0FFFF&
            Height          =   330
            Left            =   1035
            TabIndex        =   31
            Top             =   450
            Width           =   3435
         End
         Begin VB.TextBox txtAdresa 
            Height          =   330
            Left            =   1035
            TabIndex        =   33
            Top             =   855
            Width           =   3435
         End
         Begin VB.ComboBox cboKomercijalista 
            Height          =   345
            Left            =   1035
            TabIndex        =   35
            Top             =   1260
            Width           =   3480
         End
         Begin VB.Frame Frame1 
            Caption         =   "Phones:"
            Height          =   1635
            Left            =   4590
            TabIndex        =   30
            Top             =   360
            Width           =   2940
            Begin VB.TextBox txtTel 
               Height          =   330
               Index           =   2
               Left            =   135
               TabIndex        =   41
               Top             =   1170
               Width           =   2670
            End
            Begin VB.TextBox txtTel 
               Height          =   330
               Index           =   1
               Left            =   135
               TabIndex        =   40
               Top             =   720
               Width           =   2670
            End
            Begin VB.TextBox txtTel 
               Height          =   330
               Index           =   0
               Left            =   135
               TabIndex        =   39
               Top             =   315
               Width           =   2670
            End
         End
         Begin VB.TextBox txtRabat2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0FF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1035
            TabIndex        =   37
            Text            =   "0"
            Top             =   1665
            Width           =   870
         End
         Begin VB.TextBox txtKomentar 
            Height          =   1680
            Left            =   -74955
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   360
            Width           =   7485
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   45
            TabIndex        =   38
            Top             =   495
            Width           =   945
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Address:"
            Height          =   225
            Left            =   45
            TabIndex        =   36
            ToolTipText     =   "Lokacija poslovnog objekta"
            Top             =   900
            Width           =   945
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Agent:"
            Height          =   225
            Left            =   45
            TabIndex        =   34
            ToolTipText     =   "Agent of this client, employee"
            Top             =   1305
            Width           =   945
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Rebate:"
            Height          =   225
            Left            =   45
            TabIndex        =   32
            ToolTipText     =   "Default rebate used for this client orders"
            Top             =   1710
            Width           =   945
         End
      End
   End
   Begin VB.PictureBox picArtikli 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   0
      ScaleHeight     =   2220
      ScaleWidth      =   7665
      TabIndex        =   43
      Top             =   0
      Width           =   7665
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3465
         Picture         =   "frmNarucenaRoba.frx":0592
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Calculate atutomaticaly"
         Top             =   1710
         Width           =   285
      End
      Begin VB.TextBox txtNapomena 
         Height          =   1590
         Left            =   3870
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Top             =   405
         Width           =   3570
      End
      Begin VB.TextBox txtTPArtikla 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   50
         Text            =   "0"
         ToolTipText     =   "Transportnih pakovanja"
         Top             =   1305
         Width           =   915
      End
      Begin VB.TextBox txtCenaTP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   48
         Text            =   "0.00"
         Top             =   1710
         Width           =   915
      End
      Begin VB.TextBox txtCenaArtikla 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2535
         TabIndex        =   46
         Text            =   "0.00"
         Top             =   900
         Width           =   915
      End
      Begin VB.TextBox txtImeArtikla 
         BackColor       =   &H00FFC0FF&
         Height          =   330
         Left            =   135
         TabIndex        =   45
         Top             =   405
         Width           =   3615
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Comment:"
         Height          =   225
         Left            =   3870
         TabIndex        =   52
         Top             =   135
         Width           =   870
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "TP:"
         Height          =   225
         Left            =   1485
         TabIndex        =   51
         ToolTipText     =   "Transport packages"
         Top             =   1305
         Width           =   990
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "TP Price:"
         Height          =   225
         Left            =   1440
         TabIndex        =   49
         Top             =   1710
         Width           =   1050
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Price:"
         Height          =   225
         Left            =   1440
         TabIndex        =   47
         Top             =   900
         Width           =   1050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Article name:"
         Height          =   225
         Left            =   135
         TabIndex        =   44
         Top             =   135
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SQLFile As String
Public AddingNew As Boolean

Public crMode As Long

Private Sub Adodc_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
   
   With Err
      .Description = Description
      .HelpContext = HelpContext
      .Source = Source
      .Number = ErrorNumber
   End With
   
   frmSql.Show
   frmSql.WindowState = vbNormal
   
   'RaiseErr ERR_SQL, Err, False
   
   fCancelDisplay = True
End Sub


Private Sub Adodc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
   If crMode = 1 Then
      Adodc.Caption = Adodc.Recordset.Fields("Article").Value
      
      With Adodc.Recordset
         SelectComboItem cboArtikal, .Fields("Article").Value, Knjiga("Articles")
         SelectComboItem cboKlijent, .Fields("Client").Value, Knjiga("Clients")
         
         txtRabat.Text = .Fields("Rebate")
         
         txtTP.Text = .Fields("TP")
         txtKP.Text = .Fields("CP")
         txtNaplaceno.Text = .Fields("Payed")
         txtDuguje.Text = .Fields("Owing")
         txtPotrazuje.Text = .Fields("Demands")
         
         txtTP_LostFocus
         txtKP_LostFocus
         txtNaplaceno_LostFocus
         txtDuguje_LostFocus
         txtPotrazuje_LostFocus
         
         dtOtprema.Value = .Fields("Shipment Date")
         dtNaplata.Value = .Fields("Payment Date")
      End With
   ElseIf crMode = 2 Then
      Adodc.Caption = Adodc.Recordset.Fields("Name").Value
      
      With Adodc.Recordset
         SelectComboItem cboKomercijalista, .Fields("Employee").Value, Knjiga("Employees")
         
         txtTel(0) = ""
         txtTel(1) = ""
         txtTel(2) = ""
         txtAdresa = ""
         
         txtRabat2.Text = .Fields("Rebate")
         txtAdresa.Text = .Fields("Address")
         txtImeKlijenta.Text = .Fields("Name")
         txtTel(0).Text = .Fields("Phone1")
         txtTel(1).Text = .Fields("Phone2")
         txtTel(2).Text = .Fields("Phone3")

         
         txtRabat2_LostFocus
         txtKP_LostFocus
      End With
   ElseIf crMode = 3 Then
      Adodc.Caption = Adodc.Recordset.Fields("Article").Value
      
      With Adodc.Recordset
         txtImeArtikla = ""
         txtNapomena.Text = ""
         txtTPArtikla.Text = "0"
         txtCenaTP.Text = "0"
         txtCenaArtikla.Text = "0"
         
         txtImeArtikla = .Fields("Article")
         txtNapomena.Text = .Fields("Comment")
         txtTPArtikla.Text = .Fields("TP")
         txtCenaTP.Text = .Fields("Price TP")
         txtCenaArtikla.Text = .Fields("Price")
         
         txtTPArtikla_LostFocus
         txtCenaTP_LostFocus
         txtCenaArtikla_LostFocus
      End With
   End If
End Sub


Private Sub cboArtikal_LostFocus()
   If CheckForComboExistance(cboArtikal, Knjiga("Articles")) = -1 Then _
      MsgBox "Entered value is not valid, it does not exist in table!", vbCritical, "Greska": _
      cboArtikal.SetFocus
End Sub

Private Sub cboKlijent_LostFocus()
   If CheckForComboExistance(cboKlijent, Knjiga("Clients")) = -1 Then _
      MsgBox "Entered value is not valid, it does not exist in table!", vbCritical, "Greska": _
      cboKlijent.SetFocus: Exit Sub
   txtRabat.Text = Knjiga("Clients").RecordSource.Fields("Rebate")
End Sub

Private Sub cboKomercijalista_LostFocus()
   If CheckForComboExistance(cboKomercijalista, Knjiga("Employees")) = -1 Then _
      MsgBox "Entered value is not valid, it does not exist in table!", vbCritical, "Greska": _
      cboKomercijalista.SetFocus
End Sub

Private Sub Command1_Click()
   Dim npl&, znpl&
   
   With Knjiga("Articles").RecordSource
      znpl = .Fields("Price") * txtKP.Text
      npl = txtNaplaceno.Text
   End With
   
   If npl > znpl Then
      txtPotrazuje.Text = npl - znpl
   Else
      txtDuguje.Text = znpl - npl
   End If
   
   txtDuguje_LostFocus
   txtPotrazuje_LostFocus
End Sub

Private Sub Command2_Click()
   txtKP.Text = _
      Knjiga("Articles").RecordSource.Fields("TP").Value * txtTP.Text
   txtKP_LostFocus
End Sub

Private Sub cmdIzmeni_Click()
   If cboArtikal.Text = "" Then MsgBox "Article field is empty!": Exit Sub
   If cboKlijent.Text = "" Then MsgBox "Client field is empty!": Exit Sub
   
   Dim sql$, pos&
   With Adodc.Recordset
      pos = .AbsolutePosition
      
      If crMode = 1 Then
         sql = "UPDATE tblOrders SET " & _
         "TP = " & CLng(txtTP.Text) & ", " & _
         "CP = " & CLng(txtKP.Text) & ", " & _
         "Payed = " & CSng(txtNaplaceno.Text) & ", " & _
         "Owing = " & CSng(txtDuguje.Text) & ", " & _
         "Rebate = " & CLng(txtRabat.Text) & ", " & _
         "Demands = " & CSng(txtPotrazuje.Text) & ", " & _
         "Article = " & cboArtikal.ItemData(cboArtikal.ListIndex) & ", " & _
         "Client = " & cboKlijent.ItemData(cboKlijent.ListIndex) & ", " & _
         "[Payment Date] = '" & dtNaplata.Value & "', " & _
         "[Shipment Date] = '" & dtOtprema.Value & "' WHERE " & DataGrid.KeyField & _
         " = " & .Fields("ID_tblOrders")
      ElseIf crMode = 2 Then
         sql = "UPDATE tblClients SET " & _
         "Name = '" & txtImeKlijenta.Text & "', " & _
         "Address = '" & txtAdresa.Text & "', " & _
         "Employee = " & cboKomercijalista.ItemData(cboKomercijalista.ListIndex) & ", " & _
         "Rebate = " & txtRabat2.Text & ", " & _
         "Phone1 = '" & txtTel(0).Text & "', " & _
         "Phone2 = '" & txtTel(1).Text & "', " & _
         "Phone3 = '" & txtTel(2).Text & "' WHERE " & DataGrid.KeyField & _
         " = " & .Fields("ID_tblClients")
      ElseIf crMode = 3 Then
         sql = "UPDATE tblArticles SET " & _
         "Article = '" & txtImeArtikla.Text & "', " & _
         "Price = " & CSng(txtCenaArtikla.Text) & ", " & _
         "[Price TP] = " & CSng(txtCenaTP.Text) & ", " & _
         "[TP] = " & txtTPArtikla.Text & ", " & _
         "Comment = '" & txtNapomena.Text & _
         "' WHERE " & DataGrid.KeyField & _
         " = " & .Fields("ID_tblArticles")
         
      End If
      
      .ActiveConnection.Execute sql
   End With
   
   DataGrid.DoNotRefreshOnMove = True
   DataGrid.DataSource.Requery
   DataGrid.DataSource.Move pos - 1, 1
   
'  manual Update
   If Not AddingNew Then
      With DataGrid
         Dim i&
         For i = 1 To .Columns.Count
            .UpdateData Adodc.Recordset.AbsolutePosition - 1, _
                        .Columns(i).Index - 1, , False
         Next i
         .Refresh
      End With
   End If
   
   DataGrid.DoNotRefreshOnMove = False
   
   AddingNew = False
End Sub


Private Sub cmdDodaj_Click()
   Dim sSql$
   
   If crMode = 1 Then
      sSql = _
      "INSERT INTO tblOrders (Article, Client, CP, TP, Payed, [For Payment], Demands, Owing, [Payment Date], [Shipment Date]) VALUES "
      sSql = sSql _
      & "( " & cboArtikal.ItemData(cboArtikal.ListIndex) _
      & ", " & cboKlijent.ItemData(cboKlijent.ListIndex) _
      & ", " & CLng(txtKP.Text) _
      & ", " & CLng(txtTP.Text) _
      & ", " & CSng(txtNaplaceno.Text) _
      & ", " & Knjiga("Articles").RecordSource.Fields("Price") * CLng(txtKP.Text) _
      & ", " & CSng(txtPotrazuje.Text) _
      & ", " & CSng(txtDuguje.Text) _
      & ", '" & dtNaplata.Value & "'" _
      & ", '" & dtOtprema.Value & "'" _
      & ") "
   ElseIf crMode = 2 Then
      If txtImeKlijenta.Text = "" Then Exit Sub
      sSql = _
      "INSERT INTO tblClients (Employee, Address, Rebate, Name, Phone1, Phone2, Phone3) VALUES "
      sSql = sSql _
      & "( " & cboKomercijalista.ItemData(cboKomercijalista.ListIndex) _
      & ", '" & txtAdresa.Text & "'" _
      & ", " & CLng(txtRabat2.Text) _
      & ", '" & txtImeKlijenta.Text & "'" _
      & ", '" & txtTel(0).Text & "'" _
      & ", '" & txtTel(1).Text & "'" _
      & ", '" & txtTel(2).Text & "'" _
      & ") "
   ElseIf crMode = 3 Then
      If txtImeArtikla.Text = "" Then Exit Sub
      sSql = _
      "INSERT INTO tblArticles (Article, Price, [Price TP], [TP], Comment) VALUES "
      sSql = sSql _
      & "( '" & txtImeArtikla.Text & "'" _
      & ", " & CSng(txtCenaArtikla.Text) _
      & ", " & CSng(txtCenaTP.Text) _
      & ", " & CSng(txtTPArtikla.Text) _
      & ", '" & txtNapomena.Text & "'" _
      & ") "
   End If
   
   Debug.Print sSql
   
   adoActiveAdo.Recordset.ActiveConnection.Execute sSql
   
   DataGrid.DataSource.Requery
   
   DataGrid.RefreshData
   
   DataGrid.DataSource.MoveLast
   
End Sub

Private Sub cmdOcisti_Click()
   cboArtikal.Text = ""
   cboKlijent.Text = ""
   cboKomercijalista.Text = ""
   
   txtTP.Text = "0"
   txtKP.Text = "0"
   txtNaplaceno.Text = "0.00"
   txtDuguje.Text = "0.00"
   txtPotrazuje.Text = "0.00"
   
   txtRabat2.Text = "0"
   txtAdresa.Text = ""
   txtImeKlijenta.Text = ""
   txtTel(0).Text = ""
   txtTel(1).Text = ""
   txtTel(2).Text = ""

   txtImeArtikla = ""
   txtNapomena.Text = ""
   txtTPArtikla.Text = "0"
   txtCenaTP.Text = "0"
   txtCenaArtikla.Text = "0"
   
   dtNaplata.Value = Date
   dtOtprema.Value = Date
End Sub

Private Sub Command3_Click()
   txtNaplaceno.Text = Knjiga("Articles").RecordSource.Fields("Price") * CLng(txtKP.Text)
   txtNaplaceno_LostFocus
   
   Command1_Click
End Sub



Private Sub Command4_Click()
   txtCenaTP.Text = CSng(txtCenaArtikla.Text) * CSng(txtTPArtikla.Text)
   txtCenaTP_LostFocus
End Sub

Public Sub Form_Load()
   frmMain.mnuPregled.Enabled = True
   
   Form_Resize
      
   Set adoActiveAdo = Adodc
   Set adoActiveGrid = DataGrid
   
   Adodc.ConnectionString = strConnectionString
   
   RunSQL OpenSqlFile(SQLFile)
   
   frmSql.cboFiles.Text = SQLFile

   ' Namesti SQL prozor da prikazuje greske kako treba (ne dva puta)...
   ' Formular treba da ISKLJUCI polja koja ne postoje u SQL pogledu, a
   ' jos zajebanije ce biti uopste generisanje i upravljanje tim
   ' formularima jer postoje razni slucajevi,
   ' Iskljucuj izmene ako pogled ima agregate!
   ' Dodaj drop-down grid na lookup polja...
   ' Mozda je najbolje formular praviti posebno... verovatno :P, formule
   ' i sve to...
   
   If DataGrid.IsDatasetEmpty Then
      'Me.Enabled = False
      EnableEdit False
      Exit Sub
   Else
      EnableEdit True
   End If
   
   LoadLookupData
   
   adoActiveGrid.DataSource.MoveFirst
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
   Adodc.Width = Me.ScaleWidth
   Adodc.Left = 0
   
   With DataGrid
      .Move 0, Adodc.Top + Adodc.Height, Me.ScaleWidth, Me.ScaleHeight - .Top
   End With
   
   With picCommands
      .Move Me.ScaleWidth - .Width, 4
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   frmMain.mnuPregled.Enabled = False
   
   Select Case crMode
   Case 2
      Knjiga("Clients").RecordSource.Requery
   Case 3
      Knjiga("Articles").RecordSource.Requery
   End Select
   
   Adodc.Recordset.Close
End Sub

Private Sub txtCenaArtikla_LostFocus()
On Error GoTo e
   With txtCenaArtikla
      .Text = FormatNumber(.Text, 2)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0.00"
      Exit Sub
   End With
End Sub

Private Sub txtCenaTP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Command4_Click
End Sub

Private Sub txtCenaTP_LostFocus()
On Error GoTo e
   With txtCenaTP
      .Text = FormatNumber(.Text, 2)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0.00"
      Exit Sub
   End With
End Sub

Private Sub txtDuguje_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub txtDuguje_LostFocus()
On Error GoTo e
   With txtDuguje
      .Text = FormatNumber(.Text, 2)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0.00"
      Exit Sub
   End With
End Sub

Private Sub txtKP_LostFocus()
On Error GoTo e
   With txtKP
      .Text = FormatNumber(.Text, 0)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0"
      Exit Sub
   End With
End Sub

Private Sub txtNaplaceno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Command3_Click
End Sub

Private Sub txtNaplaceno_LostFocus()
On Error GoTo e
   With txtNaplaceno
      .Text = FormatNumber(.Text, 2)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0.00"
      Exit Sub
   End With
End Sub

Private Sub txtPotrazuje_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub txtPotrazuje_LostFocus()
On Error GoTo e
   With txtPotrazuje
      .Text = FormatNumber(.Text, 2)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0.00"
      Exit Sub
   End With
End Sub

Private Sub txtRabat_LostFocus()
On Error GoTo e
   With txtRabat
      .Text = Replace(.Text, "%", "", , , vbTextCompare)
      .Text = FormatNumber(.Text, 0)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0"
      Exit Sub
   End With
End Sub

Private Sub txtTP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Command2_Click
End Sub

Private Sub txtTP_LostFocus()
On Error GoTo e
   With txtTP
      .Text = FormatNumber(.Text, 0)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0"
      Exit Sub
   End With
End Sub

Public Sub LoadLookupData()
   ComboData cboArtikal, Knjiga("Articles").RecordSource
   ComboData cboKlijent, Knjiga("Clients").RecordSource
   ComboData cboKomercijalista, Knjiga("Employees").RecordSource
End Sub

Public Sub PregledProdaje()
   
End Sub

Public Sub EnableEdit(Enable As Boolean)
   cmdDodaj.Enabled = Enable
   cmdOcisti.Enabled = Enable
   cmdIzmeni.Enabled = Enable
   Command1.Enabled = Enable
   Command2.Enabled = Enable
   Command3.Enabled = Enable
   cboKlijent.Enabled = Enable
   cboArtikal.Enabled = Enable
End Sub

Public Sub SetMode(md As String)
   Select Case md
   Case "Orders"
      picRoba.Visible = True
      picKlijenti.Visible = False
      picArtikli.Visible = False
      crMode = 1
      Me.Caption = "Roba"
      DataGrid.KeyField = "tblOrders.ID"
   Case "Clients"
      picKlijenti.Visible = True
      picRoba.Visible = False
      picArtikli.Visible = False
      crMode = 2
      Me.Caption = "Clients"
      DataGrid.KeyField = "tblClients.ID"
   Case "Articles"
      picArtikli.Visible = True
      picRoba.Visible = False
      picKlijenti.Visible = False
      crMode = 3
      Me.Caption = "Articles"
      DataGrid.KeyField = "tblArticles.ID"
   End Select

   
End Sub


Private Sub txtRabat2_LostFocus()
On Error GoTo e
   With txtRabat2
      .Text = Replace(.Text, "%", "", , , vbTextCompare)
      .Text = FormatNumber(.Text, 0)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0"
      Exit Sub
   End With
End Sub

Private Sub txtTPArtikla_LostFocus()
On Error GoTo e
   With txtTPArtikla
      .Text = FormatNumber(.Text, 0)
Exit Sub
e:
      MsgBox "Entered value is not valid!", vbCritical, "Greska"
       
      .Text = "0"
      Exit Sub
   End With
End Sub
