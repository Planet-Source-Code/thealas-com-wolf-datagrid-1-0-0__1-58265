VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Small Corporation"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   9960
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0442
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuBaza 
      Caption         =   "&Program"
      Begin VB.Menu mnuBazaSadrzaj 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuBazaSql 
         Caption         =   "&SQL Window"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBazaS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBazaKraj 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuProzori 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mnuProzoriKaskada 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuProzoriTileVertikalno 
         Caption         =   "&Tile Vertically"
      End
      Begin VB.Menu mnuProzoriTileHorizontalno 
         Caption         =   "&Tile Horisontally"
      End
      Begin VB.Menu mnuProzoriOrganizuj 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuPregled 
      Caption         =   "&View"
      Begin VB.Menu mnuPregledIzmena 
         Caption         =   "&Action query"
         Begin VB.Menu mnuPregledIzmenaSve 
            Caption         =   "&All"
         End
         Begin VB.Menu mnuPregledIzmenaMesec 
            Caption         =   "&Filter by month"
         End
      End
      Begin VB.Menu mnuPregledProdaje 
         Caption         =   "&Sales grouped by articles"
         Begin VB.Menu mnuPregledProdajeSve 
            Caption         =   "&All"
         End
         Begin VB.Menu mnuPregledProdajeMesec 
            Caption         =   "&Filter by month"
         End
      End
      Begin VB.Menu mnuPregledZarade 
         Caption         =   "&Salary grouped by clients"
         Begin VB.Menu mnuPregledZaradeSve 
            Caption         =   "&All"
         End
         Begin VB.Menu mnuPregledZaradeMesec 
            Caption         =   "&Filter by month"
         End
      End
      Begin VB.Menu mnuPregledKomercijalista 
         Caption         =   "&Salary grouped by employees"
      End
      Begin VB.Menu mnuPregledS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPregledOtpremnica 
         Caption         =   "&Daily orders report"
      End
      Begin VB.Menu mnuPregledS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPregledText 
         Caption         =   "&Tekstualni mod"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPregledSortiraj 
         Caption         =   "&Sortiraj redove"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
   mnuPregled.Enabled = False
   
   
   strDataBaseFile = App.Path & "\SmallCorporation.mdb"
   strConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & strDataBaseFile & ";DefaultDir=" & ParseDir(strDataBaseFile) & ";"
   'strConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & strDataBaseFile & ";Persist Security Info=False"
    
   adoConnection.Open strConnectionString
   
   Knjiga.Add "Articles", New Recordset
   Knjiga.Add "Clients", New Recordset
   Knjiga.Add "Employees", New Recordset
   
   Knjiga("Articles").RecordSource.Open "tblArticles", adoConnection, adOpenKeyset, adLockOptimistic, adCmdTable
   Knjiga("Clients").RecordSource.Open "tblClients", adoConnection, adOpenKeyset, adLockOptimistic, adCmdTable
   Knjiga("Employees").RecordSource.Open "tblEmployees", adoConnection, adOpenKeyset, adLockOptimistic, adCmdTable
   
   Knjiga("Articles").RefreshRawData
   Knjiga("Clients").RefreshRawData
   Knjiga("Employees").RefreshRawData

   mnuPregledText.Checked = GetSetting(App.Title, "opcije", "textmode", False)

   frmContents.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Knjiga("Articles").RecordSource.Close
   Knjiga("Clients").RecordSource.Close
End Sub

Private Sub mnuBazaKraj_Click()
   Unload Me
End Sub

Private Sub mnuBazaSadrzaj_Click()
   frmContents.Show
End Sub

Private Sub mnuBazaSql_Click()
   frmSql.Show
End Sub


Private Sub mnuPregledIzmenaMesec_Click()
   With frmDetail
      frmDate.Show vbModal

      If frmDate.dt1 <> "" And frmDate.dt2 <> "" Then
         Dim sql$
         .Show
         sql = OpenSqlFile("roba_def_mesec")
         With frmDate
            sql = Replace(sql, "%KOLONA%", .sCol, , , vbTextCompare)
            sql = Replace(sql, "%DATUM1%", .sql_dt1, , , vbTextCompare)
            sql = Replace(sql, "%DATUM2%", .sql_dt2, , , vbTextCompare)
         End With
         RunSQL sql
         frmDetail.EnableEdit False
         mnuPregledOtpremnica.Enabled = False

      End If
   End With
End Sub

Private Sub mnuPregledIzmenaSve_Click()
   frmDetail.SQLFile = "roba_def"
   frmDetail.Form_Load
   frmDetail.EnableEdit True
   mnuPregledOtpremnica.Enabled = True
End Sub

Private Sub mnuPregledKomercijalista_Click()
   With frmDetail
      frmDate.Show vbModal

      If frmDate.dt1 <> "" And frmDate.dt2 <> "" Then
         Dim sql$
         .Show
         sql = OpenSqlFile("roba_pregled_zarade_prema_Employeesma")
         With frmDate
            sql = Replace(sql, "%KOLONA%", .sCol, , , vbTextCompare)
            sql = Replace(sql, "%DATUM1%", .sql_dt1, , , vbTextCompare)
            sql = Replace(sql, "%DATUM2%", .sql_dt2, , , vbTextCompare)
         End With
         RunSQL sql
         frmDetail.EnableEdit False
         mnuPregledOtpremnica.Enabled = False
      End If
   End With
End Sub

Private Sub mnuPregledOtpremnica_Click()
   With frmDetail
      If Not .crMode = 1 Then MsgBox "Orders action query must be opened!", vbExclamation: Exit Sub
      
      Dim sf$, dt As Date: dt = .Adodc.Recordset.Fields("Datum Otpreme").Value
      sf = "{tblRoba." & "Datum Otpreme" & "} = Date (" & Year(dt) & ", " & Month(dt) & ", " & Day(dt) & _
      ")" & _
      " and {tblClients.Ime} = '" & Knjiga("Clients").RecordSource.Fields("Ime").Value & _
      "'"
      Debug.Print sf
      .crOtpremnica.WindowTitle = "Otpremnica za " & dt
      .crOtpremnica.ReplaceSelectionFormula sf
      .crOtpremnica.PrintReport
   End With
End Sub

   
Private Sub mnuPregledProdajeMesec_Click()
   With frmDetail
      frmDate.Show vbModal

      If frmDate.dt1 <> "" And frmDate.dt2 <> "" Then
         Dim sql$
         .Show
         sql = OpenSqlFile("roba_pregled_prema_prodaji_artikla_mesec")
         With frmDate
            sql = Replace(sql, "%KOLONA%", .sCol, , , vbTextCompare)
            sql = Replace(sql, "%DATUM1%", .sql_dt1, , , vbTextCompare)
            sql = Replace(sql, "%DATUM2%", .sql_dt2, , , vbTextCompare)
         End With
         RunSQL sql
         frmDetail.EnableEdit False
         mnuPregledOtpremnica.Enabled = False
      End If
   End With
End Sub

Private Sub mnuPregledProdajeSve_Click()
   frmDetail.SQLFile = "roba_pregled_prema_prodaji_artikla"
   frmDetail.Form_Load
   frmDetail.EnableEdit False
   mnuPregledOtpremnica.Enabled = False
End Sub

Private Sub mnuPregledSortiraj_Click()
   With frmDetail
   End With
End Sub


Private Sub mnuPregledZaradeMesec_Click()
   With frmDetail
      frmDate.Show vbModal

      If frmDate.dt1 <> "" And frmDate.dt2 <> "" Then
         Dim sql$
         .Show
         sql = OpenSqlFile("roba_pregled_zarade_prema_Clientsma_mesec")
         With frmDate
            sql = Replace(sql, "%KOLONA%", .sCol, , , vbTextCompare)
            sql = Replace(sql, "%DATUM1%", .sql_dt1, , , vbTextCompare)
            sql = Replace(sql, "%DATUM2%", .sql_dt2, , , vbTextCompare)
         End With
         frmDetail.EnableEdit False
         mnuPregledOtpremnica.Enabled = False
         RunSQL sql
      End If
   End With
End Sub

Private Sub mnuPregledZaradeSve_Click()
   frmDetail.SQLFile = "roba_pregled_zarade_prema_Clientsma"
   frmDetail.Form_Load
   frmDetail.EnableEdit False
   mnuPregledOtpremnica.Enabled = False
End Sub

Private Sub mnuProzoriKaskada_Click()
   Me.Arrange vbCascade
End Sub

Private Sub mnuProzoriOrganizuj_Click()
   Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuProzoriTileHorizontalno_Click()
   Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuProzoriTileVertikalno_Click()
   Me.Arrange vbTileVertical
End Sub
