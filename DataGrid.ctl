VERSION 5.00
Begin VB.UserControl DataGrid 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DataBindingBehavior=   2  'vbComplexBound
   DrawWidth       =   56
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "DataGrid.ctx":0000
   Begin VB.HScrollBar hscrTable 
      Height          =   195
      Left            =   2790
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   3375
      Width           =   1815
   End
   Begin VB.VScrollBar vscrTable 
      Height          =   2265
      Left            =   4590
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   1080
      Width           =   195
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      HideSelection   =   0   'False
      Left            =   945
      TabIndex        =   2
      Top             =   810
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblDesign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wolf Grid"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   135
      TabIndex        =   3
      Top             =   135
      Width           =   1410
   End
End
Attribute VB_Name = "DataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ***************************************************************************************
' Wolf DataBase System
' Copyright (C) TheAlas Software 2004-2005
' http://www.thealas.com
' You are NOT free to use this code as you like, if you find useful parts and would like
' to use them in your software please send me email for permition.
' You are again NOT free to use/compile this code whatsoever unless you have my pertmition.
' If you dont want to leave credit in your aboutbox (that is ususaly all I require for
' permition), just use compiled version of the control if it is provided with code.
' Thank you for downloading this code, I hope you will learn something as it is the only
' reason I posted it on PSC, if not then I hope you will find some use of this control.
' If you liked this control then go to the PSC site and vote, it may not only show that
' this control is actualy useful, but it may even bring some prize to me :).
' If you do not like this control, then leave it and do not vote unless you think it is
' abusive.
' ***************************************************************************************
'
' Personal opinion about Microsoft ADO:
' !(warning: contains abusive terms)!
' Before u start changing anything u should know one thinng...
' ... **M$ ADO IS $HIT**, thats prety much all u MUST know.
' (well, vb, windows and all M$ stuff is shit as well,
' but the above mentioned is really the ultimate crap)
' I guess they'r just trying to make shity software for this shity world...
' Use SQL instead of ado/dao/oledb... and I dunno what other craps are there.
' SQL is created to bring money, and as we know M$ likes money so it likes SQL
' too, we get simple formula: money=good => sql=money => sql=good, simple huh?
' But we also know that M$ never makes good stuff! So we get something like almost
' good which is still better than shit, that as we know brings no money, it just
' smells bad (havent tasted raw human shit yet, but I suppose it tastes bad).
' ***************************************************************************************

' NOTE:
' There are some parts of code commented with "//", "???", "!!!" or "***", this means
' that commented line is either useless, not working or it is soon to change.
' NOTE:
' As you can see this control depends on certain CommonControls library. This is so
' DatePick control and Spin controls can be used in later versions, this version have no
' editable fields yet.
Option Explicit

Private mvarRawData() As Variant
Private mvarColumnCount As Integer
Private mvarRowCount As Long
Private mvarRowHeight As Integer
Private mvarBorderStyle As Long
Private mvarIndicatorWidth As Integer
Private mvarData As Variant
Private WithEvents mvarDataSource As ADODB.Recordset
Attribute mvarDataSource.VB_VarHelpID = -1
Private mvarSelectedRow As Long  ' //
Private mvarSelectedCol As Long ' ????
Private mvarSelectedRecord As Long
Private mvarCurrentRow As Long
Private mvarCurrentCol As Long
Private mvarEditMode As Boolean
Private mvarNullString As String
Private mvarLockedEdit As Boolean
Private mvarAllowDelete As Boolean
Private mvarDisableEditBox As Boolean
Private mvarHideKeyColumns As Boolean
Private mvarKeyField As String
Private mvarDoNotRefreshOnMove As Boolean
Private mvarManualUpdate As Boolean
Private mvarOnlyView As Boolean
Private mvarEnabled As Boolean

Private WithEvents mvarGridFont As StdFont
Attribute mvarGridFont.VB_VarHelpID = -1

Private lxCol As Long
Private lxPos As Long     ' increments by row height (this is only for data, not actually graphic-related)
Private lyPos As Long     ' increments by 16px-64px, moves the drawing area

Private mvarGridColor As OLE_COLOR
Private mvarTableColor As OLE_COLOR
Private mvarTextColor As OLE_COLOR
Private mvarHeaderTextColor As OLE_COLOR
Private mvarBackColor As OLE_COLOR
Private mvarCurrentRecordColor As OLE_COLOR

Private mpenGridColor As Long
Private mpenWindowText As Long
Private mpenGridText As Long
Private mpenHeaderText As Long
Private mpenButtonFace As Long
Private mpen3DLight As Long

Private mbrushButtonFace As Long
Private mbrushTableColor As Long
Private mbrushWindowBackground As Long
Private mbrushAppWorkspace As Long
Private mbrushBackColor As Long

Private mhdcIndicator As Long

Private mvarDrawGridLines As Boolean
Private mvarDrawColumnHeaders As Boolean
Private mvarDrawIndicator As Boolean
Private mvarDontPaint As Boolean

Public Columns As New CColumns
Attribute Columns.VB_VarDescription = "Class of DataGrid Columns."
Public Selection As New CSelection
Attribute Selection.VB_VarDescription = "Class for selected records."
Public ColSelection As New CSelection
Attribute ColSelection.VB_VarDescription = "Selected columns class, key is used to hold column position as ""num:[Position]"" where [Position] is left to right position of the column."

Private dcBuff As Long
Private lBmp As Long

Private fntGrid As Long

Private bInitialized As Boolean

Private bMouseDown As Boolean

Private bSelecting As Boolean
Private lFirstSelected As Long
Private lFirstSelectedShift As Long
Private bColSelecting As Boolean
Private lColFirstSelected As Long
Private lColFirstSelectedShift As Long
Private bShiftKey As Boolean
Private bCtrlKey As Boolean
Private lCurrentColSized As Long
Private bColSizing As Boolean
Private lTopSelected As Long
Private bDeleteing As Boolean
Private bDisableMouseUp As Boolean

'
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()
Public Event DblClick()
Public Event KeyPress(KeyAscii As Integer)
Public Event Resize()

Public Enum spBorderStyle
   None = 0
   FixedSingle = 1
   Flat = 2
End Enum

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
   Enabled = mvarEnabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
   mvarEnabled = vNewValue
   UserControl.Enabled = mvarEnabled
End Property


Public Property Get OnlyView() As Boolean
Attribute OnlyView.VB_Description = "Returns/sets if DataGrid cannot interact with user, used for View queries."
   OnlyView = mvarOnlyView
End Property

Public Property Let OnlyView(ByVal vNewValue As Boolean)
   mvarOnlyView = vNewValue
   
   AllowDelete = Not mvarOnlyView
   DisableEditBox = mvarOnlyView
End Property


Public Property Get ManualUpdate() As Boolean
Attribute ManualUpdate.VB_Description = "Returns/sets if raw data update should be automatic (grid data changes as DataSource changes) or manual (user must change grid data manualy)."
   ManualUpdate = mvarManualUpdate
End Property

Public Property Let ManualUpdate(ByVal vNewValue As Boolean)
   mvarManualUpdate = vNewValue
End Property


Public Property Get DoNotRefreshOnMove() As Boolean
Attribute DoNotRefreshOnMove.VB_Description = "Returns/sets if DataGrid should refresh after every DataSource Move/MoveFirst/MoveLast/MoveNext command."
   DoNotRefreshOnMove = mvarDoNotRefreshOnMove
End Property

Public Property Let DoNotRefreshOnMove(ByVal vNewValue As Boolean)
   mvarDoNotRefreshOnMove = vNewValue
End Property


Public Property Get KeyField() As String
Attribute KeyField.VB_Description = "Returns/sets key field used for action queries such as delete/update. Note that advanced features require user to specify key field as ""ID_tblCustomers"", it means that key field is ""ID"" from table ""tblCustomers""."
   KeyField = mvarKeyField
End Property

Public Property Let KeyField(ByVal vNewValue As String)
   mvarKeyField = vNewValue
End Property


Public Property Get HideKeyColumns() As Boolean
Attribute HideKeyColumns.VB_Description = "Returns/sets if DataGrid should hide columns that are associated with key fields. Note that grid cannot detect key fields, it will hide those with ""ID_"" in the left part of the name. Use names such as ""ID_tblCustomers"" or ""ID_tblEmployees"""
   HideKeyColumns = mvarHideKeyColumns
End Property

Public Property Let HideKeyColumns(ByVal vNewValue As Boolean)
   mvarHideKeyColumns = vNewValue
   
   If mvarHideKeyColumns Then
      Dim i&
      For i = 1 To Columns.Count
         If Left(Columns(i).Caption, 2) = "ID" Then
            Columns(i).Visible = False
         End If
      Next i
      If Not (mvarDataSource Is Nothing) Then _
         Refresh
   End If
   
   PropertyChanged "HideKeyColumns"
End Property


Public Property Get CurrentRecordColor() As OLE_COLOR
Attribute CurrentRecordColor.VB_Description = "Not used yet."
   CurrentRecordColor = mvarCurrentRecordColor
End Property

Public Property Let CurrentRecordColor(ByVal vNewValue As OLE_COLOR)
   mvarCurrentRecordColor = vNewValue
End Property


Public Property Get RowCount() As Long
Attribute RowCount.VB_Description = "Returns total number of rows used in grid, 0-based."
Attribute RowCount.VB_MemberFlags = "400"
   RowCount = mvarRowCount
End Property

Public Property Get DisableEditBox() As Boolean
Attribute DisableEditBox.VB_Description = "Returns/sets if DataGrid should disable automatic field editing feature."
   DisableEditBox = mvarDisableEditBox
End Property

Public Property Let DisableEditBox(ByVal vNewValue As Boolean)
   mvarDisableEditBox = vNewValue
End Property


Public Property Get RecordPosition() As Long
Attribute RecordPosition.VB_Description = "Returns current record position of the DataSet."
Attribute RecordPosition.VB_MemberFlags = "400"
   RecordPosition = mvarSelectedRecord
End Property

Public Property Let RecordPosition(ByVal vNewValue As Long)
   mvarSelectedRecord = vNewValue
On Error GoTo e
   mvarDataSource.Move mvarSelectedRecord
   Refresh
e:
End Property


Public Property Get AllowDelete() As Boolean
Attribute AllowDelete.VB_Description = "Returns/sets if it is allowed records to be deleted from DataGrid by pressing Del key."
   AllowDelete = mvarAllowDelete
End Property

Public Property Let AllowDelete(ByVal vNewValue As Boolean)
   mvarAllowDelete = vNewValue
End Property


Public Property Get LockedEdit() As Boolean
Attribute LockedEdit.VB_Description = "Returns/sets if edit box is enabled, but no editing is allowed within."
   LockedEdit = mvarLockedEdit
End Property

Public Property Let LockedEdit(ByVal vNewValue As Boolean)
   mvarLockedEdit = vNewValue
   txtEdit.Locked = mvarLockedEdit
End Property


Public Property Get DontPaint() As Boolean
Attribute DontPaint.VB_Description = "Returns/sets if grid should be painted."
   DontPaint = mvarDontPaint
End Property

Public Property Let DontPaint(ByVal vNewValue As Boolean)
   mvarDontPaint = vNewValue
End Property


Public Property Get GridFont() As StdFont
Attribute GridFont.VB_Description = "Returns/sets font used for drawing text on the grid."
   Set GridFont = mvarGridFont
End Property

Public Property Set GridFont(ByVal vNewValue As StdFont)
   Set mvarGridFont = vNewValue
   
   If (mvarGridFont Is Nothing) Then _
      Set mvarGridFont = UserControl.Font
   
   Set txtEdit.Font = mvarGridFont
   Set UserControl.Font = mvarGridFont
   
   mvarGridFont_FontChanged ""
End Property




Private Sub mvarDataSource_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If Not mvarManualUpdate Then
      Dim i&
      For i = 1 To Columns.Count
         UpdateData pRecordset.AbsolutePosition - 1, _
                     Columns(i).Index - 1, , False
      Next i
      Refresh
   End If
End Sub

Private Sub mvarDataSource_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo e
   mvarSelectedRecord = mvarDataSource.AbsolutePosition - 1
   
   Dim rc&
   
   If mvarDoNotRefreshOnMove Then Exit Sub
   
   With UserControl
      If mvarSelectedRecord < 0 Then Exit Sub
      
      rc = CLng(.ScaleHeight / mvarRowHeight) + 1
      If mvarRowCount + lyPos + 4 > rc Then
         If mvarSelectedRecord + lyPos + 4 > rc Then
            vscrTable.Value = vscrTable.Value + 1
         End If
      End If
      
      If mvarSelectedRecord + lyPos + 1 = 0 Then
         vscrTable.Value = vscrTable.Value - 1
      End If
      
      Dim lFirst&, lLast&
      
      lFirst = Abs(lyPos)
      lLast = lFirst + rc - 1
      
      If mvarSelectedRecord < lFirst Then
         vscrTable.Value = mvarSelectedRecord
      End If
      If mvarSelectedRecord > lLast Then
         Dim lp&: lp = mvarSelectedRecord - rc + 4
         If lp > 0 And lp < mvarRowCount Then _
            vscrTable.Value = lp
      End If
   End With
   
   Refresh
Exit Sub
e:
   mvarSelectedRecord = 0
End Sub

Private Sub mvarDataSource_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
End Sub

Private Sub mvarDataSource_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
End Sub

Private Sub mvarDataSource_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
End Sub


Private Sub mvarGridFont_FontChanged(ByVal PropertyName As String)
   GDIObjDelete
   GDIObjCreate


   PropertyChanged "GridFont"
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of DataGrid."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
   BackColor = mvarBackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
   mvarBackColor = vNewValue
   GDIObjDelete
   GDIObjCreate
End Property


Public Property Get HeaderTextColor() As OLE_COLOR
Attribute HeaderTextColor.VB_Description = "Returns/sets column headers text color."
Attribute HeaderTextColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   HeaderTextColor = mvarHeaderTextColor
End Property

Public Property Let HeaderTextColor(ByVal vNewValue As OLE_COLOR)
   mvarHeaderTextColor = vNewValue
   GDIObjDelete
   GDIObjCreate
End Property


Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Returns/sets color of the text used for drawing data."
Attribute TextColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   TextColor = mvarTextColor
End Property

Public Property Let TextColor(ByVal vNewValue As OLE_COLOR)
   mvarTextColor = vNewValue
   GDIObjDelete
   GDIObjCreate
   txtEdit.ForeColor = vNewValue
End Property

Public Property Get RowHeight() As Integer
Attribute RowHeight.VB_Description = "Returns/sets height of DataGrid rows and column headers."
Attribute RowHeight.VB_ProcData.VB_Invoke_Property = ";Position"
   RowHeight = mvarRowHeight
End Property

Public Property Let RowHeight(ByVal vNewValue As Integer)
   mvarRowHeight = vNewValue
End Property


Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Returns/sets color of grid lines."
Attribute GridColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   GridColor = mvarGridColor
End Function

Public Property Let GridColor(ByVal vNewValue As OLE_COLOR)
   mvarGridColor = vNewValue
   GDIObjDelete
   GDIObjCreate
   PropertyChanged "GridColor"
End Property

Public Property Get DrawGridLines() As Boolean
Attribute DrawGridLines.VB_Description = "Returns/sets if the DataGrid is drawn with grid lines."
Attribute DrawGridLines.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DrawGridLines = mvarDrawGridLines
End Property

Public Property Let DrawGridLines(ByVal vNewValue As Boolean)
   mvarDrawGridLines = vNewValue
End Property

Public Property Get DrawColumnHeaders() As Boolean
Attribute DrawColumnHeaders.VB_Description = "Returns/sets if the DataGrid will draw column headers, only text if False."
Attribute DrawColumnHeaders.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DrawColumnHeaders = mvarDrawColumnHeaders
End Property

Public Property Let DrawColumnHeaders(ByVal vNewValue As Boolean)
   mvarDrawColumnHeaders = vNewValue
End Property

Public Property Get DrawIndicator() As Boolean
Attribute DrawIndicator.VB_Description = "Returns/sets if the DataGrid is drawn with the indicator bar."
Attribute DrawIndicator.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DrawIndicator = mvarDrawIndicator
End Property

Public Property Let DrawIndicator(ByVal vNewValue As Boolean)
   mvarDrawIndicator = vNewValue
End Property

Public Property Get TableColor() As OLE_COLOR
Attribute TableColor.VB_Description = "Returns/sets background color of the DataGrid table."
Attribute TableColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   TableColor = mvarTableColor
End Property

Public Property Let TableColor(ByVal vNewValue As OLE_COLOR)
   mvarTableColor = vNewValue
   txtEdit.BackColor = vNewValue
   GDIObjDelete
   GDIObjCreate
   UserControl.BackColor = vNewValue
End Property

Public Property Get NullString() As String
Attribute NullString.VB_Description = "Returns/sets string used for empty (NULL) field."
Attribute NullString.VB_ProcData.VB_Invoke_Property = ";Data"
   NullString = mvarNullString
End Property

Public Property Let NullString(ByVal vNewValue As String)
   mvarNullString = vNewValue
End Property

Public Property Get RawData(Row As Long, Col As Long) As Variant
Attribute RawData.VB_Description = "Returns/sets data that DataGrid displays, used mostly for manual update and fast retrieval of data."
Attribute RawData.VB_MemberFlags = "400"
   RawData = mvarRawData(Row, Col)
End Property

Public Property Let RawData(Row As Long, Col As Long, ByVal vNewValue As Variant)
   mvarRawData(Row, Col) = vNewValue
End Property


Public Property Get BorderStyle() As spBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets if the DataGrid is drawn with a Border."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
   BorderStyle = mvarBorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewValue As spBorderStyle)
   mvarBorderStyle = vNewValue
   
   If mvarBorderStyle < 2 Then
      UserControl.BorderStyle = mvarBorderStyle
   Else
      UserControl.BorderStyle = 0
   End If
   
   PropertyChanged "BorderStyle"
End Property

Public Property Get IndicatorWidth() As Integer
Attribute IndicatorWidth.VB_Description = "Returns/sets width of the indicator bar"
Attribute IndicatorWidth.VB_ProcData.VB_Invoke_Property = ";Position"
   IndicatorWidth = mvarIndicatorWidth
End Property

Public Property Let IndicatorWidth(ByVal vNewValue As Integer)
   mvarIndicatorWidth = vNewValue
   
   PropertyChanged "IndicatorWidth"
End Property

Public Property Get DataSource() As ADODB.Recordset
Attribute DataSource.VB_Description = "Specifies source of grid data, ADO Recordset supported only."
Attribute DataSource.VB_MemberFlags = "200"
   Set DataSource = mvarDataSource
End Property

Public Property Set DataSource(ByVal vNewValue As ADODB.Recordset)
   Set mvarDataSource = vNewValue
   
   If (mvarDataSource Is Nothing) Then GoTo shit:
   
   UserControl.Enabled = True
   
   LoadData
   
   mvarDataSource.MoveFirst
   
shit:
   PropertyChanged "DataSource"
End Property

Private Sub LoadData()
On Error GoTo e
   '  just how much vb code can be stupid huh? I mean just look at this code that checks if something is not nothing:
   If Not (mvarDataSource Is Nothing) Then
      UserControl.Enabled = True
      
      ShowEditBox False
      
      ClearSelection
      ClearColSelection
      
      'well, lets load all now :D
      Dim i&, c&, a&
      
      If Not IsEmpty(mvarData) Then Erase mvarData
      
      mvarData = mvarDataSource.GetRows
      mvarRowCount = UBound(mvarData, 2) + 1
      
      mvarColumnCount = mvarDataSource.Fields.Count
      
      If mvarRowCount = 0 Then
         UserControl.Enabled = False
         Exit Sub
      Else
         UserControl.Enabled = True
      End If
      
      
      Dim sKey$
      'clean
del:
      For i = 1 To Columns.Count
         Columns.Remove i
         GoTo del
      Next i
      
      ' add, but check duplicate keys for eather grid and t-sql cannot make difference
      For i = 0 To mvarDataSource.Fields.Count - 1
         a = 0
         
         sKey = mvarDataSource.Fields(i).Name
         
check:
         Dim found%: found = 0
         For c = 1 To Columns.Count
            If Columns(c).Key = sKey Then _
               a = a + 1: found = 1
         Next c
                  
         If a > 0 Then _
            sKey = mvarDataSource.Fields(i).Name & CStr(a): If found Then GoTo check
         
         Columns.Add , sKey
      Next i
      
      Dim arr() As Long
    
      ReDim arr(mvarColumnCount - 1, mvarRowCount - 1)
      
      For c = 0 To mvarColumnCount - 1
         With Columns(c + 1)
            .Caption = mvarDataSource.Fields(c).Name
            .DataType = mvarDataSource.Fields(c).Type
         End With
         
         For i = 0 To mvarRowCount - 1
            If Not IsNull(mvarData(c, i)) Then
               Dim sval$: sval = CStr(mvarData(c, i))
               sval = Columns(c + 1).GetFormattedValue(sval)
               arr(c, i) = CLng(TextWidth(CStr(sval))) + TextWidth("A")
            End If
         Next i
      Next c
         
      ' sort the lenghts
      Dim j&, temp&, k&, cmpval&, ind&
      
      For i = 0 To mvarColumnCount - 1
         Dim longest&
         
         ' sort now
         arr = ShellSort(arr, i)
         
         longest = arr(i, mvarRowCount - 1)
         
         With Columns(i + 1)
            If CLng(TextWidth(mvarDataSource.Fields(i).Name)) > longest Then
               .ColumnWidth = TextWidth(mvarDataSource.Fields(i).Name) + 8
            Else
               .ColumnWidth = longest + 4
            End If
            ' what if column is bigger than control :S... many craps may occur
            If .ColumnWidth > UserControl.ScaleWidth Then _
               .ColumnWidth = UserControl.ScaleWidth - UserControl.ScaleWidth / 4
         End With
         
      Next i
   End If
   
   If bDeleteing Then
      'If lTopSelected * mvarRowHeight + mvarRowHeight > UserControl.ScaleHeight Then _
     '    lyPos = -lTopSelected + 1
   Else
      lxCol = 0
      lxPos = 0
      lyPos = 0
      vscrTable.Value = 0
      hscrTable.Value = 0
   End If
   
   bDeleteing = False
   
   If mvarHideKeyColumns Then
      For i = 1 To Columns.Count
         If Left(Columns(i).Caption, 2) = "ID" Then
            Columns(i).Visible = False
         End If
      Next i
   End If
   
   UserControl.Enabled = True
   
   UserControl_Resize
Exit Sub
e:
   mvarRowCount = 0
   mvarColumnCount = 0
   If Err = 3021 Then
      MsgBox lr(103), , lr(102)
   End If
   UserControl.Enabled = False
   Exit Sub
End Sub


Private Sub txtEdit_Click()
   RaiseEvent Click
End Sub

Private Sub txtEdit_LostFocus()
   ShowEditBox False
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
   Refresh
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub



Private Sub UserControl_InitProperties()
   BackColor = vbApplicationWorkspace
   HeaderTextColor = vbWindowText
   TextColor = vbWindowText
   BorderStyle = 1
   Set DataSource = Nothing
   GridColor = vbButtonFace
   TableColor = vbWindowBackground
   DrawGridLines = True
   DrawColumnHeaders = True
   HideKeyColumns = False
   DrawIndicator = True
   NullString = "Null"
   RowHeight = 18
   IndicatorWidth = 18
   LockedEdit = True
   AllowDelete = False
   Set Columns = Nothing
   mvarEnabled = True
   ManualUpdate = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyShift Then bShiftKey = True
   If KeyCode = vbKeyControl Then bCtrlKey = True
   
   If mvarAllowDelete Then
      If KeyCode = vbKeyDelete Then
         ' delete warning
         If MsgBox(lr(104), vbOKCancel Or vbQuestion Or vbDefaultButton2, lr(102)) = vbCancel Then
            UserControl.SetFocus ' little bug fixed, usercontrol not recieveing keys cancel
            Exit Sub
         End If
         
         If Not (Selection Is Nothing) Then
            If Selection.Count > 0 Then
               ' delete selected records
               Dim i&, Key, sql$, num&
               
               Key = Split(mvarKeyField, ".", , vbTextCompare)
               
               ' NEVER EVER use ado classes instead of sql, ado is EVIL!!
               ' It will simply not work, you may cause complete loss of data
               ' by using ado instead of action queries. Ado is free, and everything that
               ' is microsoft's and free is EVIL, it will destroy your mind (and ur data as well)
               sql = "DELETE * FROM " & Key(0) & " WHERE "
               
               mvarDoNotRefreshOnMove = True
               For i = 1 To Selection.Count
                  num = CLng(Mid(Selection(i).Key, 5))
                  mvarDataSource.Move num, 1
                  num = mvarDataSource.Fields("ID_" & Key(0))  ' table name is ALWAYS alias ID in queries
                  If i > 1 Then sql = sql & " OR "
                  sql = sql & Key(1) & "=" & num
               Next i
               mvarDoNotRefreshOnMove = False
               
               mvarDataSource.ActiveConnection.Execute sql
                              
               bDeleteing = True
               
               mvarDataSource.Requery
               
               LoadData
            Else
               MsgBox lr(101), vbExclamation, lr(102)
            End If
         End If
      End If
   End If
   
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyShift Then bShiftKey = False
   If KeyCode = vbKeyControl Then bCtrlKey = False
   
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      Enabled = .ReadProperty("Enabled", True)
      OnlyView = .ReadProperty("OnlyView", False)
      ManualUpdate = .ReadProperty("ManualUpdate", False)
      DoNotRefreshOnMove = .ReadProperty("DoNotRefreshOnMove", False)
      KeyField = .ReadProperty("KeyField", "")
      HideKeyColumns = .ReadProperty("HideKeyColumns", False)
      CurrentRecordColor = .ReadProperty("CurrentRecordColor", vbWindowBackground)
      DisableEditBox = .ReadProperty("DisableEditBox", False)
      RecordPosition = .ReadProperty("RecordPosition", 0)
      AllowDelete = .ReadProperty("AllowDelete", False)
      LockedEdit = .ReadProperty("LockedEdit", True)
      DontPaint = .ReadProperty("DontPaint", False)
      Set GridFont = .ReadProperty("GridFont", Nothing)
      BackColor = .ReadProperty("BackColor", vbApplicationWorkspace)
      HeaderTextColor = .ReadProperty("HeaderTextColor", vbWindowText)
      TextColor = .ReadProperty("TextColor", vbWindowText)
      BorderStyle = .ReadProperty("BorderStyle", 1)
      Set DataSource = .ReadProperty("DataSource", Nothing)
      GridColor = .ReadProperty("GridColor", vbButtonFace)
      TableColor = .ReadProperty("TableColor", vbWindowBackground)
      DrawGridLines = .ReadProperty("DrawGridLines", True)
      DrawColumnHeaders = .ReadProperty("DrawColumnHeaders", True)
      DrawIndicator = .ReadProperty("DrawIndicator", True)
      NullString = .ReadProperty("NullString", "Null")
      RowHeight = .ReadProperty("RowHeight", 18)
      IndicatorWidth = .ReadProperty("IndicatorWidth", 16)
      
      GDIObjDelete
      GDIObjCreate
   End With
End Sub

Private Sub UserControl_Show()
   If UserControl.Ambient.UserMode Then
      lblDesign.Visible = False
   Else
      'UserControl.Enabled = False
      lblDesign.Visible = True
   End If
   
   If mvarRowCount <= 0 Then _
      UserControl.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Enabled", Enabled, True
      .WriteProperty "OnlyView", OnlyView, False
      .WriteProperty "ManualUpdate", ManualUpdate, False
      .WriteProperty "DoNotRefreshOnMove", DoNotRefreshOnMove, False
      .WriteProperty "KeyField", KeyField, ""
      .WriteProperty "HideKeyColumns", HideKeyColumns, False
      .WriteProperty "CurrentRecordColor", CurrentRecordColor, vbWindowBackground
      .WriteProperty "DisableEditBox", DisableEditBox, False
      .WriteProperty "RecordPosition", RecordPosition, 0
      .WriteProperty "AllowDelete", AllowDelete, False
      .WriteProperty "LockedEdit", LockedEdit, True
      .WriteProperty "DontPaint", DontPaint, False
      .WriteProperty "GridFont", GridFont, Nothing
      .WriteProperty "BackColor", BackColor, vbApplicationWorkspace
      .WriteProperty "HeaderTextColor", HeaderTextColor, vbWindowText
      .WriteProperty "TextColor", TextColor, vbWindowText
      .WriteProperty "BorderStyle", BorderStyle, 1
      .WriteProperty "DataSource", DataSource, Nothing
      .WriteProperty "GridColor", GridColor, vbButtonFace
      .WriteProperty "TableColor", TableColor, vbWindowBackground
      .WriteProperty "DrawGridLines", DrawGridLines, True
      .WriteProperty "DrawColumnHeaders", DrawColumnHeaders, True
      .WriteProperty "DrawIndicator", DrawIndicator, True
      .WriteProperty "NullString", NullString, "Null"
      .WriteProperty "RowHeight", RowHeight, 18
      .WriteProperty "IndicatorWidth", IndicatorWidth, 16
   End With
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' select row
   
   bMouseDown = True
   bDisableMouseUp = False
   
   If Button = 1 Then
      
      Dim lx&, i&
      
      If Y < mvarRowHeight And X > mvarIndicatorWidth Then
         For i = Abs(lxCol) + 1 To Columns.Count
            lx = lx + Columns(i).ColumnWidth
            If X < lx + mvarIndicatorWidth + 7 And X > lx + mvarIndicatorWidth - 7 Then
               bColSizing = True
               lCurrentColSized = i
               ShowEditBox False ' edit box looks very buggy without this
               Exit For
            Else
               ' begin selection
               ClearSelection
               If Not bColSizing Then _
                  bColSelecting = True ' extra checks
            End If
         Next i
      End If
      
      Dim sch&
      
      If hscrTable.Visible Then sch = hscrTable.Top Else sch = Screen.Height
      
      If X < mvarIndicatorWidth And Y < sch And Y > mvarRowHeight Then
         ClearColSelection
         bSelecting = True
      Else
         Dim lcWidth&
         For i = 1 To Columns.Count
            lcWidth = lcWidth + Columns(i).ColumnWidth
         Next i
         
         If Y > mvarRowHeight And X < lcWidth + mvarIndicatorWidth + lxPos Then  ' dont allow to click on column
            mvarEditMode = True
            mvarCurrentRow = HitTestRow(X, Y)
            mvarCurrentCol = HitTestColumn(X)
            
            ClearSelection
            ClearColSelection
            
            ActivateEditBox mvarCurrentRow, mvarCurrentCol
            
            Refresh
         End If
      End If
         
      If bSelecting Then
         
         ShowEditBox False
         
         lFirstSelected = HitTestRow(X, Y)
         
         If Not bShiftKey Then lFirstSelectedShift = -1
         If lFirstSelectedShift = -1 Then _
            lFirstSelectedShift = lFirstSelected
         
         UserControl_MouseMove Button, Shift, X, Y
      End If
      
      If bColSelecting Then
         
         ShowEditBox False
         
         lColFirstSelected = HitTestColumn(X) - 1
         
         If Not bShiftKey Then lColFirstSelectedShift = -1
         If lColFirstSelectedShift = -1 Then _
            lColFirstSelectedShift = lColFirstSelected
         
         UserControl_MouseMove Button, Shift, X, Y
      End If
   End If
   
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With UserControl
      Dim CP&, i&, lx&, c&
      
      .MousePointer = 0
      
      If X < mvarIndicatorWidth And Y > mvarRowHeight Then
         If .Tag <> "RIGHT" Then .MouseIcon = LoadResPicture("RIGHT", vbResCursor): .Tag = "RIGHT"
         .MousePointer = 99
      End If
      
      If Not bSelecting Then
         If Y < mvarRowHeight And X > mvarIndicatorWidth Then
            Dim bSizeCur As Boolean
            For i = Abs(lxCol) + 1 To Columns.Count
               lx = lx + Columns(i).ColumnWidth
               If X < lx + mvarIndicatorWidth + 7 And X > lx + mvarIndicatorWidth - 7 Then
                  If .Tag <> "SIZE_LR" Then .MouseIcon = LoadResPicture("SIZE_LR", vbResCursor): .Tag = "SIZE_LR"
                  .MousePointer = 99
                  bSizeCur = True
                  Exit For
               End If
            Next i
            If Not bSizeCur And Not bColSizing Then
               If .Tag <> "TOP" Then .MouseIcon = LoadResPicture("TOP", vbResCursor): .Tag = "TOP"
               .MousePointer = 99
            End If
         End If
      End If
   
      ' Size columns
      If bColSizing Then
         lx = 0
         For c = Abs(lxCol) + 1 To lCurrentColSized - 1
            lx = lx + Columns(c).ColumnWidth
         Next c
         
         Columns(lCurrentColSized).ColumnWidth = X - mvarIndicatorWidth - lx
         
         If Columns(lCurrentColSized).ColumnWidth < 10 Then _
            Columns(lCurrentColSized).ColumnWidth = 10
         
         Refresh
      End If
   
   
      Dim lf&, ls&
   
      ' Selection
      If bMouseDown And bSelecting Then
         .MousePointer = 99
         
         mvarSelectedRow = HitTestRow(X, Y)
         
         If Not bCtrlKey And Not bShiftKey Then _
            ClearSelection
            
         ' There are strict rules with shift selecting
         If lFirstSelected < mvarSelectedRow Then
            ShiftSelect lFirstSelected, mvarSelectedRow
         Else
            ShiftSelect mvarSelectedRow, lFirstSelected
         End If
         
         If bShiftKey Then
               ClearSelection
               
               lf = lFirstSelectedShift
               ls = mvarSelectedRow
                              
               If lf < ls Then
                  ShiftSelect lf, ls
               Else
                  ShiftSelect ls, lf
               End If
         End If
         
         Refresh
      End If
      
      ' Column Selection
      If bMouseDown And bColSelecting And Not bSizeCur And Not bColSizing Then
         .MousePointer = 99 ' ?????
         
         mvarSelectedCol = HitTestColumn(X) - 1
         
         If Not bCtrlKey And Not bShiftKey Then _
            ClearColSelection
            
         If lColFirstSelected < mvarSelectedCol Then
            ShiftSelectColumn lColFirstSelected, mvarSelectedCol
         Else
            ShiftSelectColumn mvarSelectedCol, lColFirstSelected
         End If
         
         If bShiftKey Then
               ClearColSelection
               
               lf = lColFirstSelectedShift
               ls = mvarSelectedCol
                              
               If lf < ls Then
                  ShiftSelectColumn lf, ls
               Else
                  ShiftSelectColumn ls, lf
               End If
         End If
         
         Refresh
      End If
   End With
   
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If (mvarDataSource Is Nothing) Then Exit Sub
   If Not bDisableMouseUp Then
      bMouseDown = False
      bSelecting = False
      bColSelecting = False
      bColSizing = False
      
      If Button = 1 Then
         If Y > mvarRowHeight Then
            mvarSelectedRecord = HitTestRow(X, Y)
            If mvarSelectedRecord < 0 Then mvarSelectedRecord = 0
            mvarDataSource.Move mvarSelectedRecord, 1
            Refresh
         End If
      End If
   
      UserControl.MousePointer = 0
   
      bDisableMouseUp = False
   End If
   
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub ShiftSelect(lStart As Long, lEnd As Long)
Attribute ShiftSelect.VB_Description = "Selects records from one row to another, first row must be smaller or error will occur."
   Dim i&, c&

   If lStart < 0 Then lStart = 0: If lEnd < 0 Then lEnd = 0

   
   For c = lStart To lEnd
      For i = 1 To Selection.Count
         If Selection(i).Key = "num:" & CStr(c) Then GoTo already_there
      Next i
      
      Selection.Add "num:" & CStr(c)
      
      ' dont use recordset, we need speed here, use raw data instead
      For i = 0 To mvarDataSource.Fields.Count - 1
         Selection("num:" & CStr(c)).Value(i) = mvarData(i, c)
      Next i
already_there:
   Next c
End Sub

Public Sub ShiftSelectColumn(lStart As Long, lEnd As Long)
Attribute ShiftSelectColumn.VB_Description = "Selects columns from one column to another, first column must be smaller or error will occur."
   Dim i&, c&

   If lStart < 0 Then lStart = 0: If lEnd < 0 Then lEnd = 0

   
   For c = lStart To lEnd
      For i = 1 To ColSelection.Count
         If ColSelection(i).Key = "num:" & CStr(c) Then GoTo already_there
      Next i
      
      ColSelection.Add "num:" & CStr(c)
      
      ColSelection("num:" & CStr(c)).Value(0) = Columns(c + 1).Caption
      ColSelection("num:" & CStr(c)).Value(1) = mvarDataSource.Fields(c).Name
already_there:
   Next c
End Sub

Public Function ActivateEditBox(lRow As Long, lCol As Long) As Boolean
Attribute ActivateEditBox.VB_Description = "Activates edit box for editing data."
   ' returns false if position cannot be found
   If mvarDisableEditBox Then Exit Function
   
   If Not Columns(lCol).Visible Then _
      ActivateEditBox = False: Exit Function
   
   If lRow > mvarRowCount - 1 Or lCol > Columns.Count Or lCol <= 0 Then _
      ActivateEditBox = False: Exit Function
           
   ' now, rows are 0-based, but cols 1-based :), keep that in mind
   With txtEdit
      Dim lx&, ly&, lw&, lh&, c&
      
      ly = (lRow + 1) * mvarRowHeight + (lyPos * mvarRowHeight)
      
      For c = 1 To lCol - 1
         lx = lx + Columns(c).ColumnWidth
      Next c
      lx = lx + mvarIndicatorWidth + lxPos
      
      lw = Columns(lCol).ColumnWidth
      lh = mvarRowHeight
      
      .Move lx + 1, ly + 1, lw - 2, lh - 2
      
      mvarDataSource.Move lRow, 1
      bDisableMouseUp = True
            
      If Not IsNull(mvarDataSource.Fields(lCol - 1).Value) Then
         .Text = Columns(lCol).GetFormattedValue((mvarDataSource.Fields(lCol - 1).Value))
         If Not mvarLockedEdit Then .Locked = False
      Else
         .Locked = True
         .Text = mvarNullString
      End If
      
      .SelStart = 0: .SelLength = Len(.Text)
      
      Select Case Columns(lCol).DataAlign
         Case DT_LEFT: .Alignment = 0
         Case DT_RIGHT: .Alignment = 1
         Case DT_CENTER: .Alignment = 2
      End Select
      
      If lCol <= Abs(lxCol) Then
         .Visible = False
      Else
         .Visible = True
         .SetFocus
      End If
      
   End With
End Function

Public Function HitTestColumn(X As Single) As Long
Attribute HitTestColumn.VB_Description = "Gives index of the column that is currently under specified coordinates."
   ' returns 0 if no column is selected
   Dim cx&, i&
   
   For i = 1 To Columns.Count
      cx = cx + Columns(i).ColumnWidth
      
      If X >= (cx - Columns(i).ColumnWidth) + mvarIndicatorWidth + lxPos _
      And X <= cx + mvarIndicatorWidth + lxPos Then
         HitTestColumn = i
         Exit Function
      End If
   Next i
End Function

Public Function HitTestRow(X As Single, Y As Single) As Long
Attribute HitTestRow.VB_Description = "Gives position of the row that is currently under specified coordinates."
   Dim rc&, CP&, i&
  
   ' now we must do intersection, it must have best precision
   With UserControl
      rc = .ScaleHeight / mvarRowHeight + 1
      If mvarRowCount + lyPos < rc Then rc = mvarRowCount + lyPos
   
      If Y > .ScaleHeight Or Y > rc * mvarRowHeight Then _
         HitTestRow = rc - lyPos - 1: Exit Function ' just return the last one, -1 is column headers panel
   
      For i = 0 To rc
         CP = i * mvarRowHeight
         If Y > CP And Y < CP + mvarRowHeight Or Y = CP Then
            HitTestRow = (Abs(lyPos) + CP / mvarRowHeight) - 1 ' -1 cuz there is row with columns
            Exit Function
         End If
      Next i
   End With
End Function
Private Sub UserControl_Paint()
   If mvarDontPaint Then Exit Sub
   
   Refresh
End Sub


Private Sub UserControl_Terminate()
   DeleteObject lBmp
   DeleteDC dcBuff
   GDIObjDelete
End Sub



Public Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the control."
   ' draw the part that user can see
   Dim c&, r&, rc&, tablew&
   Dim lcWidth&
   Dim hdc&
   Dim rtWnd As RECT
   Dim rtBuff As RECT
   Dim pt(1 To 2) As POINTAPI, ptold As POINTAPI
   
   With UserControl
      If Not Ambient.UserMode Then Exit Sub ' leave if design-time
      
      hdc = GetDC(.hwnd)
         
      ' get client window rect
      GetClientRect .hwnd, rtWnd
      
      ' fill it with default color
      SelectObject dcBuff, mbrushBackColor
      SelectObject dcBuff, mpenWindowText
      Rectangle dcBuff, 0, 0, rtWnd.Right, rtWnd.Bottom
      
      If Not .Enabled Then GoTo finish
      
      ' calculate columns width
      For c = 0 To mvarColumnCount - 1
         lcWidth = lcWidth + Columns(c + 1).ColumnWidth
      Next c
      
      ' calculate how much rows we will be drawing here
      rc = CLng(.ScaleHeight / mvarRowHeight) + 1
      If mvarRowCount + lyPos < rc Then rc = mvarRowCount + lyPos
            
      If (mvarDataSource Is Nothing) Then GoTo finish
            
      ' data
      For r = 0 To rc - 1
      
         lcWidth = 0
                  
         For c = 0 To mvarColumnCount - 1
            If Columns(c + 1).Visible Then
               pt(1).X = mvarIndicatorWidth + lxPos + lcWidth - 1
               pt(1).Y = r * mvarRowHeight + mvarRowHeight - 1
               pt(2).X = mvarIndicatorWidth + lxPos + lcWidth + Columns(c + 1).ColumnWidth
               pt(2).Y = r * mvarRowHeight + mvarRowHeight + mvarRowHeight
               
               If mvarDrawGridLines Then
                  SelectObject dcBuff, mpenGridColor
                  SelectObject dcBuff, mbrushTableColor  ' CELL COLOR
                  Rectangle dcBuff, pt(1).X, pt(1).Y, pt(2).X, pt(2).Y
               End If
                 
               Dim strVal$
               
               If Not IsNull(mvarData(c, -lyPos + r)) Then
                  strVal = mvarData(c, -lyPos + r)
               Else
                  strVal = NullString
               End If
               
               With Columns(c + 1)
                  strVal = .GetFormattedValue(strVal)
               End With
               
   '            If TextWidth(strVal) + 4 > Columns(c + 1).ColumnWidth Then
   '                yep, we must cut it, the text is too big...
   '               Do Until TextWidth(strVal) + TextWidth("...") + 4 <= Columns(c + 1).ColumnWidth
   '                  strVal = Left(strVal, Len(strVal) - 1)
   '               Loop
   '               strVal = strVal & "..."
   '            End If
               
               Dim rtText As RECT
               With rtText
                  .Left = pt(1).X + 3
                  .Right = pt(2).X - 2
                  .Top = pt(1).Y + 2
                  .Bottom = pt(2).Y
               End With
               
               SelectObject dcBuff, fntGrid
               SetTextColor dcBuff, mvarTextColor
               DrawTextA dcBuff, strVal, -1, rtText, Columns(c + 1).DataAlign Or DT_VCENTER
               
               lcWidth = lcWidth + Columns(c + 1).ColumnWidth
            End If
         Next c
      Next r
      
'      ' selected record, works, but I need transparent rectangle api
'      SetBkMode dcBuff, 1
'      Rectangle dcBuff, mvarIndicatorWidth, _
         (mvarSelectedRecord + lyPos) * mvarRowHeight + mvarRowHeight - 1, _
         lcWidth + mvarIndicatorWidth + lxPos, _
         (mvarSelectedRecord + lyPos) * mvarRowHeight + mvarRowHeight * 2
      
      
      lcWidth = 0
      ' column headers
      For c = 0 To mvarColumnCount - 1
         If mvarDrawColumnHeaders And Columns(c + 1).Visible Then
            SelectObject dcBuff, mbrushButtonFace
            SelectObject dcBuff, mpenWindowText
            Rectangle dcBuff, _
               lxPos + lcWidth + mvarIndicatorWidth - 1, 0, _
               lxPos + lcWidth + Columns(c + 1).ColumnWidth + mvarIndicatorWidth, mvarRowHeight
            ' 3d
            SelectObject dcBuff, mpen3DLight
            ' vertical
            MoveToEx dcBuff, lxPos + lcWidth + mvarIndicatorWidth, 1, ptold
            LineTo dcBuff, _
               lxPos + lcWidth + mvarIndicatorWidth, mvarRowHeight - 1
            ' horizontal
            MoveToEx dcBuff, lxPos + lcWidth + mvarIndicatorWidth + 1, 1, ptold
            LineTo dcBuff, _
               lxPos + lcWidth + mvarIndicatorWidth + 1 + Columns(c + 1).ColumnWidth, 1
         End If
         
         Dim strHeader$
         
         strHeader = Columns(c + 1).Caption
         
         If TextWidth(strHeader) + 4 > Columns(c + 1).ColumnWidth Then
            ' damn, now we must cut it, the text is too big...
            Do Until TextWidth(strHeader) + TextWidth("...") + 4 <= Columns(c + 1).ColumnWidth
               If Len(strHeader) Then Exit Do
               strHeader = Left(strHeader, Len(strHeader) - 1)
            Loop
            strHeader = strHeader & "..."
         End If
         
         SetTextColor dcBuff, mvarHeaderTextColor
         DrawTextA dcBuff, strHeader, Len(strHeader), _
            Pt2Rt(lxPos + lcWidth + mvarIndicatorWidth, 2, _
                  lxPos + lcWidth + mvarIndicatorWidth + 4 + Columns(c + 1).ColumnWidth, _
                  CLng(mvarIndicatorWidth)), _
            Columns(c + 1).CaptionAlign Or DT_VCENTER

         lcWidth = lcWidth + Columns(c + 1).ColumnWidth
      Next c


      ' indicator (the thingy to the left where we select whole rows)
      rc = CLng(.ScaleHeight / mvarRowHeight) + 1
      If mvarRowCount + lyPos < rc Then rc = mvarRowCount + lyPos

      If mvarDrawIndicator Then
         If mvarRowCount > 0 Then
            For r = 0 To rc - 1
               SelectObject dcBuff, mpenWindowText
               Rectangle dcBuff, 0, r * mvarRowHeight + mvarRowHeight - 1, mvarIndicatorWidth, r * mvarRowHeight + mvarRowHeight + mvarRowHeight
               '3d
               SelectObject dcBuff, mpen3DLight
               ' vertical
               MoveToEx dcBuff, 1, r * mvarRowHeight + mvarRowHeight, ptold
               LineTo dcBuff, 1, r * mvarRowHeight + 2 * mvarRowHeight - 1
               ' horizontal
               MoveToEx dcBuff, 1, r * mvarRowHeight + mvarRowHeight, ptold
               LineTo dcBuff, mvarIndicatorWidth - 1, r * mvarRowHeight + mvarRowHeight
            Next r
            SelectObject dcBuff, mpenWindowText
            Rectangle dcBuff, 0, 0, mvarIndicatorWidth, mvarRowHeight
            
            SelectObject dcBuff, mpen3DLight
            ' vertical
            MoveToEx dcBuff, 1, 1, ptold
            LineTo dcBuff, 1, mvarRowHeight - 1
            ' horizontal
            MoveToEx dcBuff, 1, 1, ptold
            LineTo dcBuff, mvarIndicatorWidth - 1, 1
         End If
      End If
      
      ' a little indicator of current selection
      If mvarSelectedRecord + lyPos + 1 > 0 Then _
         BitBlt dcBuff, 4, mvarRowHeight * (mvarSelectedRecord + lyPos) + mvarRowHeight + 3, _
            6, 11, mhdcIndicator, 0, 0, vbSrcAnd

      ' selection
      Dim i&
      For i = 1 To Selection.Count
         Dim lnum&: lnum = CLng(Mid(Selection(i).Key, 5))
         
         If lnum + lyPos >= 0 Then
            BitBlt dcBuff, 0, (lnum + 1 + lyPos) * mvarRowHeight, _
               lcWidth + mvarIndicatorWidth + lxPos, mvarRowHeight, dcBuff, _
                0, (lnum + 1 + lyPos) * mvarRowHeight, vbDstInvert
         End If
      Next i
      
      For i = 1 To ColSelection.Count
         Dim cw&: cw = 0 ' memory will not be erased, vb i stupid, u never know where and when this might occur
         For c = 1 To CLng(Mid(ColSelection(i).Key, 5)) + 1
            cw = cw + Columns(c).ColumnWidth
         Next c
         cw = cw - Columns(CLng(Mid(ColSelection(i).Key, 5)) + 1).ColumnWidth _
            + lxPos + mvarIndicatorWidth
         
         If Not lxCol = -CLng(Mid(ColSelection(i).Key, 5)) - 1 Then _
            If Not cw + lxPos >= .ScaleWidth Then _
               BitBlt dcBuff, cw, 0, _
                  Columns(CLng(Mid(ColSelection(i).Key, 5)) + 1).ColumnWidth _
                  , (rc + 1) * mvarRowHeight, dcBuff, 0, 0, vbDstInvert
      Next i
          
      ' oh, and cover that little hole in the buttom right corner :)
      If hscrTable.Visible Or vscrTable.Visible Then
         SelectObject dcBuff, mpenButtonFace
         SelectObject dcBuff, mbrushButtonFace
         Rectangle dcBuff, .ScaleWidth - vscrTable.Width, .ScaleHeight - hscrTable.Height _
         , .ScaleWidth, .ScaleHeight
      End If
      
finish:
      BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, dcBuff, 0, 0, vbSrcCopy
   
      ReleaseDC .hwnd, hdc
   
      ' BUG BUG ERROR SHIT CRAP ???? !!!!
      'If mvarEditMode Then _
         ActivateEditBox mvarCurrentRow, mvarCurrentCol
   End With
End Sub


Private Sub hscrTable_Change()
   lxCol = -hscrTable.Value
   Dim i&
   lxPos = 0
   For i = 1 To Abs(lxCol)
      lxPos = lxPos - Columns(i).ColumnWidth
   Next i
   Refresh
End Sub

Private Sub hscrTable_Scroll()
   hscrTable_Change
End Sub

Private Sub UserControl_Initialize()
   ' change some default data
   With UserControl
      .ScaleMode = vbPixels
   End With

   TableColor = vbWindowBackground
   
   Set mvarGridFont = UserControl.Font
   
   bDisableMouseUp = True
   
   If Not bInitialized Then
      
      lFirstSelectedShift = -1
      
      GDIObjDelete
      GDIObjCreate
   
      Dim tmpDC&: tmpDC = GetDC(0)
      dcBuff = CreateCompatibleDC(0)
      lBmp = CreateCompatibleBitmap(tmpDC, Screen.Width / Screen.TwipsPerPixelX, _
         Screen.Height / Screen.TwipsPerPixelY)
      SelectObject dcBuff, lBmp
      ReleaseDC 0, tmpDC
         
      SetBkMode dcBuff, 1
      
      If Columns.Count = 0 Then
         Columns.Add
      End If
      
      bInitialized = True
   End If
End Sub




Private Sub UserControl_Resize()
On Error Resume Next
   With UserControl
      .Tag = .ScaleHeight
      vscrTable.Move .ScaleWidth - vscrTable.Width, 0, vscrTable.Width, .ScaleHeight - hscrTable.Height
      hscrTable.Move 0, .ScaleHeight - hscrTable.Height, .ScaleWidth - vscrTable.Width, hscrTable.Height
   
      Dim c&, lcWidth&
      For c = 0 To mvarColumnCount - 1
         lcWidth = lcWidth + Columns(c + 1).ColumnWidth
      Next c
   
      With hscrTable
         .Max = Columns.Count - 1
         .SmallChange = 1
         .LargeChange = 4
      End With
      If .ScaleWidth - 32 > lcWidth Then _
         hscrTable.Visible = False Else _
         hscrTable.Visible = True
      
      With vscrTable
         .Max = mvarRowCount - 1
         .SmallChange = 1
         .LargeChange = 4
      End With
      If .ScaleHeight - 32 > mvarRowCount * mvarRowHeight Then _
         vscrTable.Visible = False Else _
         vscrTable.Visible = True
         
   End With
   
   Refresh
   
On Error Resume Next
   UserControl.SetFocus ' vb scroll bar bug >:(

   RaiseEvent Resize
End Sub


Private Sub vscrTable_Change()
   ShowEditBox False
   
   lyPos = -vscrTable.Value
   Refresh
End Sub

Private Sub vscrTable_KeyDown(KeyCode As Integer, Shift As Integer)
   UserControl_KeyDown KeyCode, Shift
End Sub

Private Sub vscrTable_KeyUp(KeyCode As Integer, Shift As Integer)
   UserControl_KeyUp KeyCode, Shift
End Sub

Private Sub hscrTable_KeyDown(KeyCode As Integer, Shift As Integer)
   UserControl_KeyDown KeyCode, Shift
End Sub

Private Sub hscrTable_KeyUp(KeyCode As Integer, Shift As Integer)
   UserControl_KeyUp KeyCode, Shift
End Sub


Private Sub vscrTable_Scroll()
   vscrTable_Change
End Sub



Private Sub GDIObjCreate()
   mpenGridColor = CreatePen(vbSolid, 1, TranslateColor(mvarGridColor))
   mpenWindowText = CreatePen(vbSolid, 1, TranslateColor(vbWindowText))
   mpenButtonFace = CreatePen(vbSolid, 1, TranslateColor(vbButtonFace))
   mpen3DLight = CreatePen(vbSolid, 1, TranslateColor(vb3DLight))
   mpenGridText = CreatePen(vbSolid, 1, TranslateColor(mvarHeaderTextColor))
   mpenHeaderText = CreatePen(vbSolid, 1, TranslateColor(mvarTextColor))
   
   mbrushButtonFace = CreateSolidBrush(TranslateColor(vbButtonFace))
   mbrushTableColor = CreateSolidBrush(TranslateColor(mvarTableColor))
   mbrushWindowBackground = CreateSolidBrush(TranslateColor(vbWindowBackground))
   mbrushAppWorkspace = CreateSolidBrush(TranslateColor(vbApplicationWorkspace))
   mbrushBackColor = CreateSolidBrush(TranslateColor(mvarBackColor))

   mhdcIndicator = CreateCompatibleDC(0)
   SelectObject mhdcIndicator, LoadResPicture("INDICATOR", vbResBitmap).Handle

   With mvarGridFont
      If Not (mvarGridFont Is Nothing) Then _
         fntGrid = FromStdFont(.Name, .Size, 0, .Bold, .Italic, .Underline, .Charset)
   End With
End Sub

Private Sub GDIObjDelete()
   DeleteObject (mpenGridColor)
   DeleteObject (mpenWindowText)
   DeleteObject (mbrushButtonFace)
   DeleteObject (mbrushTableColor)
   DeleteObject (mbrushWindowBackground)
   DeleteObject (mbrushAppWorkspace)
   DeleteObject (mpenButtonFace)
   DeleteObject (mpen3DLight)
   DeleteObject (mpenGridText)
   DeleteObject (mpenHeaderText)
   DeleteObject (mbrushBackColor)
   
   DeleteDC mhdcIndicator
   
   DeleteObject fntGrid
End Sub

Public Sub ShowEditBox(Optional bShow As Boolean = True)
Attribute ShowEditBox.VB_Description = "Displays the edit box at the position it was last time."
   txtEdit.Visible = bShow
   mvarEditMode = bShow
End Sub

Public Sub ClearSelection()
Attribute ClearSelection.VB_Description = "Deselects all selected rows."
   Dim i&
del:
   For i = 1 To Selection.Count
      Selection.Remove i
      GoTo del
   Next i
End Sub

Public Sub ClearColSelection()
Attribute ClearColSelection.VB_Description = "Deselects all selected columns."
   Dim i&
del:
   For i = 1 To ColSelection.Count
      ColSelection.Remove i
      GoTo del
   Next i
End Sub

Public Sub RefreshData()
Attribute RefreshData.VB_Description = "Refreshes raw data used for displaying grid."
   bDeleteing = True
   mvarDataSource.Requery
   LoadData
End Sub

Public Sub UpdateData(Row As Long, Col As Long, Optional MoveToRecord As Boolean = False, Optional ApplyRefresh As Boolean = True)
Attribute UpdateData.VB_Description = "Simply updates raw data used for displaying records."
   If MoveToRecord Then _
      mvarDataSource.Move Row + 1, 1

   mvarData(Col, Row) = _
      mvarDataSource.Fields(Col).Value
   
   If ApplyRefresh Then _
      Refresh
End Sub


Public Sub NewAdded()
Attribute NewAdded.VB_Description = "For better performance, this function will simply add new row and require new record to fill the newly created row with data. After adding single record to DataSource call this function to refresh grid faster so new record will appear at the end."
   ReDim Preserve mvarData(UBound(mvarData, 1), UBound(mvarData, 2) + 1)
   
   mvarDataSource.MoveLast
   
   Dim i&
   With mvarDataSource
      For i = 0 To .Fields.Count - 1
         mvarData(i, mvarRowCount) = .Fields(i).Value
      Next i
   End With
         
   mvarRowCount = mvarRowCount + 1

   Refresh
End Sub

Public Function IsDatasetEmpty() As Boolean
Attribute IsDatasetEmpty.VB_Description = "Returns true if DataSource has no records or there was an error opening it."
   If (mvarRowCount = 0) Then IsDatasetEmpty = True
End Function

Public Sub About()
Attribute About.VB_Description = "Displays About Box."
Attribute About.VB_UserMemId = -552
   frmAbout.Show vbModal, Me
End Sub
