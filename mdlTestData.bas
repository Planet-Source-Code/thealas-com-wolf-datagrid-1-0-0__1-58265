Attribute VB_Name = "mdlTestData"
Option Explicit

Public strConnectionString As String
Public strDataBaseFile As String

Public adoActiveAdo As Adodc
Public adoActiveGrid As DataGrid
Public adoConnection As New ADODB.Connection

Public Knjiga As New CBooks

Public blShowIDS As Boolean

Public Const ERR_SQL As String = "Error in sql syntax!"

Public Function RaiseErr(CausedBy As String, ErrorObject As ErrObject, Optional CanContinueOrResume As Boolean = False) As VbMsgBoxResult
   
   Dim MsgRes As VbMsgBoxResult
   Dim Buttons As VbMsgBoxStyle
   
   Buttons = vbOKOnly
   
   If CanContinueOrResume Then Buttons = vbAbortRetryIgnore
   
   MsgRes = MsgBox(CausedBy & ":" & vbCrLf & vbCrLf & "Broj Greske: " & ErrorObject.Number & vbCrLf & vbCrLf & "Opis greske:" & vbCrLf & ErrorObject.Description, Buttons, "Greska")
   
   RaiseErr = MsgRes
End Function

Public Sub RunSQL(Statement As String)
   frmSql.txtSql = Statement
   frmSql.WindowState = vbMinimized
   frmSql.cmdIzvrsi_Click
End Sub

Public Function OpenSqlFile(strFile As String) As String
   Dim buf$
   Open App.Path & "\" & strFile & ".sql" For Input As #1
      Do Until EOF(1)
         Line Input #1, buf
         OpenSqlFile = OpenSqlFile & buf & vbCrLf
      Loop
   Close #1
End Function

Public Sub SaveSqlFile(strFile As String, Command As String)
   Dim buf$
   Open App.Path & "\" & strFile & ".sql" For Output As #1
      Print #1, Command
   Close #1
End Sub

Public Function ParseDir(strFileOrServer As String) As String
Dim pos As Integer
    pos = InStrRev(strFileOrServer, "\", , vbTextCompare)
    ParseDir = Mid(strFileOrServer, 1, pos)
End Function

Public Sub ComboData(cCombo As ComboBox, rs As ADODB.Recordset, Optional sField As Long = 1)
   rs.MoveFirst
   Dim arr: arr = rs.GetRows
   Dim i&
   For i = 0 To UBound(arr, 2)
      cCombo.AddItem CStr(arr(sField, i))
      cCombo.ItemData(i) = arr(0, i) ' ID
   Next i
   
   cCombo.ListIndex = 0
End Sub

Public Sub SelectComboItem(cCombo As ComboBox, strItem As String, Optional rsBook As CBook = Nothing)
   Dim i&, c&: c = -1
   For i = 0 To cCombo.ListCount - 1
      If cCombo.List(i) = strItem Then
         c = i
         Exit For
      End If
   Next i
   
   If c > -1 Then
      cCombo.ListIndex = c
      If Not (rsBook Is Nothing) Then
         With rsBook
            .RecordSource.Move c, 1
         End With
      End If
   End If
End Sub

Public Function CheckForComboExistance(cCombo As ComboBox, Optional rsBook As CBook = Nothing) As Long
   Dim i&
   CheckForComboExistance = -1
   With cCombo
      For i = 0 To .ListCount - 1
         If .List(i) = .Text Then
            CheckForComboExistance = .ItemData(i)
            .ListIndex = i
            If Not (rsBook Is Nothing) Then _
               rsBook.RecordSource.Move i, 1
            Exit Function
         End If
      Next i
   End With
End Function
