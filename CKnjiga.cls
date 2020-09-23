VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private WithEvents mvarRecordSource As Recordset

Public RawData As Variant

Public Property Get RecordSource() As Recordset
Attribute RecordSource.VB_MemberFlags = "200"
   Set RecordSource = mvarRecordSource
End Property

Public Property Set RecordSource(ByVal vNewValue As Recordset)
   Set mvarRecordSource = vNewValue
End Property


Public Sub RefreshRawData()
   If Not (mvarRecordSource Is Nothing) Then _
      RawData = mvarRecordSource.GetRows _
   Else _
      Erase RawData
End Sub

Private Sub mvarRecordSource_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Dim i&
   With mvarRecordSource
      For i = 0 To .Fields.Count - 1
         RawData(i, .AbsolutePosition - 1) = .Fields(i).Value
      Next i
   End With
End Sub

