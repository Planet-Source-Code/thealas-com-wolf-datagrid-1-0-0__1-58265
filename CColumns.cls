VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection of DataGrid columns"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CColumn"
Attribute VB_Ext_KEY = "Member0" ,"CColumn"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'# this class was generated by VB6 Class Builder

'local variable to hold collection
Private mCol As Collection

Public Function Add(Optional ColumnWidth As Integer = 100, Optional sKey As String) As CColumn
Attribute Add.VB_Description = "Creates and returns new column object."
   'create a new object
   Dim objNewMember As CColumn
   Set objNewMember = New CColumn
   
   
   'set the properties passed into the method
   objNewMember.ColumnWidth = ColumnWidth
   objNewMember.Key = sKey
   
   If Len(sKey) = 0 Then
      mCol.Add objNewMember
   Else
      mCol.Add objNewMember, sKey
   End If
   
   objNewMember.Index = mCol.Count
   objNewMember.Visible = True
   
   
   'return the object created
   Set Add = objNewMember
   Set objNewMember = Nothing
   
End Function

Public Property Get Item(vntIndexKey As Variant) As CColumn
Attribute Item.VB_Description = "Returns/sets column object associated with specified index or key."
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns total number of columns used in collection."
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_Description = "Deletes column object associated with specified index or key."
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

