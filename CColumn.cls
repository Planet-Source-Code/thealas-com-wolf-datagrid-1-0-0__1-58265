VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Column object for DataGrid"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'# this class was generated by VB6 Class Builder

Private mvarColumnWidth As Integer 'local copy
Private mvarDataType As DataTypeEnum 'local copy
Private mvarCaption As String
Private mvarDecimalPlaces As Integer
Private mvarDataAlign As Long
Private mvarCaptionAlign As Long
Private mvarKey As String
Private mvarIndex As Long
Private mvarVisible As Boolean

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns/sets if column is drawn on table."
   Visible = mvarVisible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
   mvarVisible = vNewValue
End Property


Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns left to right index of column."
   Index = mvarIndex
End Property

Public Property Let Index(ByVal vNewValue As Long)
   mvarIndex = vNewValue
End Property


Public Property Get Key() As String
Attribute Key.VB_Description = "Returns/sets key string used for this column."
   Key = mvarKey
End Property

Public Property Let Key(ByVal vNewValue As String)
   mvarKey = vNewValue
End Property


Public Property Get CaptionAlign() As Long
Attribute CaptionAlign.VB_Description = "Returns/sets alignment of column caption."
   CaptionAlign = mvarCaptionAlign
End Property

Public Property Let CaptionAlign(ByVal vNewValue As Long)
   mvarCaptionAlign = vNewValue
End Property


Public Property Get DataAlign() As Long
Attribute DataAlign.VB_Description = "Returns/sets alignment used for drawing data associated with this column (field)."
   DataAlign = mvarDataAlign
End Property

Public Property Let DataAlign(ByVal vNewValue As Long)
   mvarDataAlign = vNewValue
End Property


Public Property Get DecimalPlaces() As Integer
Attribute DecimalPlaces.VB_Description = "Returns/sets decimal places used for currency data type."
   DecimalPlaces = mvarDecimalPlaces
End Property

Public Property Let DecimalPlaces(ByVal vNewValue As Integer)
   mvarDecimalPlaces = vNewValue
End Property


Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets caption of this column, used for drawing."
   Caption = mvarCaption
End Property

Public Property Let Caption(ByVal vNewValue As String)
   mvarCaption = vNewValue
End Property


Public Property Let DataType(ByVal vData As DataTypeEnum)
Attribute DataType.VB_Description = "Returns/sets data type used for this column (field)."
   mvarDataType = vData
   
   If vData = adCurrency Then DataAlign = DT_RIGHT
End Property

Public Property Get DataType() As DataTypeEnum
   DataType = mvarDataType
End Property

Public Property Let ColumnWidth(ByVal vData As Integer)
Attribute ColumnWidth.VB_Description = "Returns/sets width of column."
   mvarColumnWidth = vData
End Property

Public Property Get ColumnWidth() As Integer
   ColumnWidth = mvarColumnWidth
   
   If Not mvarVisible Then ColumnWidth = 0
End Property

Public Function GetFormattedValue(Value As String) As String
Attribute GetFormattedValue.VB_Description = "Returns formatted data for data types such as currency or date, instead of raw data."
On Error GoTo e
   Select Case mvarDataType
   Case adCurrency
      If mvarDecimalPlaces > 0 Then _
         GetFormattedValue = FormatCurrency(Value, mvarDecimalPlaces) _
      Else _
         : GetFormattedValue = FormatCurrency(Value)
   Case adDecimal, adInteger, adNumeric, adSingle, adVarNumeric
      GetFormattedValue = FormatNumber(Value, mvarDecimalPlaces)
   Case Else
      GetFormattedValue = Value
   End Select
Exit Function
e:
   GetFormattedValue = "ERROR"
End Function

Private Sub Class_Initialize()
   mvarCaptionAlign = DT_CENTER
End Sub