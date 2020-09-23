VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextBoxLimiter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mvarUsedTextBox As TextBox
Attribute mvarUsedTextBox.VB_VarHelpID = -1

Private mvarDecimalPlaces As Long

Private Sub Class_Initialize()
   mvarDecimalPlaces = 2
End Sub

Public Property Get UsedTextBox() As TextBox
   Set UsedTextBox = mvarUsedTextBox
End Property

Public Property Set UsedTextBox(ByVal vNewValue As TextBox)
   Set mvarUsedTextBox = vNewValue
End Property


Public Property Get DecimalPlaces() As Variant
   DecimalPlaces = mvarDecimalPlaces
End Property

Public Property Let DecimalPlaces(ByVal vNewValue As Variant)
   mvarDecimalPlaces = vNewValue
End Property

Private Sub mvarUsedTextBox_LostFocus()
   With mvarUsedTextBox
      .Text = FormatNumber(.Text, mvarDecimalPlaces)
   End With
End Sub
