Attribute VB_Name = "mdlPub"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000

Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WM_CLOSE = &H10
Public Const WM_FULL = -1

Public Const HDS_BUTTONS = &H2

Public Const SWP_DRAWFRAME = &H20
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME


Public Function Sort_NumericInsert(arr() As Long, arraySize As Long, lDimension As Long) As Long()
   Dim i As Long, j As Long, temp As Long, k As Long, Comparisonval&
   
   For i = 1 To arraySize
      For j = 0 To i - 1
      
         Comparisonval = Comparisonval + 1
         If arr(lDimension, i) < arr(lDimension, j) Then
            temp = arr(lDimension, i)
            For k = i To j + 1 Step -1
               arr(lDimension, k) = arr(lDimension, k - 1)
            Next k
            arr(lDimension, j) = temp
         End If
         
      Next j
   Next i
   Sort_NumericInsert = arr
End Function

Public Sub StaticBorder(lhWnd As Long)
    SetWindowLong lhWnd, (-20), GetWindowLong(lhWnd, (-20)) And Not &H20000 Or &H200&
    SetWindowLong lhWnd, (-20), &H20000
    RemoveBorder lhWnd
End Sub

Public Sub RemoveBorder(lhWnd As Long)
    Dim lStyle As Long
    lStyle = GetWindowLong(lhWnd, GWL_STYLE) And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
    SetWindowLong lhWnd, GWL_STYLE, lStyle
    SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Function lr(id As Long) As String
   lr = LoadResString(id)
End Function
