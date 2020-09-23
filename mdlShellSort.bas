Attribute VB_Name = "mdlShellSort"
'********************************************************************
' The following code is written by Vladimir S. Pekulas
' http://www.expinion.net/
' You can use the code at your own risk it is completely free
'********************************************************************

Option Explicit

Dim new_first_pointer As Integer, new_last_pointer As Integer
Dim distance As Integer, fwd_comp_limit As Integer, record_count As Integer, Holder As Integer
Dim first_pointer As Integer, last_pointer As Integer, Exchange_switch As Boolean
Private arr() As Long
Private SortDim As Long

Public Function ShellSort(SortArray() As Long, Dimension As Long) As Long()
    arr = SortArray
    SortDim = Dimension
    
    record_count = UBound(arr, 2)

    distance = record_count / 2
    
    Do Until distance < 1
       Call Passes
    Loop
    
    ' shit that I must fix somehow >:(
    
    On Error GoTo e
    Dim i&, pass&
    For i = 0 To record_count
      If arr(Dimension, i) = 0 Then
         pass = arr(Dimension, 0)
         arr(Dimension, 0) = 0
         arr(Dimension, i) = pass
         Exit For
      End If
    Next i
    
e:
    ShellSort = arr
End Function


Private Function Passes()
    fwd_comp_limit = record_count - distance
    first_pointer = 0
   
  Do Until first_pointer > fwd_comp_limit
    Call Forward_Compare
    first_pointer = first_pointer + 1
  Loop
  distance = distance / 2
End Function

Private Function Forward_Compare()
    new_first_pointer = first_pointer
    last_pointer = first_pointer + distance
  
  If Int(arr(SortDim, new_first_pointer)) > Int(arr(SortDim, last_pointer)) Then
     Call Exchange
  End If
End Function

Private Function Exchange()
    Exchange_switch = True
  
  Do Until Exchange_switch = False
    Call Propagate_Swaps
  Loop
End Function

Private Function Propagate_Swaps()
  Call Swap
  Call Back_compare
End Function

Private Function Back_compare()
    new_last_pointer = new_first_pointer
    new_first_pointer = new_first_pointer - distance
    
    If new_first_pointer <= 1 Then
        Exchange_switch = False
    Else
        If Int(arr(SortDim, new_first_pointer)) > Int(arr(SortDim, new_last_pointer)) Then
            last_pointer = new_last_pointer
        Else
            Exchange_switch = False
        End If
    End If
End Function

Private Function Swap()
    Holder = arr(SortDim, new_first_pointer)
    arr(SortDim, new_first_pointer) = arr(SortDim, last_pointer)
    arr(SortDim, last_pointer) = Holder
End Function



