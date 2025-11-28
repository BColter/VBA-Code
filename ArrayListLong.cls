VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayLIstLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' Module Name: ArrayListLong
' Purpose    : Provides and array list to store longs
' Author     : Burton Colter
' Created    : 6/12/2025 2:26:44 PM (last modified 11/28/2025)
' Notes      : Uses Long, good for moderatly sized arrays < 100000 elements
'            : Please note that this is 0 indexed
'            : When the array reaches capacity, it will double in size
'------------------------------------------------------------------------------
Private DATA() As Long
Private index As Long

'Constructor
Private Sub Class_Initialize()
    ReDim DATA(9)   'Inital size of 10
    index = 0
End Sub

'Add an element to the end of the array
Public Sub Add(item As Long)
    If index >= UBound(DATA) Then ReDim Preserve DATA((UBound(DATA) + 1) * 2) ' Double the array size if full
    DATA(index) = item
    index = index + 1
End Sub

'Return an element from the specified index location
Public Function GetItem(i As Long) As Long
    If i >= 0 And i < index Then
        GetItem = DATA(i)
    Else
        Err.Raise 9, , "Index out of bounds"
    End If
End Function

'Removes specific element form the array and shifts elements
Public Sub Remove(i As Long)
    If i >= 0 And i < index Then
        Dim j As Long
        For j = i To index - 2: DATA(j) = DATA(j + 1): Next j
        index = index - 1
        DATA(index) = 0
    Else
        Err.Raise 9, , "Index out of bounds"
    End If
End Sub

'Clear the using ReDim (destructive)
Public Sub Clear()
    ReDim DATA(9)
    index = 0
End Sub

'Return a trimmed array with no empty slots
Public Function ToArray() As Long()
    Dim trimmed() As Long
    If index = 0 Then
        Err.Raise vbObjectError + 1000, "ArrayList", "No data in array"
    Else
        ReDim trimmed(index - 1)
        Dim i As Long
        For i = 0 To UBound(trimmed): trimmed(i) = DATA(i): Next
        ToArray = trimmed
    End If
End Function

' Check if the value exists in the arraylist, return True if found
Public Function Contains(x As Long) As Boolean
    Contains = (IndexOf(x) >= 0)
End Function

' If value is found, return the index location in the array
Public Function IndexOf(x As Long, Optional startPos As Long = 0) As Long
    Dim i As Long
    IndexOf = -1
    If startPos < 0 Then startPos = 0
    If startPos >= index Then Exit Function
    
    For i = startPos To index - 1
        If DATA(i) = x Then
            IndexOf = i
            Exit Function
        End If
    Next
End Function

' Sort the data ascending
Public Sub Sort(Optional ByVal descending As Boolean = False)
    If index < 2 Then Exit Sub
    QuickSort 0, index - 1
    If descending Then Reverse
End Sub

' Quick sort O(log n)
Private Sub QuickSort(ByVal low As Long, ByVal high As Long)
    Dim pivot As Long, i As Long, j As Long, mid As Long
    
    Do While low < high
        ' Set mid point
        mid = low + (high - low) \ 2 'Note the reverse of divide "/", this will keep it as an integer and not a double eliminating the cast back to long... CRAZY!
        ' Ensure values are ordered correctly, calling Swap as needed
        If DATA(low) > DATA(mid) Then Swap DATA(low), DATA(mid)
        If DATA(low) > DATA(high) Then Swap DATA(low), DATA(high)
        If DATA(mid) > DATA(high) Then Swap DATA(mid), DATA(high)
        
        ' Set Pivot
        pivot = DATA(mid)
        Swap DATA(mid), DATA(high)
        
        i = low: j = high - 1
        Do
            Do While i <= high And DATA(i) < pivot: i = i + 1: Loop
            Do While j >= low And DATA(j) > pivot: j = j - 1: Loop
            If i >= j Then Exit Do
            Swap DATA(i), DATA(j)
            i = i + 1: j = j - 1
        Loop
        ' Reset Pivot
        Swap DATA(i), DATA(high)
        high = i - 1
        ' Recursive Calls
        If (i - low) < (high - i) Then
            QuickSort low, high
            low = i + 1
        Else
            QuickSort i + 1, high
        End If
    Loop
End Sub

' Helper for clean value swap called by QuickSort
Private Sub Swap(ByRef a As Long, ByRef b As Long)
    Dim t As Long
    t = a: a = b: b = t
End Sub

' Reverse the order of the array.
' NOTE: if the data has not been previously sorted, this
' Sub is of no value. Please use with caution
Public Sub Reverse()
    Dim i As Long, j As Long, t As Long
    
    j = index - 1
    For i = 0 To (index \ 2) - 1
        t = DATA(i)
        DATA(i) = DATA(j)
        DATA(j) = t
        j = j - 1
    Next
End Sub

Public Property Get LongArray() As Long()
    LongArray = ToArray()
End Property

'Return the entire array, including empty slots
Public Property Get dataArray() As Long()
    dataArray = DATA
End Property

'Returns the size of the array
Public Property Get size() As Long
    size = index
End Property

'Returns the last index of the array
Public Property Get GetIndex() As Long
    GetIndex = index - 1
End Property

' Get the capacity of the underlying array
Public Property Get Capacity() As Long
    Capacity = UBound(DATA) + 1 ' UBound is 0-based, capacity is 1-based
End Property


'Allows user to set the capacity of the array if known
Public Property Let Capacity(newCapacity As Long)
    If newCapacity < index Then Err.Raise 500, , "Capacity cannot be less than current size"
    If newCapacity < 1 Then Err.Raise 500, , "Capacity must be at least 1"
    ReDim Preserve DATA(newCapacity - 1) ' Adjust to 0-based index
End Property


