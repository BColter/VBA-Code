VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayListLong"
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
' Notes      : Good for moderatly sized arrays < 100000 elements
'            : Please note that this is 0 indexed
'            : When the array reaches capacity, it will double in size
'            : Current implementation will only use the binary search when
'            : the data is sorted (fresh), otherwise search will be linear.
'------------------------------------------------------------------------------

Private DATA() As Long
Private Const INITIAL_LAST_INDEX As Long = 9  'Inital size of 10 elements (0-9 zero based array)
Private index As Long
Private isSorted As Boolean
Private isDescending As Boolean

' Error Codes
Private Const ERR_INDEX_OUT_OF_RANGE As Long = vbObjectError + 1
Private Const ERR_STARTPOS_OUT_OF_RANGE As Long = vbObjectError + 2
Private Const ERR_CAPACITY_TOO_SMALL As Long = vbObjectError + 3
Private Const ERR_CAPACITY_INVALID As Long = vbObjectError + 4


' Constructor
Private Sub Class_Initialize()
    ResetArray
End Sub

' Reset Array will clear the data and scale it down to the default size
Private Sub ResetArray()
    ReDim DATA(INITIAL_LAST_INDEX)
    index = 0
    isSorted = False
    isDescending = False
End Sub

' Clear the using ReDim (destructive)
Public Sub Clear()
    ResetArray
End Sub

' Add an element to the end of the array. Sets sorted flag to false
Public Sub Add(Item As Long)
    If index >= UBound(DATA) Then ReDim Preserve DATA(UBound(DATA) * 2 + 1) ' Double the array size if full
    DATA(index) = Item
    index = index + 1
    isSorted = False
End Sub

' Return an element from the specified index location
Public Function GetItem(i As Long) As Long
    If i >= 0 And i < index Then
        GetItem = DATA(i)
    Else
        Err.Raise ERR_INDEX_OUT_OF_RANGE, "ArrayListLong", "Index out of bounds"
    End If
End Function

' Removes specific element form the array and shifts elements. NOTE: O(n), time consuming on large arrays.
' Note that the value will remain in the last location, but the index is adjusted.
Public Sub Remove(i As Long)
    If i >= 0 And i < index Then
        Dim j As Long
        For j = i To index - 2: DATA(j) = DATA(j + 1): Next j
        index = index - 1
    Else
        Err.Raise ERR_INDEX_OUT_OF_RANGE, "ArrayListLong", "Index out of bounds"
    End If
End Sub

' Allows the user to properly manage memory allocation for large data sets. If no data, reset to initial size
' Reseting to initial size prevents a zero length array
Public Sub ShrinkToFit()
    If index > 0 Then
        ReDim Preserve DATA(index - 1)
    Else
        ReDim DATA(INITIAL_LAST_INDEX)
    End If
End Sub

' Return a trimmed array with no empty slots. If the index is zero and empty array will be returned. UBound(trimmed) will return an error
Public Function ToArray() As Long()
    Dim trimmed() As Long
    If index = 0 Then
        ReDim trimmed(0 To -1)
        ToArray = trimmed
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
' Note: If the array is not sorted, it will run an o(n) linear search
' It is recommended that the array be sorted before calling this function
Public Function IndexOf(x As Long, Optional startPos As Long = 0) As Long
    Dim i As Long
    IndexOf = -1
    If startPos < 0 Then Err.Raise ERR_INDEX_OUT_OF_RANGE, "ArrayListLong", "Start positon out of bounds, less than zero"
    If startPos >= index Then
        Err.Raise ERR_STARTPOS_OUT_OF_RANGE, "ArrayListLong", "Start position out of bounds"
        Exit Function
    End If
    
    If isSorted Then
        IndexOf = BinarySearch(x, startPos)
    Else
        For i = startPos To index - 1
            If DATA(i) = x Then
                IndexOf = i
                Exit Function
            End If
        Next
    End If
End Function

' Binary search only used on sorted data, returns -1 if not found.
Private Function BinarySearch(x As Long, ByVal startPos As Long) As Long
    Dim min As Long, mid As Long, max As Long
    min = startPos: max = index - 1
    
    Do While min <= max
        mid = min + (max - min) \ 2
        If DATA(mid) = x Then
            BinarySearch = mid
            Exit Function
        ElseIf isDescending Then
            If DATA(mid) < x Then
                max = mid - 1
            Else
                min = mid + 1
            End If
        Else
            If DATA(mid) > x Then
                max = mid - 1
            Else
                min = mid + 1
            End If
        End If
    Loop
    BinarySearch = -1
End Function

' Sort the data (ascending by default, descending if flagged true)
Public Sub Sort(Optional ByVal Descending As Boolean = False)
    If index < 2 Then Exit Sub
    isDescending = Descending
    QuickSort 0, index - 1
    isSorted = True
End Sub

' Quick Sort O(n log n) on average
Private Sub QuickSort(ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, m As Long
   
    Do While low < high
        i = low: j = high
        m = DATA(low + (high - low) \ 2)
        
        Do While i <= j
            If isDescending Then
                Do While DATA(i) > m: i = i + 1: Loop
                Do While DATA(j) < m: j = j - 1: Loop
            Else
                Do While DATA(i) < m: i = i + 1: Loop
                Do While DATA(j) > m: j = j - 1: Loop
            End If
        
            If i <= j Then
                Swap DATA(i), DATA(j)
                i = i + 1: j = j - 1
            End If
        Loop
        
        ' Tail recursion optimization
        If (j - low) < (high - i) Then
            If low < j Then QuickSort low, j
            low = i ' Continue loop for larger partition
        Else
            If i < high Then QuickSort i, high
            high = j ' Continue loop for larger partition
        End If
    Loop
End Sub

' Helper for clean value swap called by QuickSort
Private Sub Swap(ByRef a As Long, ByRef b As Long)
    Dim t As Long
    t = a: a = b: b = t
End Sub

' Class Properties

' Returns the size of the array
Public Property Get Size() As Long
    Size = index
End Property

' Returns the last index of the array
Public Property Get LastIndex() As Long
    LastIndex = index - 1
End Property

' Get the available capacity of the underlying array
Public Property Get Capacity() As Long
    Capacity = UBound(DATA) + 1 ' UBound is 0-based, capacity is 1-based
End Property

' Return an array containing data only
Public Property Get LongArray() As Long()
    LongArray = ToArray()
End Property

' Return the entire array, including empty slots
Public Property Get DataArray() As Long()
    DataArray = DATA
End Property

Public Property Get Sorted() As Boolean
    Sorted = isSorted
End Property

Public Property Get Descending() As Boolean
    Descending = isDescending
End Property

' Return an element from the specified index location
Public Property Get Item(i As Long) As Long
    If i >= 0 And i < index Then
        Item = DATA(i)
    Else
        Err.Raise ERR_INDEX_OUT_OF_RANGE, "ArrayListLong", "Index out of bounds"
    End If
End Property

' Allows user to set the capacity of the array, should the user know the size of the data set
' Specifically. Best used when after the ArrayListLong is created.
Public Property Let Capacity(newCapacity As Long)
    If newCapacity < index Then Err.Raise ERR_CAPACITY_TOO_SMALL, "ArrayListLong", "Capacity cannot be less than current size"
    If newCapacity < 1 Then Err.Raise ERR_CAPACITY_INVALID, "ArrayListLong", "Capacity must be at least 1"
    ReDim Preserve DATA(newCapacity - 1) ' Adjust to 0-based index
End Property
