VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------
' Module Name: TextFile
' Purpose    : Manage all types of text based files (read, write, append, open)
' Author     : Burton Colter
' Created    : 2025-06-14 11:08:45 AM
' Notes      : Should manage all text file types. File ext will be needed for loading
'            : writing and appending. Data can be returned as a string, array or 2d array
'            : When reading, the class should ignore blank lines and trim all rows of text
'-----------------------------------------------------------------------------------------

Private fso As Object
Private fStream As Object
Private Const textEditor As String = "C:\Windows\notepad.exe " ' Adjust as needed for desired text editor

'Constructor
Private Sub Class_Initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

'Destructor
Private Sub Class_Terminate()
    If Not fso Is Nothing Then Set fso = Nothing
    If Not fStream Is Nothing Then Set fStream = Nothing
End Sub

'Initialize fStream for writing a new text file (destructive)
Private Sub CreateNewFile(fp As String)
    On Error GoTo ErrHandler
    ' Check if the file system folder exists, if not create it
    Dim fPath As String
    fPath = fso.GetParentFolderName(fp)
    If Not fso.FolderExists(fPath) Then fso.CreateFolder fPath
    Set fStream = fso.CreateTextFile(fp, True)
    Exit Sub
    
ErrHandler:
    HandleError "CreateNewFile", Err.Number, Err.Description
    Exit Sub
End Sub

'Initalize fStream for appending an existing text file
Private Sub AppendExistingFile(fp As String)
    On Error GoTo ErrHandler
    Set fStream = fso.OpenTextFile(fp, 8, True)
    Exit Sub
    
ErrHandler:
    HandleError "AppendExistingFile", Err.Number, Err.Description
    Exit Sub
End Sub

'Subroutine to open a file, called by other subroutines
Private Sub OpenExistingFile(fp As String)
    On Error GoTo ErrHandler
    Set fStream = fso.OpenTextFile(fp, 1)
    Exit Sub
    
ErrHandler:
    HandleError "OpenExistingFile", Err.Number, Err.Description
    Exit Sub
End Sub

'Writes to a new file from a string (destructive)
Public Sub NewFileLine(fp As String, DATA As String)
    On Error GoTo ErrHandler
    CreateNewFile (fp)
    fStream.Write (DATA)
    fStream.Close
    Exit Sub
    
ErrHandler:
    HandleError "NewFileLine", Err.Number, Err.Description
    Exit Sub
End Sub

'Writes to a new file from a string array (destructive)
Public Sub NewFileArray(fp As String, DATA() As String)
    On Error GoTo ErrHandler
    CreateNewFile (fp)
    
    Dim Size As Long, i As Long
    
    Size = UBound(DATA)
    For i = 0 To Size - 1: fStream.WriteLine (DATA(i)): Next
    fStream.Write DATA(Size)
    fStream.Close
    Exit Sub
    
ErrHandler:
    HandleError "NewFileArray", Err.Number, Err.Description
    Exit Sub
End Sub

'Appends a line to an existing file with a string
Public Sub AppendFileLine(fp As String, DATA As String)
    On Error GoTo ErrHandler
    
    'Read the end of the File
    Dim tsCheck As Object ' This can be avoided using a reference to
                          ' Microsoft Scripting Runtime (Tools > References)
    Set tsCheck = CreateObject("Scripting.FileSystemObject").OpenTextFile(fp, 1)
    Dim fileText As String
    If fso.FileExists(fp) Then
        Set tsCheck = fso.OpenTextFile(fp, 1) ' 1 = ForReading
        fileText = tsCheck.ReadAll
        tsCheck.Close
    End If
    AppendExistingFile (fp)
    
    'Add a new line if it does not exist
    If Len(fileText) > 0 Then
        If right(fileText, 2) <> vbCrLf Then fStream.Write vbCrLf
    End If
    fStream.Write (DATA)
    fStream.Close
    Exit Sub
    
ErrHandler:
    HandleError "AppendFileLine", Err.Number, Err.Description
    Exit Sub
End Sub

'Appends to an existing file with an string array
Public Sub AppendFileArray(fp As String, DATA() As String)
    On Error GoTo ErrHandler
    AppendExistingFile (fp)
    
    Dim i As Long
    For i = 0 To UBound(DATA): fStream.WriteLine (DATA(i)): Next
    fStream.Close
    Exit Sub
    
ErrHandler:
    HandleError "AppendFileArray", Err.Number, Err.Description
    Exit Sub
End Sub

'Reads data from a file and returns a single string variable
Public Function ReadFileStr(fp As String) As String
    On Error GoTo ErrHandler
    Dim DATA As String
    
    OpenExistingFile (fp)
    DATA = fStream.ReadAll
    fStream.Close
    ReadFileStr = DATA
    Exit Function
    
ErrHandler:
    HandleError "ReadFileStr", Err.Number, Err.Description
    Exit Function
End Function

'Reads data from a file and returns a string array, each line is an entry in the array
Public Function ReadFileArray(fp As String) As String()
    On Error GoTo ErrHandler
    Dim DATA As String
    Dim temp() As String
    Dim rowIndex As Long
    
    OpenExistingFile (fp)
    DATA = fStream.ReadAll
    fStream.Close
    temp = Split(DATA, vbCrLf)
    'Check for blank line at the end of the data and remove if found
    rowIndex = UBound(temp)
    If temp(rowIndex) = "" Then ReDim Preserve temp(rowIndex - 1)
    ReadFileArray = temp
    Exit Function
    
ErrHandler:
    HandleError "ReadFileArray", Err.Number, Err.Description
    Exit Function
End Function

'Reads data from delimited file and returns an array
Public Function ReadFileDelimArray(fp As String, delim As String) As String()
    On Error GoTo ErrHandler
    Dim DATA As String
    Dim temp() As String
    Dim rows() As String
    Dim row() As String
    Dim rowIndex As Long, colIndex As Long
    
    OpenExistingFile (fp)
    DATA = fStream.ReadAll
    rows = Split(DATA, vbCrLf)
    rowIndex = UBound(rows)
    'Check for a blank line at the end of the data and ignore if found
    If rows(rowIndex) = "" Then rowIndex = rowIndex - 1
    
    'Loop through the data line by line and check the length
    'This is to adjust for a ragged array
    Dim i As Long, j As Long
    For i = 0 To rowIndex
        j = UBound(Split(rows(i), delim))
        If colIndex < j Then colIndex = j
    Next
    
    'Resize the array and load
    ReDim temp(rowIndex, colIndex)
    For i = 0 To rowIndex
        row = Split(rows(i), delim)
        'Prevent index out of bounds errors due to ragged data
        For j = 0 To colIndex
            If j <= UBound(row) Then temp(i, j) = row(j)
        Next
    Next
    
    fStream.Close
    ReadFileDelimArray = temp
    Exit Function
    
ErrHandler:
    HandleError "ReadFileDelimArray", Err.Number, Err.Description
    Exit Function
End Function

'Shell subroutine to open file in text editor
Public Sub OpenTextFile(filePath As String)
    On Error GoTo ErrHandler
    Call Shell(textEditor & filePath, vbNormalFocus)
    Exit Sub
    
ErrHandler:
    HandleError "OpenTextFile", Err.Number, Err.Description
    Exit Sub
End Sub

'Error Handling
Private Sub HandleError(moduleName As String, ErrorNumber As Long, ErrorDescription As String)
    Dim ErrorMsg As String
    ErrorMsg = "Sub Routine: " & moduleName & " - Error " & ErrorNumber & ": " & ErrorDescription
    
    'Clean up to prevent memory problems and advise user
    Set fStream = Nothing
    Err.Raise vbObjectError + 69, "File Error", ErrorMsg
End Sub
