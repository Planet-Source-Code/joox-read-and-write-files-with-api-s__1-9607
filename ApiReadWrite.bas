Attribute VB_Name = "ApiWriteRead"
Public Declare Function ReadFileNO Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CreateFileNS Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WriteFileNO Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long


Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_ALWAYS = 2
Public Const CREATE_NEW = 1
Public Const OPEN_ALWAYS = 4
Public Const OPEN_EXISTING = 3
Public Const TRUNCATE_EXISTING = 5
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Public Const FILE_FLAG_NO_BUFFERING = &H20000000
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000

Public Function ReadFile(File As String) As String
Dim filesizelow As Long
Dim filesizehigh As Long
Dim longbuffer As Long
Dim stringbuffer As String
Dim numread As Long
Dim hFile As Long
Dim retval As Long

hFile = CreateFileNS(File, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    'get an handle for the file
If hFile = -1 Then
    'there is an error! maybe the file doesn't exist
    ReadFile = "-1"
    Exit Function
End If

filesizelow = GetFileSize(hFile, filesizehigh)
    'get the size of the file
stringbuffer = Space(filesizelow)
    'file a string with spaces
retval = ReadFileNO(hFile, ByVal stringbuffer, filesizelow, numread, 0)
    'get the whole data of the file
If numread = 1 Then
    'if numread is 1 then everything is ok!
   ReadFile = stringbuffer
Else
    ReadFile = -1
End If

retval = CloseHandle(hFile)
    'important: close filehandle!
End Function

Public Function WriteFile(File As String, Data As String) As Long
Dim hFile As Long
Dim retval As Long
Dim numwritten As Long

hFile = CreateFileNS(File, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
    'get an handle for the file
If hFile = -1 Then
    'there is an error! maybe the file doesn't exist
    WriteFile = "-1"
    Exit Function
End If

retval = WriteFileNO(hFile, ByVal Data, Len(Data), numwritten, 0)
    'write the data
WriteFile = numwritten
    'returns the number of written bytes
CloseHandle hFile
    'important: close filehandle!
End Function
