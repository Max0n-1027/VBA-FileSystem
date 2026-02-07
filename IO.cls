VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleDescription "FileSystemObject Wrapper Class"
'@Folder("RyoYoshizawa.FileSystem")
'@PredeclaredId
Option Explicit

Private Function fs() As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
End Function

'@Description "Appends a name to an existing path."
Public Function BuildPath(ByVal Path As String, ByVal Name As String) As String
    BuildPath = fs.BuildPath(Path, Name)
End Function

'@Description "Copies one or more directories from one location to another."
Public Function CopyDirectory(ByVal Source As String, ByVal Destination As String, Optional OverWriteFiles As Boolean = False) As Boolean
    On Error GoTo Catch
    fs.CopyFolder Source, Destination, OverWriteFiles
    CopyDirectory = True
    Exit Function
Catch:
    CopyDirectory = False
End Function

'@Description "Copies one or more files from one location to another."
Public Function CopyFile(ByVal Source As String, ByVal Destination As String, Optional OverWriteFiles As Boolean = False) As Boolean
    On Error GoTo Catch
    fs.CopyFile Source, Destination, OverWriteFiles
    CopyFile = True
    Exit Function
Catch:
    CopyFile = False
End Function

'@Description "Creates a new directory."
Public Function CreateDirectory(ByVal DirectoryName As String) As Boolean
    On Error GoTo Catch
    fs.CreateFolder DirectoryName
    CreateDirectory = True
    Exit Function
Catch:
    CreateDirectory = False
End Function

'@Description "Deletes one or more specified directories."
Public Function DeleteDirectory(ByVal DirectorySpec As String, Optional Force As Boolean = False) As Boolean
    On Error GoTo Catch
    fs.DeleteFolder DirectorySpec, Force
    DeleteDirectory = True
    Exit Function
Catch:
    DeleteDirectory = False
End Function

'@Description "Deletes one or more specified files."
Public Function DeleteFile(ByVal FileSpec As String, Optional Force As Boolean = False) As Boolean
    On Error GoTo Catch
    fs.DeleteFolder FileSpec, Force
    DeleteFile = True
    Exit Function
Catch:
    DeleteFile = False
End Function

'@Description "Checks if a specified directory exists."
Public Function ExistsDirectory(ByVal DirectorySpec As String) As Boolean
    ExistsDirectory = fs.FolderExists(DirectorySpec)
End Function

'@Description "Checks if a specified drive exists."
Public Function ExistsDrive(ByVal DirveSpec As String) As Boolean
    ExistsDrive = fs.DriveExists(DirveSpec)
End Function

'@Description "Checks if a specified file exists."
Public Function ExistsFile(ByVal FileSpec As String) As Boolean
    ExistsFile = fs.FileExists(FileSpec)
End Function

'@Description "Returns the complete path from the root of the drive for the specified path."
Public Function GetAbsolutePathName(ByVal Path As String) As String
    GetAbsolutePathName = fs.GetAbsolutePathName(Path)
End Function

'@Description "Returns the base name of a specified file or directory."
Public Function GetBaseName(ByVal Path As String) As String
    GetBaseName = fs.GetBaseName(Path)
End Function

'@Description "Returns a collection of all the file names in a specified directory."
Public Function GetDirectoryFiles(ByVal FolderPath As String) As Collection
    Set GetDirectoryFiles = New Collection
    Dim dir As Object: Set dir = fs.GetFolder(FolderPath)
    Dim file As Object: For Each file In dir.Files
        GetDirectoryFiles.Add file.Path
    Next
End Function

'@Description "Returns the drive name of a specified path."
Public Function GetDirectorySize(ByVal Path As String) As LongLong
    If Not Me.ExistsDirectory(Path) Then
        GetDirectorySize = -1
        Exit Function
    End If
    Dim dir As Object: Set dir = fs.GetFolder(Path)
    GetDirectorySize = dir.Size
End Function

'@Description "Returns the drive name of a specified path."
Public Function GetDriveName(ByVal Path As String) As String
    GetDriveName = fs.GetDriveName(Path)
End Function

'@Description "Returns a collection consisting of all drive names available on the local machine."
Public Function GetDriveNames() As Collection
    Set GetDriveNames = New Collection
    Dim drive As Object: For Each drive In fs.Drives
        GetDriveNames.Add drive.Path
    Next
End Function

'@Description "Returns the file extension name for the last component in a specified path."
Public Function GetExtensionName(ByVal Path As String) As String
    GetExtensionName = fs.GetExtensionName(Path)
End Function

'@Description "Returns the file name or directory name for the last component in a specified path."
Public Function GetFileName(ByVal Path As String) As String
    GetFileName = fs.GetFileName(Path)
End Function

'@Description "Returns the size, in bytes, of a specified file."
Public Function GetFileSize(ByVal Path As String) As LongLong
    If Not Me.ExistsFile(Path) Then
        GetFileSize = -1
        Exit Function
    End If
    Dim file As Object: Set file = fs.GetFile(Path)
    GetFileSize = file.Size
End Function

'@Description "Returns the name of the parent directory of the last component in a specified path."
Public Function GetParentDirectoryName(ByVal Path As String) As String
    GetParentDirectoryName = fs.GetParentFolderName(Path)
End Function

'@Description "Collection of all Directory objects contained within a directory name."
Public Function GetSubDirectories(ByVal Path As String) As Collection
    Set GetSubDirectories = New Collection
    If Not Me.ExistsDirectory(Path) Then
        Exit Function
    End If
    
    Dim dir As Object: Set dir = fs.GetFolder(Path)
    Dim subDir As Object: For Each subDir In dir.SubFolders
        GetSubDirectories.Add subDir.Path
    Next
End Function

'@Description "Returns the temporary folder specified."
Public Function GetTempDirectoryName() As String
    GetTempDirectoryName = fs.GetSpecialFolder(2)
End Function

'@Description "Returns a randomly generated temporary file or folder."
Public Function GetTempName() As String
    GetTempName = fs.GetTempName
End Function

'@Description "Moves one or more directories from one location to another."
Public Function MoveDirectory(ByVal Source As String, ByVal Destination As String) As Boolean
    On Error GoTo Catch
    fs.MoveFolder Source, Destination
    MoveDirectory = True
    Exit Function
Catch:
    MoveDirectory = False
End Function

'@Description "Moves one or more files from one location to another."
Public Function MoveFile(ByVal Source As String, ByVal Destination As String) As Boolean
    On Error GoTo Catch
    fs.MoveFile Source, Destination
    MoveFile = True
    Exit Function
Catch:
    MoveFile = False
End Function
