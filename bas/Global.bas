Attribute VB_Name = "Global"
Option Explicit

Private Declare Function PathFileExistsA Lib "shlwapi" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectoryA Lib "shlwapi" (ByVal pszPath As String) As Long


Public Function PathExist(ByVal FileName As String) As Boolean: PathExist = PathFileExistsA(FileName) <> 0: End Function
Public Function PathIsDirectory(ByVal sPath As String) As Boolean: PathIsDirectory = PathIsDirectoryA(sPath): End Function
Public Function GetFilePath(ByVal Path As String) As String
On Error Resume Next
    GetFilePath = Left(Path, Len(Path) - Len(Right(Path, Len(Path) - InStrRev(Path, "\"))))
End Function
Function GetFileTitle(Path As String, Optional ByVal IncludeExt As Boolean) As String
Dim Ext    As Long
    GetFileTitle = Right(Path, Len(Path) - InStrRev(Path, "\"))
    If Not IncludeExt Then Ext = InStr(1, GetFileTitle, ".")
    If Ext <> 0 Then GetFileTitle = Left(GetFileTitle, Ext - 1)
End Function

Function GetFileExt(Path As String) As String
    GetFileExt = Right(Path, Len(Path) - InStrRev(Path, "."))
End Function



Public Sub ImportFromIconFile(ByVal FileName As String, Coll As Collection)
Dim Import   As cIconEntry
Dim Bmp      As cGDIPBitmap
Dim i        As Long

    Set Coll = New Collection
    Set Import = New cIconEntry
    
    Import.OpenIconFile FileName
    For i = 0 To Import.IconCount - 1
        Set Bmp = Import.GdipBitmap(i)
        Coll.Add Bmp
    Next
    Set Import = Nothing
    
End Sub

Public Sub ExportToIconFile(ByVal FileName As String, Coll As Collection)
Dim Export   As cIconEntry
Dim Bmp      As cGDIPBitmap
Dim i        As Long

    If Coll Is Nothing Then Exit Sub
    Set Export = New cIconEntry
    For Each Bmp In Coll
        Export.Add Bmp
    Next
    Export.SaveIconFile FileName
    Set Export = Nothing
    
    If Err.Number Then Debug.Print Err.Description
End Sub

Public Function ExportBitmapToIcon(mBmp As cGDIPBitmap, ByVal FileName As String) As Boolean
Dim Export As cIconEntry

    Set Export = New cIconEntry
    If Export.Add(mBmp) Then
        ExportBitmapToIcon = Export.SaveIconFile(FileName)
    End If
    Set Export = Nothing
End Function


Public Function PatchIconFile(ByVal FileName As String) As Boolean
Dim Patch As cIconEntry
    Set Patch = New cIconEntry
    If Patch.OpenIconFile(FileName) Then
        Call Patch.Truncate
        PatchIconFile = Patch.OpenIconFile(FileName)
    End If
    Set Patch = Nothing
End Function

Public Function GdipBitmap_(Optional Source As Variant) As cGDIPBitmap
On Error GoTo e
    Set GdipBitmap_ = New cGDIPBitmap
    GdipBitmap_.LoadImage Source
e:
End Function
