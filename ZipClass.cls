VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZipClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type typCenteralDirEnd
   EndOFCentralDirSignature As Long
   NumberOfThisDisk As Integer
   NumberOfDiskWithCentralDir As Integer
   EntriesInTheCentralDirThisOnDisk As Integer
   EntriesInTheCentralDir As Integer
   SizeOfCentralDir As Long
   OffSetOfCentralDir As Long
   ZipFileCommentLength As Integer
End Type

Private ZipFileComment As String
Private CentralDirEnd As typCenteralDirEnd
Private colFiles As Collection

Public Property Get FileCount() As Long
   FileCount = colFiles.Count
End Property

Public Property Let Comment(s As String)
   ZipFileComment = s
   CentralDirEnd.ZipFileCommentLength = Len(s)
End Property

Public Function AddFile(FilePath As String, Optional FileName As String, Optional Comment As String) As Boolean
   Dim f As ZipFile
   
   If Len(Dir(FilePath)) = 0 Then
      Exit Function
   End If
   
   Set f = New ZipFile
   f.FilePath = FilePath
   If (Len(FileName) > 0) Then f.NewFileName = FileName
   If (Len(Comment) > 0) Then f.Comment = Comment
   
   colFiles.Add f
   AddFile = True
End Function

Public Function WriteZip(FilePath As String, Optional Overwrite As Boolean) As Boolean
   Dim f As ZipFile
   Dim l As Long
   Dim fh As Long
   
   If Dir(FilePath) <> "" Then
      If Overwrite Then
         Kill FilePath
      Else
         Exit Function
      End If
   End If
   
   fh = FreeFile
   Open FilePath For Binary As #fh
   For Each f In colFiles
      f.WriteLocalFileHeader fh
   Next
   l = Seek(fh)
   For Each f In colFiles
      f.WriteCentralFileHeader fh
   Next
   With CentralDirEnd
      .EndOFCentralDirSignature = &H6054B50
      .EntriesInTheCentralDirThisOnDisk = colFiles.Count
      .EntriesInTheCentralDir = colFiles.Count
      .SizeOfCentralDir = Seek(fh) - l
      .OffSetOfCentralDir = l - 1
   End With
   Put #fh, , CentralDirEnd
   Put #fh, , ZipFileComment
   Close #fh
   
   WriteZip = True
   
End Function

Private Sub Class_Initialize()
   Set colFiles = New Collection
End Sub

Private Sub Class_Terminate()
   Set colFiles = Nothing
End Sub
