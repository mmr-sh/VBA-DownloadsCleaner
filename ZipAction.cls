Option Explicit
Implements IFileAction

Private Sub IFileAction_Execute(ByVal dto As FileDTO)
    Dim targetDir As String
    ' 7日以上前なら削除予定、そうでなければzip
    If dto.LastModified < Date - 7 Then
        targetDir = Utils.GetOrCreateFolder(dto.DestDesktopPath & "\削除予定")
    Else
        targetDir = Utils.GetOrCreateFolder(dto.DestDesktopPath & "\zip")
    End If
    Utils.SafeMoveFile dto.FilePath, targetDir
End Sub
