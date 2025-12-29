Option Explicit
Implements IFileAction

Private Sub IFileAction_Execute(ByVal dto As FileDTO)
    Dim targetDir As String
    ' 30日以上前なら削除予定、そうでなければpdf
    If dto.LastModified < Date - 30 Then
        targetDir = Utils.GetOrCreateFolder(dto.DestDesktopPath & "\削除予定")
    Else
        targetDir = Utils.GetOrCreateFolder(dto.DestDesktopPath & "\pdf")
    End If
    Utils.SafeMoveFile dto.FilePath, targetDir
End Sub
