Option Explicit
Implements IFileAction

Private Sub IFileAction_Execute(ByVal dto As FileDTO)
    Dim targetDir As String
    ' 10日以上前なら削除予定、そうでなければその他ファイル
    If dto.LastModified < Date - 10 Then
        targetDir = Utils.GetOrCreateFolder(dto.DestDesktopPath & "\削除予定")
    Else
        targetDir = Utils.GetOrCreateFolder(dto.DestDesktopPath & "\その他ファイル")
    End If
    Utils.SafeMoveFile dto.FilePath, targetDir
End Sub
