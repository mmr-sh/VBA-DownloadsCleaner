Option Explicit
Implements IFileAction

Private Sub IFileAction_Execute(ByVal dto As FileDTO)
    ' すべて画像フォルダへ
    Dim targetDir As String: targetDir = Utils.GetOrCreateFolder(dto.DestDesktopPath & "\画像")
    Utils.SafeMoveFile dto.FilePath, targetDir
End Sub
