Option Explicit
Implements IFileAction

Private Sub IFileAction_Execute(ByVal dto As FileDTO)
    ' すべて音楽フォルダへ
    Dim targetDir As String: targetDir = Utils.GetOrCreateFolder(dto.DestDesktopPath & "\音楽")
    Utils.SafeMoveFile dto.FilePath, targetDir
End Sub
