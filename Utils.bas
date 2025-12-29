Option Explicit

' フォルダの存在確認と作成
Public Function GetOrCreateFolder(ByVal path As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
    GetOrCreateFolder = path
End Function

' 重複回避ロジック付きファイル移動
Public Sub SafeMoveFile(ByVal srcPath As String, ByVal destDir As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileName As String: fileName = fso.GetFileName(srcPath)
    Dim baseName As String: baseName = fso.GetBaseName(fileName)
    Dim ext As String: ext = fso.GetExtensionName(fileName)
    Dim targetPath As String: targetPath = destDir & "\" & fileName
    
    ' 1. 重複なし
    If Not fso.FileExists(targetPath) Then
        fso.MoveFile srcPath, targetPath
        Exit Sub
    End If
    
    ' 2. 日付付与 (_yyyymmdd)
    Dim dateStr As String: dateStr = "_" & Format(Date, "yyyymmdd")
    targetPath = destDir & "\" & baseName & dateStr & "." & ext
    If Not fso.FileExists(targetPath) Then
        fso.MoveFile srcPath, targetPath
        Exit Sub
    End If
    
    ' 3. 連番付与 (2)～(99)
    Dim n As Integer, success As Boolean: success = False
    For n = 2 To 99
        targetPath = destDir & "\" & baseName & dateStr & "(" & n & ")." & ext
        If Not fso.FileExists(targetPath) Then
            fso.MoveFile srcPath, targetPath
            success = True
            Exit For
        End If
    Next n
    
    ' 4. それでもダメならエラー報告
    If Not success Then
        If ErrorManager.ReportError(errCollisionLimitReached, fileName) Then
            End ' 全体停止
        End If
    End If
End Sub
