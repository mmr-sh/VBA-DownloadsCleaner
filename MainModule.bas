Option Explicit

Sub CleanupDownloadsFolder()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dlPath As String: dlPath = GetDownloadsPath(fso)
    Dim dsPath As String: dsPath = GetDesktopPath(fso)
    
    If dlPath = "" Then
        MsgBox "ダウンロードフォルダが見つかりませんでした。処理を終了します。"
        Exit Sub
    End If
    
    Dim factory As New ActionFactory
    Dim dto As FileDTO
    Dim fileObj As Object
    Dim results As Object: Set results = CreateObject("Scripting.Dictionary")
    
    ErrorManager.ResetErrorCount
    
    ' メインループ
    For Each fileObj In fso.GetFolder(dlPath).Files
        Set dto = New FileDTO
        
        ' 変更点: データを個別にセットせず、オブジェクトごと渡してLoadさせる
        dto.Load fileObj, dsPath
        
        ' 変更点: バリデーション結果を確認 (無効なファイルならスキップ)
        If dto.IsValid Then
            Dim action As IFileAction: Set action = factory.GetAction(dto)
            action.Execute dto
            
            ' 集計
            Dim actName As String: actName = TypeName(action)
            results(actName) = results(actName) + 1
        End If
        
    Next fileObj
    
    ' 完了報告
    ShowSummary results
End Sub

' ダウンロードフォルダ取得（Environ優先、FallbackでShell）
Private Function GetDownloadsPath(fso As Object) As String
    Dim path As String: path = Environ("USERPROFILE") & "\Downloads"
    
    If Not fso.FolderExists(path) Then
        Dim desk As String: desk = GetDesktopPath(fso)
        path = fso.GetParentFolderName(desk) & "\Downloads"
    End If
    
    If fso.FolderExists(path) Then GetDownloadsPath = path Else GetDownloadsPath = ""
End Function

' デスクトップパス取得
Private Function GetDesktopPath(fso As Object) As String
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    GetDesktopPath = shell.SpecialFolders("Desktop")
End Function

' 集計表示
Private Sub ShowSummary(dict As Object)
    Dim msg As String, k As Variant
    msg = "清掃が完了しました！移動結果：" & vbCrLf
    
    If dict.Count = 0 Then
        msg = msg & "移動対象のファイルはありませんでした。"
    Else
        For Each k In dict.Keys
            msg = msg & "・" & k & ": " & dict(k) & " 件" & vbCrLf
        Next
    End If
    
    MsgBox msg, vbInformation, "終了"
End Sub
