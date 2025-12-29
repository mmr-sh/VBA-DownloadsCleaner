Option Explicit

' メンバ変数はPrivateにして隠蔽する
Private pFileName As String
Private pFilePath As String
Private pExtension As String
Private pFileType As String
Private pLastModified As Date
Private pDestDesktopPath As String
Private pIsValid As Boolean ' バリデーション結果

' --- プロパティ(読み取り専用) ---

Public Property Get FileName() As String
    FileName = pFileName
End Property

Public Property Get FilePath() As String
    FilePath = pFilePath
End Property

Public Property Get Extension() As String
    Extension = pExtension
End Property

Public Property Get FileType() As String
    FileType = pFileType
End Property

Public Property Get LastModified() As Date
    LastModified = pLastModified
End Property

Public Property Get DestDesktopPath() As String
    DestDesktopPath = pDestDesktopPath
End Property

Public Property Get IsValid() As Boolean
    IsValid = pIsValid
End Property

' --- 初期化とバリデーション ---

' FSOのFileオブジェクトと宛先パスを受け取り、自身を初期化する
Public Sub Load(ByVal fsoFile As Object, ByVal destPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    pFileName = fsoFile.Name
    pFilePath = fsoFile.Path
    
    ' 【バリデーションロジック】
    ' 例: Officeの一時ファイル(~$で始まる)は処理対象外とする
    If Left(pFileName, 2) = "~$" Then
        pIsValid = False
        Exit Sub
    End If
    
    ' 拡張子をここで取得し、小文字に統一しておく（正規化）
    pExtension = LCase(fso.GetExtensionName(pFilePath))
    
    pFileType = fsoFile.Type
    pLastModified = fsoFile.DateLastModified
    pDestDesktopPath = destPath
    
    ' ここまで到達できれば有効なデータとする
    pIsValid = True
End Sub
