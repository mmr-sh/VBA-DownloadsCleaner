Option Explicit

' 条件に応じて適切なクラスインスタンスを生成する
Public Function GetAction(ByVal dto As FileDTO) As IFileAction
    ' DTO側で既に小文字化(正規化)されているため、そのまま比較する
    Select Case dto.Extension
        Case "zip": Set GetAction = New ZipAction
        Case "pdf": Set GetAction = New PdfAction
        Case "jpg", "jpeg", "png", "gif", "bmp": Set GetAction = New ImageAction
        Case "mp3": Set GetAction = New MusicAction
        Case Else: Set GetAction = New OtherAction
    End Select
End Function

