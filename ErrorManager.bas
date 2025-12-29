Option Explicit

' エラーコードの定義
Public Enum AppErrorCode
    errNone = 0
    errCollisionLimitReached = 5001 ' 重複100回超え
    errTotalLimitExceeded = 9999   ' 合計エラー10回超え
End Enum

Private pErrorCount As Integer

Public Sub ResetErrorCount()
    pErrorCount = 0
End Sub

' エラーを記録し、継続不能ならTrueを返す
Public Function ReportError(ByVal code As AppErrorCode, ByVal detail As String) As Boolean
    pErrorCount = pErrorCount + 1
    
    If code = errCollisionLimitReached Then
        MsgBox "リネーム上限(100回)に達したためスキップします: " & detail, vbExclamation
    End If
    
    ' 合計10回エラーが出たら中断フラグを立てる
    If pErrorCount >= 10 Then
        MsgBox "累積エラーが10回に達したため、処理を中断します。", vbCritical
        ReportError = True
    Else
        ReportError = False
    End If
End Function
