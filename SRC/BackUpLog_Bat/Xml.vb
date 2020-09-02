Public Class SetXml
    Public Property bklogpath As String          'バックアップログファイル格納先パス
    Public Property errmsg As String             'ログファイル内エラーメッセージ文字列
    Public Property normalwebhook As String = "" '正常時 Web hook Url
    Public Property errwebhook As String = ""    '異常時 Web hook Url
End Class
