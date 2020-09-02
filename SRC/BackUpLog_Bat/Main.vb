Imports System.Diagnostics
Imports System.IO
Imports System.Console
Imports System.Text
Imports System.Xml.Serialization

Public Class Main
    '*XMLファイル取得情報*****************************************************************
    Private Shared LogFdPath As String() = Nothing          'バックアップサーバログファイルパス
    Private Shared str_ErrVal As String() = Nothing         'バックアップログエラー文字列
    Private Shared str_WebhookNormal As String = ""         '正常時WebhookURL
    Private Shared str_WebhookErr As String = ""            '異常時WebhookURL
    '*************************************************************************************

    'Xmlファイル名
    Const XmlFilePath As String = "Setting.xml"

    'バッチ実行時の日付
    Private Shared Day As String = Now.ToString("yyyy/MM/dd", New Globalization.CultureInfo("ja-JP"))

    '曜日別チェックファイル番号
    Private Shared hshtable As Hashtable

    <STAThread()> Shared Sub Main()
        Try

            If InitMain() = False Then
                'ログチェック異常終了メッセージ送信
                SendMsg(str_WebhookErr, "バックアップログチェック中に処理が異常終了しました。ログを確認して下さい")
                Exit Try
            End If

            'サーババックアップログチェック
            If CheckLog() = False Then
                Exit Try
            End If

            'ログチェック完了メッセージ送信
            SendMsg(str_WebhookNormal, "バックアップログチェックが正常に終了しました。チェック結果をバックアップログチェック記録簿に記入してください。")

        Catch ex As Exception
            'エラーログ出力
            PutErrorLog(ex.ToString)
            'ログチェック異常終了メッセージ送信
            SendMsg(str_WebhookErr, "バックアップログチェック中に処理が異常終了しました。ログを確認して下さい")
        End Try
    End Sub

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <returns>Boolean  : 正常終了/異常終了</returns>
    ''' <remarks></remarks>
    Private Shared Function InitMain() As Boolean
        Dim BolFlg As Boolean = False
        Dim str_ErrMsg As String = ""        'エラーメッセージ文字列
        Dim str_BkFdPath As String = ""      'バックアップサーバログ格納先
        Dim xmlPath As String = ""           'Xmlファイル格納パス
        Dim Result As SetXml = Nothing       'XMLファイル情報

        Try
            '曜日別チェックファイル番号格納
            hshtable = New Hashtable
            hshtable.Add("Monday", "05")
            hshtable.Add("Tuesday", "01")
            hshtable.Add("Wednesday", "02")
            hshtable.Add("Thursday", "03")
            hshtable.Add("Friday", "04")
            hshtable.Add("Saturday", "05")

            'Xmlファイルパスを生成
            xmlPath = My.Application.Info.DirectoryPath
            xmlPath = System.IO.Path.Combine(System.IO.Directory.GetParent(xmlPath).ToString(), "XML", XmlFilePath)

            'XMLファイル存在チェック
            If Not System.IO.File.Exists(xmlPath) Then
                PutErrorLog("InitMain", "XMLファイルが存在しません。" & "指定ファイルパス：" & xmlPath)
                Exit Try
            End If

            'Xmlから情報を取得
            Result = GetXml(xmlPath)

            'XMLから取得したデータを各変数に格納
            str_BkFdPath = Result.bklogpath          'サーバログチェックファイル格納パス
            LogFdPath = str_BkFdPath.Split(",")

            str_WebhookNormal = Result.normalwebhook '正常時Webhook URL
            str_WebhookErr = Result.errwebhook       '異常時Webhook URL

            str_ErrMsg = Result.errmsg               'ログファイルのエラーメッセージ文字列
            str_ErrVal = str_ErrMsg.Split(",")

            BolFlg = True

        Catch ex As Exception
            'エラーログ出力
            PutErrorLog(ex.ToString)
        End Try
        Return BolFlg
    End Function

    ''' <summary>
    ''' サーバログチェック
    ''' </summary>
    ''' <returns> Boolean  : 正常終了/異常終了</returns>
    ''' <remarks></remarks>
    Private Shared Function CheckLog() As Boolean
        Dim FileNm As String = ""                  'ログチェックファイル名
        Dim SplitValFilePath As String() = Nothing 'ファイルパス格納用
        Dim BolFlg As Boolean = False
        Dim BolErrFlg As Boolean = False

        Try
            'ログチェック対象ファイル名を取得
            If GetCheckFileNm(FileNm, BolErrFlg) = False Then
                Exit Try
            End If

            'ログファイル件数が0件の場合、後続処理をしない
            If String.IsNullOrEmpty(FileNm) Then
                Exit Try
            End If

            '取得したファイルパス文字列をカンマ区切りで分割
            SplitValFilePath = FileNm.Split(",")
            For i As Integer = 0 To SplitValFilePath.Length - 1
                'ログファイル存在チェック
                If ChkFileExists(SplitValFilePath(i)) = False Then
                    SendMsg(str_WebhookErr, "ログファイルが存在しません。基盤委員会のメンバーに報告して下さい。" & vbCrLf & "エラーファイルパス：" & SplitValFilePath(i))
                    BolErrFlg = True
                    Continue For
                End If

                'ログファイル空ファイルチェック
                Dim fi As New System.IO.FileInfo(SplitValFilePath(i))
                If fi.Length <= 0 Then
                    SendMsg(str_WebhookErr, "ログファイルが空で出力されています。基盤委員会のメンバーに報告して下さい。" & vbCrLf & "指定ファイルパス：" & SplitValFilePath(i))
                    PutErrorLog("ログファイルが空で出力されています。基盤委員会のメンバーに報告して下さい。" & vbCrLf & "指定ファイルパス：" & SplitValFilePath(i))
                    BolErrFlg = True
                    Continue For
                End If

                'ログファイル内容チェック
                If ReadLogFile(SplitValFilePath(i)) = False Then
                    SendMsg(str_WebhookErr, "ログファイルにエラーが出力されています。基盤委員会のメンバーに報告して下さい。" & vbCrLf & "指定ファイルパス：" & SplitValFilePath(i))
                    PutErrorLog("ログファイルにエラーが出力されています。基盤委員会のメンバーに報告して下さい。" & vbCrLf & "指定ファイルパス：" & SplitValFilePath(i))
                    BolErrFlg = True
                End If
            Next

            'エラーチェック該当なしの場合
            If BolErrFlg = False Then
                BolFlg = True
            End If

        Catch ex As Exception
            'エラーログ出力
            PutErrorLog("CheckLog", ex.ToString)
        End Try
        Return BolFlg
    End Function

    ''' <summary>
    ''' ログチェック対象ファイル名を取得する
    ''' </summary>
    ''' <param name="FileNm">ファイルパス格納文字列</param>
    ''' <returns>Boolean  : 正常終了/異常終了</returns>
    ''' <remarks></remarks>
    Private Shared Function GetCheckFileNm(ByRef FileNm As String, ByRef ErrFlg As Boolean) As Boolean
        Dim BolFlg As Boolean = False
        Dim BkFIleChkFlg As Boolean = False 'バックアップログチェックフラグ
        Dim strFileNm As String() = Nothing
        Try

            Day = Day.Replace("/", "")
            Dim theDay = New DateTime(Day.Substring(0, 4), Day.Substring(4, 2), Day.Substring(6, 2))
            Dim culture = System.Globalization.CultureInfo.GetCultureInfo("en-US")
            Dim TheWeek As String = theDay.ToString("dddd", culture)

            For i As Integer = 0 To LogFdPath.Length - 1

                'フォルダ存在チェック
                If ChkFoldExists(LogFdPath(i)) = False Then
                    PutErrorLog("GetFileName", "指定したバックアップログ格納フォルダが存在しません。" & vbCrLf & "指定フォルダパス：" & LogFdPath(i))
                    SendMsg(str_WebhookErr, "指定したバックアップログ格納フォルダが存在しません。" & vbCrLf & "指定フォルダパス：" & LogFdPath(i))
                    ErrFlg = True
                End If

                strFileNm = GetFileName(LogFdPath(i), hshtable.Item(TheWeek))

                'バックアップログファイル存在チェック
                If strFileNm.Length = 0 Then
                    PutErrorLog("GetCheckFileNm", "バックアップログファイルの件数が0件です。ファイルが存在するか確認して下さい" & vbCrLf & "指定フォルダパス：" & LogFdPath(i))
                    SendMsg(str_WebhookErr, "バックアップログファイルの件数が0件です。ファイルが存在するか確認して下さい" & vbCrLf & "指定フォルダパス：" & LogFdPath(i))
                    ErrFlg = True
                End If

                For j As Integer = 0 To strFileNm.Length - 1

                    'ファイル名が複数ある場合はカンマで区切る
                    If FileNm.Length <> 0 Then
                        FileNm = FileNm + ","
                    End If

                    FileNm = FileNm + strFileNm(j)
                Next
            Next

            BolFlg = True
        Catch ex As Exception
            'エラー出力
            PutErrorLog("GetCheckFileNm", ex.ToString)
        End Try
        Return BolFlg
    End Function

    ''' <summary>
    ''' 指定したファイルの内容を読み込む
    ''' </summary>
    ''' <param name="FilePath">ファイルパス</param>
    ''' <returns> Boolean  : 正常終了/異常終了</returns>
    ''' <remarks></remarks>
    Private Shared Function ReadLogFile(ByVal FilePath As String) As Boolean
        Dim BolFlg As Boolean = False

        Dim Str_Line As String
        Dim Str_Table As String = ""
        Dim sr As New System.IO.StreamReader(FilePath, System.Text.Encoding.UTF8)
        Try

            While sr.Peek() > -1
                '１行取得
                Str_Line = sr.ReadLine
                '指定文字列を検索
                For i As Integer = 0 To str_ErrVal.Length - 1
                    If Trim(UCase(Str_Line)).Contains(str_ErrVal(i)) = True Then
                        Return BolFlg
                    End If
                Next
            End While
            BolFlg = True
        Catch ex As Exception
            'エラー出力
            PutErrorLog("ReadLogFile", ex.ToString)
        Finally
            'ファイルクローズ
            sr.Close()
            sr.Dispose()
        End Try
        Return BolFlg
    End Function

    ''' <summary>
    ''' 指定したファイルの存在チェック
    ''' </summary>
    ''' <param name="strFilePath">ファイルパス</param>
    ''' <returns>Boolean  : 正常終了/異常終了</returns>
    ''' <remarks></remarks>
    Private Shared Function ChkFileExists(ByVal strFilePath As String) As Boolean
        Dim bolFlg As Boolean = False
        Try
            'ファイル存在チェック
            If Not System.IO.File.Exists(strFilePath) Then
                PutErrorLog("ChkFileExists", "ファイルが存在しません" & "指定フォルダパス：" & strFilePath)
                Return bolFlg
            End If
            bolFlg = True
        Catch ex As Exception
            PutErrorLog("ChkFileExists", ex.ToString)
        End Try
        Return bolFlg
    End Function

    ''' <summary>
    ''' 指定したフォルダの存在チェック
    ''' </summary>
    ''' <param name="strFoldPath">ファイルパス</param>
    ''' <returns> Boolean  : 正常終了/異常終了</returns>
    ''' <remarks></remarks>
    Private Shared Function ChkFoldExists(ByVal strFoldPath As String) As Boolean
        Dim bolFlg As Boolean = False
        Try
            'フォルダ存在チェック
            If Not System.IO.Directory.Exists(strFoldPath) Then
                Exit Try
            End If
            bolFlg = True
        Catch ex As Exception
            PutErrorLog("ChkFoldExists", ex.ToString)
        End Try
        Return bolFlg
    End Function

    ''' <summary>
    ''' 指定したフォルダのファイル名を取得する
    ''' </summary>
    ''' <param name="strFoldPath">対象フォルダパス</param>
    ''' <param name="BeforeVal">検索指定文字</param>
    ''' <returns>取得ファイル名格納配列</returns>
    ''' <remarks></remarks>
    Private Shared Function GetFileName(ByVal strFoldPath As String, Optional ByVal BeforeVal As String = "") As String()
        Dim strArrayWK As String() = Nothing
        Dim strArray As String() = Nothing

        Try

            strArrayWK = System.IO.Directory.GetFiles(strFoldPath, "*" & BeforeVal & "*")

            ReDim strArray(strArrayWK.Length - 1)

            For i As Integer = 0 To strArrayWK.Length - 1
                strArray(i) = strArrayWK(i)
            Next
        Catch ex As Exception
            PutErrorLog("GetFileNamexml", ex.ToString)
        End Try
        Return strArray
    End Function

    ''' <summary>
    ''' 指定した文字列をTeamsへ送信する
    ''' </summary>
    ''' <param name="Webhook">WebhookUrl</param>
    ''' <param name="Msg">送信本文</param>
    ''' <returns>Boolean  : 正常終了/異常終了</returns>
    ''' <remarks></remarks>
    Private Shared Function SendMsg(ByVal Webhook As String, ByVal Msg As String) As Boolean
        Dim BolFlg As Boolean = False
        Try


            'TLS1.1 および 1.2 を有効化
            Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls11 Or Net.SecurityProtocolType.Tls12

            Dim sWebhookURL As String
            Dim objWC As New Net.WebClient
            Dim sData As String
            Dim sText As String

            'Webhook URL設定
            sWebhookURL = Webhook

            '送信したいメッセージを設定
            sText = Msg

            '半角の\がメッセージに含まれているとエラーになるので、全角に変換
            sText = sText.Replace("\", "￥")
            sData = "{""text"":""" & sText & """}"
            objWC.Headers.Add(Net.HttpRequestHeader.ContentType, "application/json;charset=UTF-8")

            '文字エンコーディングをUTF=8に設定
            objWC.Encoding = System.Text.Encoding.UTF8

            'POSTメソッドを使用して、指定したリソースに指定した文字列をアップロード
            'objWC.UploadString(sWebhookURL, sData)

            BolFlg = True
        Catch ex As Exception
            PutErrorLog("SendMsg", ex.ToString)
        End Try
        Return BolFlg
    End Function

    ''' <summary>
    ''' エラー情報を出力する
    ''' </summary>
    ''' <param name="PStr_ErrRes">エラー詳細</param>
    ''' <param name="PStr_ErrAddRes">エラー詳細(追記)</param>
    ''' <returns>Boolean  : 正常終了/異常終了</returns>
    ''' <remarks></remarks>
    Private Shared Function PutErrorLog(ByVal PStr_ErrRes As String, Optional ByVal PStr_ErrAddRes As String = "") As Boolean
        Dim Bool_Rtn As Boolean = False
        Dim Str_FName As String = ""
        Dim Int_FNo As Integer
        Dim GUser_Env = ""
        Dim LogFdPath As String = ""               'ログファイルパス
        Try
            '------------------------
            'ファイルの設定、編集
            '------------------------

            With GUser_Env
                LogFdPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()), "LOG")
                If ChkFoldExists(LogFdPath) = False Then
                    'LOGフォルダ作成
                    Dim di As New System.IO.DirectoryInfo(LogFdPath)
                    di.Create()
                End If

                Str_FName = System.IO.Path.Combine("..\LOG\", Now().ToString("yyyyMMdd") & ".LOG")
            End With

            '--------------------------
            'ファイルをオープン＆書込み
            '--------------------------

            'ファイルをオープン
            Int_FNo = -1
            Int_FNo = FreeFile()
            Call FileOpen(Int_FNo, Str_FName, OpenMode.Append)

            'ログ内容を書き込み
            Call PrintLine(Int_FNo, Now())                                                 ' 発生日時
            Call PrintLine(Int_FNo, Chr(9) & "エラー詳細  = " & PStr_ErrRes)               'エラー内容

            'ログ内容を書き込み(追記が存在すれば)
            If PStr_ErrAddRes <> "" Then
                Call PrintLine(Int_FNo, Chr(9) & "エラー詳細（追記）  = " & PStr_ErrAddRes) 'エラー内容追記分
            End If

            Bool_Rtn = True

        Catch exp As Exception
            '例外処理

        Finally
            'ファイルをクローズ
            If Int_FNo <> -1 Then
                Call FileClose(Int_FNo)
            End If

        End Try

        Return Bool_Rtn

    End Function

End Class
Module Xml
    ''' <summary>
    ''' XMLから情報を取得する
    ''' </summary>
    ''' <remarks></remarks>
    Function GetXml(ByVal xmlFilePath As String)
        'XML作成
        Dim xmlSerializer = New XmlSerializer(GetType(SetXml))
        Dim result As SetXml
        Dim xmlSetting = New System.Xml.XmlReaderSettings() _
                         With {.CheckCharacters = False}
        Dim StreamReader = New StreamReader(xmlFilePath, Encoding.UTF8)
        Dim xmlReader = System.Xml.XmlReader.Create(StreamReader, xmlSetting)
        result = CType(xmlSerializer.Deserialize(xmlReader), SetXml)
        Return result
    End Function

End Module
