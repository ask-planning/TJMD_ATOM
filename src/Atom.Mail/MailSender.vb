\xEF\xBB\xBFImports System.Net.Mail

''' <summary>メール送信（現行 BASP21 + 共有フォルダExcel出力 の置換）。</summary>
Public Class MailSender

#Region "フィールド"
    Private ReadOnly _smtpHost As String
    Private ReadOnly _smtpPort As Integer
#End Region

#Region "コンストラクター"
    ''' <summary>SMTP接続先を指定して初期化する。</summary>
    Public Sub New(smtpHost As String, smtpPort As Integer)
        Me._smtpHost = smtpHost
        Me._smtpPort = smtpPort
    End Sub
#End Region

#Region "公開メソッド"
    ''' <summary>メールを送信する。</summary>
    Public Sub Send(fromAddr As String, toAddrs As String, subject As String, body As String, Optional attachmentPath As String = Nothing)
        Using message As New MailMessage()
            message.From = New MailAddress(fromAddr)
            ' 宛先（;区切り）を1件ずつ追加する
            For Each address As String In toAddrs.Split(";"c)
                If address.Trim().Length > 0 Then
                    message.To.Add(address.Trim())
                End If
            Next
            message.Subject = subject
            message.Body = body
            ' 添付があるとき：添付を付ける
            If Not String.IsNullOrEmpty(attachmentPath) Then
                message.Attachments.Add(New Attachment(attachmentPath))
            End If
            ' SMTPで送信する
            Using client As New SmtpClient(Me._smtpHost, Me._smtpPort)
                client.Send(message)
            End Using
        End Using
    End Sub
#End Region

End Class
