Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Configuration
Imports System.Net.Mail
Imports Atom.Core.Logging

''' <summary>
''' メールを送信します。SMTP の設定は Web.config の mailSettings から取得します。
''' </summary>
Public Class MailSender

#Region "定数"
    ''' <summary>ログの発生元名。</summary>
    Private Const LogSource As String = "MailSender"

    ''' <summary>差出人アドレスの設定キー。</summary>
    Private Const FromAddressKey As String = "Mail.FromAddress"
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' メールを送信します。
    ''' </summary>
    ''' <param name="toAddresses">宛先アドレスの一覧。</param>
    ''' <param name="subject">件名。</param>
    ''' <param name="body">本文。</param>
    ''' <param name="loginCode">操作している利用者のログインコード。</param>
    Public Sub Send(toAddresses As IEnumerable(Of String), subject As String, body As String, loginCode As String)
        Using message As New MailMessage()
            message.From = New MailAddress(GetFromAddress())
            message.Subject = subject
            message.Body = body
            message.IsBodyHtml = False

            For Each address As String In toAddresses
                message.To.Add(address)
            Next

            Try
                Using client As New SmtpClient()
                    client.Send(message)
                End Using
                Logger.WriteInformation(LogSource, "メールを送信しました。件名 " & subject, loginCode)
            Catch ex As SmtpException
                Logger.WriteError(LogSource, "メールの送信に失敗しました。件名 " & subject, ex, loginCode)
                Throw
            End Try
        End Using
    End Sub
#End Region

#Region "内部処理"
    ''' <summary>差出人アドレスを取得します。</summary>
    ''' <returns>差出人アドレス。</returns>
    Private Shared Function GetFromAddress() As String
        Dim address As String = ConfigurationManager.AppSettings(FromAddressKey)
        ' 設定が無いとき
        If String.IsNullOrEmpty(address) Then
            Throw New ConfigurationErrorsException("設定 " & FromAddressKey & " が指定されていません。")
        End If
        Return address
    End Function
#End Region

End Class
