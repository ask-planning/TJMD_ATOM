Option Strict On
Option Explicit On
Option Infer Off

Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient

''' <summary>
''' データベース接続の生成を担います。接続文字列は Web.config の connectionStrings から取得します。
''' 接続オブジェクトは呼び出し側が Using で破棄してください。
''' </summary>
Public NotInheritable Class Database

#Region "定数"
    ''' <summary>接続文字列の名前。</summary>
    Public Const ConnectionName As String = "AtomDb"
#End Region

#Region "コンストラクタ"
    ''' <summary>インスタンス化を禁止します。</summary>
    Private Sub New()
    End Sub
#End Region

#Region "プロパティ"
    ''' <summary>
    ''' いま繋ぎに行くデータベース名です。
    ''' </summary>
    ''' <remarks>
    ''' 画面上部の環境の帯に「接続先」として出すために使います。
    ''' **接続文字列そのものを画面へ出してはいけません。**ここではデータベース名だけを返します。
    ''' 設定が無いときや解析できないときは空文字を返します（表示のためだけなので落としません）。
    ''' </remarks>
    Public Shared ReadOnly Property CurrentDatabaseName As String
        Get
            Try
                Return Atom.Core.Configuration.AppEnvironment.GetDatabaseName(GetConnectionString())
            Catch ex As ConfigurationErrorsException
                ' 接続文字列が未設定のとき
                Return String.Empty
            End Try
        End Get
    End Property
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' 接続を生成して開きます。
    ''' </summary>
    ''' <returns>開いた状態の接続。呼び出し側で破棄してください。</returns>
    Public Shared Function CreateOpenConnection() As SqlConnection
        Dim connection As New SqlConnection(GetConnectionString())
        Try
            connection.Open()
        Catch
            connection.Dispose()
            Throw
        End Try
        Return connection
    End Function

    ''' <summary>
    ''' 接続できるかを確認します。フッターの再接続処理から呼び出します。
    ''' </summary>
    ''' <returns>接続できたとき True。</returns>
    Public Shared Function TestConnection() As Boolean
        Using connection As SqlConnection = CreateOpenConnection()
            Return connection.State = ConnectionState.Open
        End Using
    End Function
#End Region

#Region "内部処理"
    ''' <summary>
    ''' 接続文字列を取得します。
    ''' </summary>
    ''' <returns>接続文字列。</returns>
    Private Shared Function GetConnectionString() As String
        ' 環境ごとの接続（AtomDb@Local など）を優先して取ります
        Return Atom.Core.Configuration.AppEnvironment.GetConnectionString(ConnectionName)
    End Function
#End Region

End Class
