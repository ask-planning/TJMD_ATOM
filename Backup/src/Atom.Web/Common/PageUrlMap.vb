Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic

Namespace Common

    ''' <summary>
    ''' メニュー番号と画面 URL の対応を保持します。
    ''' WinForms 版の ScreenFactory に相当します。未実装の画面は登録しません。
    ''' </summary>
    Public NotInheritable Class PageUrlMap

#Region "フィールド"
        ''' <summary>メニュー番号ごとの画面 URL。</summary>
        Private Shared ReadOnly _urls As Dictionary(Of Integer, String) = CreateUrls()
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' メニュー番号に対応する画面 URL を取得します。
        ''' </summary>
        ''' <param name="menuNo">メニュー番号。</param>
        ''' <returns>画面の URL。未実装のときは空文字。</returns>
        Public Shared Function GetUrl(menuNo As Integer) As String
            Dim url As String = Nothing
            ' 登録があるとき
            If _urls.TryGetValue(menuNo, url) Then
                Return url
            End If
            ' 登録が無いとき
            Return String.Empty
        End Function

        ''' <summary>
        ''' 画面が実装済みかどうかを返します。
        ''' </summary>
        ''' <param name="menuNo">メニュー番号。</param>
        ''' <returns>実装済みのとき True。</returns>
        Public Shared Function IsImplemented(menuNo As Integer) As Boolean
            Return _urls.ContainsKey(menuNo)
        End Function
#End Region

#Region "内部処理"
        ''' <summary>対応表を作成します。画面を実装したらここに追記します。</summary>
        ''' <returns>メニュー番号と URL の対応表。</returns>
        Private Shared Function CreateUrls() As Dictionary(Of Integer, String)
            Dim urls As New Dictionary(Of Integer, String)()

            ' 上海輸入定期便＆AIR便
            ' urls.Add(101, "~/Pages/Import/InvoiceImport.aspx")
            ' urls.Add(102, "~/Pages/Import/InvoiceConfirm.aspx")
            ' urls.Add(103, "~/Pages/Import/InvoiceHeaderEntry.aspx")
            ' urls.Add(104, "~/Pages/Import/PurchaseOrderImport.aspx")
            ' urls.Add(105, "~/Pages/Import/PurchaseOrderConfirm.aspx")
            ' urls.Add(106, "~/Pages/Import/StockTransfer.aspx")
            ' urls.Add(107, "~/Pages/Import/SapSlipAmountEntry.aspx")

            ' その他輸入関連
            ' urls.Add(201, "~/Pages/Procurement/PurchaseOrderImport.aspx")
            ' urls.Add(202, "~/Pages/Procurement/PurchaseOrderConfirm.aspx")
            ' urls.Add(203, "~/Pages/Procurement/DeliveryRequest.aspx")

            ' 輸入経費入力
            ' urls.Add(301, "~/Pages/Expense/ExpenseEntry.aspx")

            ' マスタ関連
            ' urls.Add(902, "~/Pages/Master/ReferenceNumber.aspx")
            ' urls.Add(903, "~/Pages/Master/ItemMaster.aspx")
            ' urls.Add(904, "~/Pages/Master/Calendar.aspx")
            ' urls.Add(905, "~/Pages/Master/Supplier.aspx")
            ' urls.Add(906, "~/Pages/Master/ProductName.aspx")
            urls.Add(907, "~/Pages/Master/ExpenseCode.aspx")

            Return urls
        End Function
#End Region

    End Class
End Namespace
