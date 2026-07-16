\xEF\xBB\xBFImports System.Windows.Forms
Imports Atom.App.Forms.Expense

Namespace Common
    ''' <summary>メニュー番号 → 画面 の対応（現行 指定フォームオープン 相当）。実装済みは New して返し、未実装は Nothing。</summary>
    Public Module ScreenFactory
        ''' <summary>menuNo に対応する画面を生成して返す（未実装は Nothing）。</summary>
        ''' <param name="menuNo">メニュー番号。</param>
        Public Function Create(menuNo As Integer) As Form
            ' menuNo ごとに対応画面を返す（実装後に該当 Case で New する）
            Select Case menuNo
                Case 101, 102, 103, 104, 105, 106, 107 ' 上海輸入定期便＆AIR便
                Case 201, 202, 203 ' その他輸入関連（調達）
                Case 301 ' 諸経費入力
                Case 902, 903, 904, 905, 906 ' マスタ関連
                Case 907 ' 経費コード登録
                    Return New ExpenseCodeMaintenanceForm()
            End Select
            Return Nothing ' 未実装
        End Function
    End Module
End Namespace
