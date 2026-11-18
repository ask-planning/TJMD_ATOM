Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' 輸入進捗（dbo.tbl_t_import_progress）を進める処理の種類です。
    ''' </summary>
    ''' <remarks>
    ''' 旧 Access 版の共通処理 `schedule_update` の引数 `syori_no` に対応します。
    ''' **値そのものが旧システムと同じでなければなりません。**更新する列と進捗状況の値が
    ''' この番号から決まるためです。
    '''
    ''' 旧 VBA のコメントに載っていた一覧は次のとおりです（1・2 は欠番）。
    ''' 3=上海出港 / 4=上海連絡 書類 / 5=上海連絡 ヘッダ / 6=検査品 / 7=通関依頼 /
    ''' 8=調達メール連絡 / 9=調達通関許可連絡 / 10=配送連絡⇒ロジ / 11=配送連絡⇒工場 /
    ''' 12=配送連絡⇒貿易（日新）/ 13=配送連絡⇒貿易 / 14=配送連絡⇒調達 / 15=保険 /
    ''' 16=ACCESS処理 / 17=経費入力
    '''
    ''' ここには**ATOM で使うものだけ**を定義しています。
    ''' </remarks>
    Public Enum ImportProgressSyori

        ''' <summary>通関依頼。配送手配書を発行したとき（menuNo=103 §16）。</summary>
        TsukanIrai = 7

        ''' <summary>配送連絡⇒ロジ。入荷明細書を発行したとき（menuNo=103 §16）。</summary>
        HaisoLogi = 10

        ''' <summary>配送連絡⇒貿易（日新）。入荷スケジュールを出力したとき（menuNo=103 §16）。</summary>
        HaisoNisshin = 12

    End Enum

End Namespace
