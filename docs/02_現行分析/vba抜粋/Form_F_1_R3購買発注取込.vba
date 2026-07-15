Attribute VB_Name = "Form_F_1_R3購買発注取込"
Attribute VB_Base = "0{92D6DAA1-2617-4B17-B8A5-3577869E5B54}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit
    
    Dim intRef As Integer


Private Sub bnt_menu_Click()

    'メインメニューへ
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_hacchu_check_Click()

    '購買発注確認へ
    strFormName = "F_1_R3購買発注_MAIN"
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 202
    Else
        intMenuNo = 105
    End If
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_import_click()

    '///////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    'IMPORTテーブルはIMPORTだけ、フィールド名がどうなるか分からないので、実際のチェックはチェック用テーブルで行う
    
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////

    If IsNull(Me![PASS]) Then
        MsgBox ("取り込み元のパスを入力してください。")
        Exit Sub
    End If
    If IsNull(Me![text_name]) Then
        MsgBox ("取り込み元のテキスト名を入力してください。")
        Exit Sub
    End If
    
    'テキスト名称の拡張子が無かったらタイプによって拡張子をつける
    If Not IsNull(Me![text_name]) Then
        If Me![frm_text_type] = 1 Then
            If Right(Me![text_name], 4) <> ".xls" And Right(Me![text_name], 5) <> ".xlsx" Then
                Me![text_name] = Me![text_name] & ".xls"
            End If
        Else
            If Right(Me![text_name], 4) <> ".txt" Then
                Me![text_name] = Me![text_name] & ".txt"
            End If
        End If
    End If
    
    '上海以外の購買情報取り込みの場合、INVOICE_NOを入力する
    If Me![shanghai_chotatsu] = 2 Then
        If Me![frm_text_select] = 1 Then
            If IsNull(Me![REF_NO]) Then
                ReturnValue = MsgBox("購買発注データの取り込みの場合、INVOICE_NOを入力して下さい。" _
                    & vbLf & "在庫転送、出荷伝票の場合、INVOCE_NOは不要です。" & vbLf & "入力しますか？（入力／はい　入れない／いいえ）", 4)
                If ReturnValue = vbYes Then
                    Exit Sub
                End If
            Else
                strInvoiceNo = Me![REF_NO]
            End If
        End If
    End If
    
    '抜取品データの取込み時は、INVOICE_NOは必須
    If Me![frm_text_select] = 4 Then
        If IsNull(Me![REF_NO]) Or Len([REF_NO]) <= 0 Then
            MsgBox ("抜取品の取り込みの場合は、INVOICE_NOを入力してください。")
            Exit Sub
        End If
    End If
    
    Dim strZT053Text As String
    Dim strTextEdit As String
    
    Dim strImportTeigi As String
    Dim strChkTableName As String
    Dim strSyori As String
    Dim strOKFlg As String
    Dim intRCount As Integer
    Dim intLoopFlg As Integer
    Dim intLoopCount As Integer
    Dim intSeqNo As Integer
    
    '購買発注のチェックフィールド番号（フォーマットが変わることがあるので項目マスタに登録）
    Dim intDenpyoTypeField As Integer
    Dim intDenpyoHidukeField As Integer
    Dim intNonyuKijitsuField As Integer
    Dim intHinmokuCodeField As Integer
    Dim intSyanaiSansyoNoField As Integer
    Dim intSyukkaDenpyoField As Integer
    Dim intShiireCodeField As Integer
    
    Dim strSyanaiSansyoNo3 As String
    Dim intDenpyoKbn As Integer
    Dim strShiireCode As String
    
    '太田2F発注チェック用、価格適用期間
    Dim intKakakuKikanCount As Integer
    Dim intTaisyoYyyyMM() As Double
    Dim datTekiyoKaishiYmd() As Date
    Dim datTekiyoSyuryoYmd() As Date
    Dim dblShanghaiTajimaYyyyMm() As Double
    Dim datShanghaiTajimaKaishiYmd() As Date
    Dim datShanghaiTajimaSyuryoYmd() As Date
    Dim intKakakuMatchFlg As Integer
    
    Dim datDenpyoHiduke As Date
    
    Set cmd.ActiveConnection = conn
    
    'エラーの内容によってメッセージをどうするかをフラグで判断　⇒　strSyori
    strSyori = "tempdelete"

    '作業用テーブル名を変数に取っておく(増えた時のため、CASE分にした）
    On Error Resume Next
    '何の処理かの識別⇒intRef(1:購買発注、2:在庫転送、3:出荷伝票、5:2+3 9:支給包装材)
    intRef = 0
    Select Case Me![frm_text_select]
        Case 1, 5
            '購買発注、または、大田2F発注
            strChkTableName = "tbl_t_R3_koubai_hacchu_import_check_temp"               'チェック用テーブル
            If Me![frm_text_type] = 1 Then
                strTableName = "tbl_t_R3_koubai_hacchu_import_excel_temp_" & Me![login_code]             'EXCEL用IMPORTテーブル
            Else
                strTableName = "tbl_t_R3_koubai_hacchu_import_temp_" & Me![login_code]       'TEXT用IMPORTテーブル
            End If
            strFileName = Me![PASS] & "\" & Me![text_name]          '取り込むファイル名
            strImportTeigi = "ZT019_IMPORT"                         'TEXT取り込みようのIMPORT定義の名称
            
            '識別に購買発注をセット　⇒　取込むときに在庫転送と出荷伝票は別識別する
            If Me![frm_text_select] = 1 Then
                intRef = 1
            Else
                intRef = 10
            End If
            
        Case 4
            strChkTableName = "tbl_t_nukitori_kensa_import_check_temp"
            If Me![frm_text_type] = 1 Then
                strTableName = "T_抜取検査品_IMPORT_TEMP_" & Me![login_code]
            Else
                strTableName = "T_抜取検査品_IMPORT_TEXT_TEMP_" & Me![login_code]
            End If
            strFileName = Me![PASS] & "\" & Me![text_name]
            
            '識別に抜取検査品をセット
            strImportTeigi = "NUKITORI_IMPORT"
            intRef = 8
            
        Case Else
            strChkTableName = "tbl_t_R3_koubai_hacchu_import_check_temp"
            If Me![frm_text_type] = 1 Then
                strTableName = "tbl_t_R3_koubai_hacchu_import_excel_temp_" & Me![login_code]
            Else
                strTableName = "tbl_t_R3_koubai_hacchu_import_temp_" & Me![login_code]
            End If
            strFileName = Me![PASS] & "\" & Me![text_name]
            strImportTeigi = "hacchu_import"
    End Select
    
    'IMPORT用テンポラリファイルの削除＝EXCEL、クリア＝TEXT（削除時のエラーは無視）⇒deleteobjectはSQL SERVER2014では使用不可
    'If Me![frm_text_type] = 1 Then
    '    DoCmd.DeleteObject acTable, [strTableName]
    'Else
        With cmd
            .CommandText = "DELETE " & strTableName & " FROM " & strTableName & " "
            .CommandType = adCmdText
            .Execute
        End With
    'End If
    
    On Error GoTo 0
    
    strSyori = "tempdelete_end"
    
    '取り込みエラーのときは多分フォーマット違いの為、メッセージを表示する、たぶんだけど・・・
    On Error GoTo Err_btn_import_Click
    
    'EXCELはヘッダーありで、TEXTはテーブルにIMPORT
    If Me![frm_text_type] = 1 Then
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, [strTableName], strFileName, True
    'Else
    '    DoCmd.TransferText acImportDelim, [strImportTeigi], [strTableName], strFileName, False, ""
    End If
    
    On Error GoTo 0

    'チェック用テーブルのクリア
    With cmd
        .CommandText = "DELETE " & strChkTableName & " FROM " & strChkTableName & " WHERE pk_login_code = '" & Me![login_code] & "' "
        .CommandType = adCmdText
        .Execute
    End With
    
    '項目マスタからテーブルの何フィールド目をチェックするかというフィールド番号を取り込む
    With cmd
        .CommandText = "usp_koumoku_master_pickup"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "IMP"
    End With
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
        rs1.MoveFirst
        n = 0
        Do Until rs1.EOF
            Select Case rs1![contents1].Value
                Case "伝票タイプ"
                    intDenpyoTypeField = rs1![pk_seq_no].Value
                    n = n + 1
                Case "伝票日付"
                    intDenpyoHidukeField = rs1![pk_seq_no].Value
                    n = n + 1
                Case "納入期日"
                    intNonyuKijitsuField = rs1![pk_seq_no].Value
                    n = n + 1
                Case "品目コード"
                    intHinmokuCodeField = rs1![pk_seq_no].Value
                    n = n + 1
                Case "社内参照番号"
                    intSyanaiSansyoNoField = rs1![pk_seq_no].Value
                    n = n + 1
                Case "出荷伝票"
                    intSyukkaDenpyoField = rs1![pk_seq_no].Value
                    n = n + 1
                Case "仕入先"
                    intShiireCodeField = rs1![pk_seq_no].Value
                    n = n + 1
            End Select
            rs1.MoveNext
        Loop
        rs1.Close
    End If
    
    'ありませんように
    If n <> 7 Then
        MsgBox ("取込みフィールドの設定の不足があります。" & vbLf & "ＩＴ室に問い合わせてください。")
        GoTo Exit_btn_import_Click
    End If
    
    
    '太田2F発注取込用、価格適用期間
    intKakakuKikanCount = 0
    With cmd
        .CommandText = "usp_kansan_rate_count"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Execute
    End With
    If Not IsNull(cmd.Parameters(1)) Then
        intKakakuKikanCount = cmd.Parameters(1)
    End If
    
    ReDim intTaisyoYyyyMM(intKakakuKikanCount) As Double
    ReDim datTekiyoKaishiYmd(intKakakuKikanCount) As Date
    ReDim datTekiyoSyuryoYmd(intKakakuKikanCount) As Date
    ReDim dblShanghaiTajimaYyyyMm(intKakakuKikanCount) As Double
    ReDim datShanghaiTajimaKaishiYmd(intKakakuKikanCount) As Date
    ReDim datShanghaiTajimaSyuryoYmd(intKakakuKikanCount) As Date
    For n = 0 To intKakakuKikanCount
        intTaisyoYyyyMM(n) = 0
        datTekiyoKaishiYmd(n) = #1/1/1900#
        datTekiyoSyuryoYmd(n) = #1/1/1900#
        dblShanghaiTajimaYyyyMm(n) = 0
        datShanghaiTajimaKaishiYmd(n) = #1/1/1900#
        datShanghaiTajimaSyuryoYmd(n) = #1/1/1900#
    Next
    
    '価格適用期間抽出
    With cmd
        .CommandText = "usp_kansan_rate_pickup"
        .CommandType = adCmdStoredProc
    End With
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
        n = 0
        rs1.MoveFirst
        Do Until rs1.EOF
            intTaisyoYyyyMM(n) = rs1![pk_yyyymm].Value
            If Not IsNull(rs1![shanghai_price_ymd_kaishi].Value) Then
                datTekiyoKaishiYmd(n) = rs1![shanghai_price_ymd_kaishi].Value
            End If
            If Not IsNull(rs1![shanghai_price_ymd_syuryo].Value) Then
                datTekiyoSyuryoYmd(n) = rs1![shanghai_price_ymd_syuryo].Value
            End If
            If rs1![pk_yyyymm].Value >= 201307 Then
                If Not IsNull(rs1![shanghai_tajima_kakakuhyo_yyyymm].Value) Then
                    dblShanghaiTajimaYyyyMm(n) = rs1![shanghai_tajima_kakakuhyo_yyyymm].Value
                    datShanghaiTajimaKaishiYmd(n) = DateSerial(Left(rs1![pk_yyyymm].Value, 4), Right(rs1![pk_yyyymm].Value, 2), 1)
                    datShanghaiTajimaSyuryoYmd(n) = DateSerial(Year(datShanghaiTajimaKaishiYmd(n)), Month(datShanghaiTajimaKaishiYmd(n)) + 1, 1) - 1
                End If
            Else
                dblShanghaiTajimaYyyyMm(n) = rs1![pk_yyyymm].Value
                If Not IsNull(rs1![shanghai_price_ymd_kaishi].Value) Then
                    datShanghaiTajimaKaishiYmd(n) = rs1![shanghai_price_ymd_kaishi].Value
                End If
                If Not IsNull(rs1![shanghai_price_ymd_syuryo].Value) Then
                    datShanghaiTajimaSyuryoYmd(n) = rs1![shanghai_price_ymd_syuryo].Value
                End If
            End If
            n = n + 1
            rs1.MoveNext
        Loop
    End If
    rs1.Close
    
    'INSERT句の固定部分
    If Me![frm_text_select] = 4 Then
        strInsertKu = "INSERT INTO tbl_t_nukitori_kensa_import_check_temp ( pk_login_code, pk_invoice_no_main, hinmoku_code, kensa_kbn ) "
        strInsertKu = strInsertKu & "VALUES ( '" & Me![login_code] & "', '" & Me![REF_NO] & "', "
    Else
        strInsertKu = "INSERT INTO tbl_t_R3_koubai_hacchu_import_check_temp ( pk_login_code, pk_koubai_hacchu_no, koubai_soshiki, denpyo_type, "
        strInsertKu = strInsertKu & "denpyo_ymd, nounyu_kijitsu_ymd, koubai_group, shiiresaki_code, shiiresaki_name, kyoukyu_plant, "
        strInsertKu = strInsertKu & "line_seq, hinmoku_code, hinmoku_code_text, plant, hokan_basyo, hinmoku_group, koubai_hacchu_suryo, "
        strInsertKu = strInsertKu & "suryo_tani, syomi_kakaku, syomi_kingaku, tsuka, syanai_sansyo_no, touroku_tanto, nounyuzumi_flg, "
        strInsertKu = strInsertKu & "delete_flg, syukka_denpyo_no , nyuko_denpyo_no, denpyo_suryo, suryo_tani2, torihiki_type, hacchu_riyu, sansyo_code, "
        strInsertKu = strInsertKu & "syanai_sansyo_no3, denpyo_kbn, syanai_sansyo_no_original, taisyo_yyyymm, shanghai_price, err_contents, err_status ) "
        strInsertKu = strInsertKu & "VALUES ( '" & Me![login_code] & "', "
    End If

    
    'IMPORTテーブルの読み込みと、チェック用テンポラリのクリアとレコードセット取得
            
    'On Error GoTo Err_btn_import_Click
    strSyori = "format_err"
    
    If Me![frm_text_type] = 2 Then
    
        'テキストファイル(.txt)からの取込み（抜きとり検査は無いです。）
        
        'レコードの読み込み
        intFileNo = FreeFile(0)
        Open strFileName For Input As #intFileNo
    
    
        Do While Not EOF(intFileNo)
        
            'カンマデータがあるので、LINE INPUTで処理
            Line Input #intFileNo, strZT053Text
            
            t = 0
            strTextEdit = ""
            strValuesKu = ""
            intFieldCount = 0
            strSyanaiSansyoNo = ""
            strSyanaiSansyoNo3 = ""
            intDenpyoKbn = 0
            
            If Len(strZT053Text) > 0 Then
            
                For n = 1 To Len(strZT053Text)
                    'タブ区切り　⇒　１文字ずつチェックして、タブが合ったらフィールドのセット
                    If Mid(strZT053Text, n, 1) = vbTab Then
                        'If Len(strValuesKu) > 0 Then
                        '    strValuesKu = strValuesKu & ", "
                        'End If
                        
                        'SQL分作成（VALUE句のデータ部分）
                        '購買発注/在庫転送/出荷伝票の場合フィールド番号で元データをチェック用テーブルに
                        If intFieldCount = intSyanaiSansyoNoField Then
                            If Len(LTrim(RTrim(strTextEdit))) > 0 Then
                                strSyanaiSansyoNo = LTrim(RTrim(strTextEdit))
                            End If
                            If IsNumeric(Mid(LTrim(RTrim(strTextEdit)), 3, 1)) And Not IsNumeric(Left(LTrim(RTrim(strTextEdit)), 2)) Then
                                
                                '社内参照番号の３バイト目が数字で、先頭２バイトがテキストタイプの時、先頭２バイトを使用する。⇒　KH、LAなど
                                strSyanaiSansyoNo3 = Left(LTrim(RTrim(strTextEdit)), 2)
                                
                            Else
                                
                                If (Left(LTrim(RTrim(strTextEdit)), 2) = Right(Year(Date), 2) _
                                    Or Left(LTrim(RTrim(strTextEdit)), 2) = Right(Year(Date) - 1, 2)) _
                                    And Not IsNumeric(Mid(LTrim(RTrim(strTextEdit)), 3, 1)) Then
                                    
                                    '先頭２バイトが今年、または昨年の年の下二桁で、３バイト目が文字タイプの時、年（下二桁）＋A　⇒　06A、06Bなどは06Aにする
                                    strSyanaiSansyoNo3 = Right(Year(Date), 2) & "A"
                                
                                Else
                                    
                                    If Left(LTrim(RTrim(strTextEdit)), 2) = "SH" Then
                                        '先頭２バイトが『SH』は、固定でSH
                                        strSyanaiSansyoNo3 = Left(LTrim(RTrim(strTextEdit)), 2)
                                    
                                    Else
                                        'その他は、先頭３バイトを使用する
                                        strSyanaiSansyoNo3 = Left(LTrim(RTrim(strTextEdit)), 3)
                                    End If
                                End If
                            End If
                        End If
                        '伝票タイプ＝NBは、購買発注、伝票タイプ＝UBは、出荷伝票があれば出荷伝票、無ければ在庫転送
                        '在庫転送と出荷伝票は一緒の場合があるので、明細に識別をセットする
                        If intFieldCount = intDenpyoTypeField Then
                            If LTrim(RTrim(strTextEdit)) = "UB" Then
                                If Len(LTrim(RTrim(strTextEdit))) > 0 And Not IsNull(LTrim(RTrim(strTextEdit))) Then
                                    
                                    If intRef <> 5 Then
                                        If intRef = 2 Then
                                            intRef = 5          'この後の処理用の区分に在庫転送、出荷伝票混合をセット
                                        Else
                                            intRef = 3          'この後の処理用の区分に出荷伝票をセット
                                        End If
                                    End If
                                    intDenpyoKbn = 3
                                Else
                                    If intRef <> 5 Then
                                        If intRef = 3 Then
                                            intRef = 5          'この後の処理用の区分に在庫転送、出荷伝票混合をセット
                                        Else
                                            intRef = 2          'この後の処理用の区分に在庫転送をセット
                                        End If
                                    End If
                                    intDenpyoKbn = 2
                                End If
                            Else
                                intDenpyoKbn = 1
                            End If
                        End If
                        If intFieldCount = intShiireCodeField Then
                            If Len(LTrim(RTrim(strTextEdit))) > 0 Then
                                strShiireCode = LTrim(RTrim(strTextEdit))
                            End If
                        End If
                    
                        If intFieldCount = intDenpyoHidukeField Or intFieldCount = intNonyuKijitsuField Then
                            '日付フィールドは２パターンで取込　⇒　textファイル（2006.01.01）、EXCELファイル（2006/01/01）
                            If IsDate(LTrim(RTrim(strTextEdit))) = True Then
                                strValuesKu = strValuesKu & "'" & LTrim(RTrim(strTextEdit)) & "', "
                                If n = intDenpyoHidukeField Then
                                    datDenpyoHiduke = LTrim(RTrim(strTextEdit))
                                End If
                            Else
                                If IsNumeric(Left(LTrim(RTrim(strTextEdit)), 4)) Then
                                    strValuesKu = strValuesKu & "'" & DateSerial(Left(LTrim(RTrim(strTextEdit)), 4), Mid(LTrim(RTrim(strTextEdit)), 6, 2), Right(LTrim(RTrim(strTextEdit)), 2)) & "', "
                                    If n = intDenpyoHidukeField Then
                                        datDenpyoHiduke = DateSerial(Left(LTrim(RTrim(strTextEdit)), 4), Mid(LTrim(RTrim(strTextEdit)), 6, 2), Right(LTrim(RTrim(strTextEdit)), 2))
                                    End If
                                Else
                                    strValuesKu = strValuesKu & "null, "
                                End If
                            End If
                        Else
                            '品目コードの編集
                            If intFieldCount = intHinmokuCodeField Then
                                If IsNumeric(LTrim(RTrim(strTextEdit))) Then
                                    strValuesKu = strValuesKu & "'" & Right("000000000000000000" & LTrim(RTrim(strTextEdit)), 18) & "', "
                                Else
                                    strValuesKu = strValuesKu & "'" & LTrim(RTrim(strTextEdit)) & "', "
                                End If
                            Else
                                If intFieldCount = 15 Or intFieldCount = 17 Or intFieldCount = 18 Or intFieldCount = 26 Or intFieldCount = 28 Then
                                    If Len(LTrim(RTrim(strTextEdit))) > 0 Then
                                        strValuesKu = strValuesKu & LTrim(RTrim(strTextEdit)) & ", "
                                    Else
                                        strValuesKu = strValuesKu & "null, "
                                    End If
                                Else
                                    If Len(LTrim(RTrim(strTextEdit))) > 0 Then
                                        strValuesKu = strValuesKu & "'" & LTrim(RTrim(strTextEdit)) & "', "
                                    Else
                                        strValuesKu = strValuesKu & "null, "
                                    End If
                                End If
                            End If
                        End If
                    
                        strTextEdit = ""
                        'タブ毎にフィールド数をカウントアップ
                        intFieldCount = intFieldCount + 1
                    Else
                    
                        'タブ以外は、くっつける
                        strTextEdit = strTextEdit & Mid(strZT053Text, n, 1)
                    End If
                Next
            
            
                '最後のフィールドのSQL分作成（VALUE句のデータ部分）
                If Len(strTextEdit) > 0 And Len(LTrim(RTrim(strTextEdit))) > 0 Then
                    strValuesKu = strValuesKu & "'" & LTrim(RTrim(strTextEdit)) & "', "
                    intFieldCount = intFieldCount + 1
                Else
                    strValuesKu = strValuesKu & "null, "
                End If
                                                
                'PO番号の編集（移動先の決定用）、特殊なPOはナカナカ無理なので、後でエラー表示する
                If Len(strSyanaiSansyoNo3) > 0 Then
                    strValuesKu = strValuesKu & "'" & strSyanaiSansyoNo3 & "', "
                Else
                    strValuesKu = strValuesKu & "null, "
                End If
                
                '伝票タイプ＝NBは、購買発注、伝票タイプ＝UBは、出荷伝票があれば出荷伝票、無ければ在庫転送
                '在庫転送と出荷伝票は一緒の場合があるので、明細に識別をセットする
                strValuesKu = strValuesKu & intDenpyoKbn & ", "
                
                '社内参照番号ORIGINAL
                If Len(strSyanaiSansyoNo) > 0 Then
                    strValuesKu = strValuesKu & "'" & strSyanaiSansyoNo & "', "
                Else
                    strValuesKu = strValuesKu & "null, "
                End If
            
                '価格適用月
                intKakakuMatchFlg = 0
                If Len(strShiireCode) > 0 Then
                    For t = 0 To intKakakuKikanCount - 1
                        If strShiireCode = "0000133128" Then
                            If datDenpyoHiduke >= datShanghaiTajimaKaishiYmd(t) And datDenpyoHiduke <= datShanghaiTajimaSyuryoYmd(t) Then
                                strValuesKu = strValuesKu & dblShanghaiTajimaYyyyMm(t)
                                intKakakuMatchFlg = 1
                                Exit For
                            End If
                        Else
                            If datDenpyoHiduke >= datTekiyoKaishiYmd(t) And datDenpyoHiduke <= datTekiyoSyuryoYmd(t) Then
                                strValuesKu = strValuesKu & intTaisyoYyyyMM(t) & ", "
                                intKakakuMatchFlg = 1
                                Exit For
                            End If
                        End If
                    Next
                End If
                If intKakakuMatchFlg = 0 Then
                    strValuesKu = strValuesKu & "0, "
                End If
                strValuesKu = strValuesKu & "0, null, 0 ) "
                
                'INSERT句とVALUE句を結合
                strSql = strInsertKu & strValuesKu
                
                'Debug.Print strSql
                'Debug.Print strZT053Text
                
                '更新
                With cmd
                    .CommandText = strSql
                    .CommandType = adCmdText
                    .Execute
                End With
                
            End If
        Loop
        
        Close #intFileNo
    
    
    Else
    
        'EXCELファイル
        With cmd
            .CommandText = "SELECT * FROM " & strTableName & " "
            .CommandType = adCmdText
        End With
        
        Set rs1 = cmd.Execute(intRcnt)
        If intRcnt <> 0 Then
    
            'IMPORTテーブルに中身があるか？
            'IMPORTテーブルはフィールド名が当てにならないので、フィールド番号でチェック
    
            rs1.MoveFirst
            
            intSeqNo = 0
            Do Until rs1.EOF
                
                'まず、『NG』をセットして、OKだったら『OK』をセットする
                '　　⇒必須フィールドに必ずデータがある時にOKだけど、今回は何でもOK、でも処理は残しておく
                strOKFlg = "NG"
                If Len(rs1.Fields(0).Value) > 0 Then
                        strOKFlg = "OK"
                End If
                
                '『OK』だったらチェック用テンポラリに取り込み
                If strOKFlg = "OK" Then
                
                    strValuesKu = ""
                    Select Case Me![frm_text_select]
                        Case 1, 5

                            '購買発注/在庫転送/出荷伝票の場合フィールド番号で元データをチェック用テーブルに
                            For n = 0 To rs1.Fields.Count - 1
                                    
                                If n = intDenpyoHidukeField Or n = intNonyuKijitsuField Then
                                    '日付フィールドは２パターンで取込　⇒　textファイル（2006.01.01）、EXCELファイル（2006/01/01）
                                    If IsDate(rs1.Fields(n).Value) = True Then
                                        strValuesKu = strValuesKu & "'" & LTrim(RTrim(rs1.Fields(n).Value)) & "', "
                                        If n = intDenpyoHidukeField Then
                                            datDenpyoHiduke = LTrim(RTrim(rs1.Fields(n).Value))
                                        End If
                                    Else
                                        If IsNumeric(Left(rs1.Fields(n).Value, 4)) Then
                                            strValuesKu = strValuesKu & "'" & DateSerial(Left(rs1.Fields(n).Value, 4), Mid(rs1.Fields(n).Value, 6, 2), Right(rs1.Fields(n).Value, 2)) & "', "
                                            If n = intDenpyoHidukeField Then
                                                datDenpyoHiduke = DateSerial(Left(rs1.Fields(n).Value, 4), Mid(rs1.Fields(n).Value, 6, 2), Right(rs1.Fields(n).Value, 2))
                                            End If
                                        Else
                                            strValuesKu = strValuesKu & "null, "
                                        End If
                                    End If
                                Else
                                    '品目コードの編集
                                    If n = intHinmokuCodeField Then
                                        If IsNumeric(rs1.Fields(intHinmokuCodeField).Value) Then
                                            strValuesKu = strValuesKu & "'" & Right("000000000000000000" & rs1.Fields(intHinmokuCodeField).Value, 18) & "', "
                                        Else
                                            strValuesKu = strValuesKu & "'" & rs1.Fields(intHinmokuCodeField).Value & "', "
                                        End If
                                    Else
                                        If n = 9 Or n = 15 Or n = 17 Or n = 18 Or n = 26 Or n = 28 Then
                                            If Len(LTrim(RTrim(rs1.Fields(n).Value))) > 0 Then
                                                strValuesKu = strValuesKu & LTrim(RTrim(rs1.Fields(n).Value)) & ", "
                                            Else
                                                strValuesKu = strValuesKu & "null, "
                                            End If
                                        Else
                                            If Len(LTrim(RTrim(rs1.Fields(n).Value))) > 0 Then
                                                strValuesKu = strValuesKu & "'" & LTrim(RTrim(rs1.Fields(n).Value)) & "', "
                                            Else
                                                strValuesKu = strValuesKu & "null, "
                                            End If
                                        End If
                                    End If
                                End If
                                    
                            Next
                                            
                            'PO番号の編集（移動先の決定用）、特殊なPOはナカナカ無理なので、後でエラー表示する
                            If IsNumeric(Mid(rs1.Fields(intSyanaiSansyoNoField).Value, 3, 1)) And Not IsNumeric(Left(rs1.Fields(intSyanaiSansyoNoField).Value, 2)) Then
                                
                                '社内参照番号の３バイト目が数字で、先頭２バイトがテキストタイプの時、先頭２バイトを使用する。⇒　KH、LAなど
                                strValuesKu = strValuesKu & "'" & Left(rs1.Fields(intSyanaiSansyoNoField).Value, 2) & "', "
                                
                            Else
                                
                                If (Left(rs1.Fields(intSyanaiSansyoNoField).Value, 2) = Right(Year(Date), 2) _
                                    Or Left(rs1.Fields(intSyanaiSansyoNoField).Value, 2) = Right(Year(Date) - 1, 2)) _
                                    And Not IsNumeric(Mid(rs1.Fields(intSyanaiSansyoNoField).Value, 3, 1)) Then
                                    
                                    '先頭２バイトが今年、または昨年の年の下二桁で、３バイト目が文字タイプの時、年（下二桁）＋A　⇒　06A、06Bなどは06Aにする
                                    strValuesKu = strValuesKu & "'" & Right(Year(Date), 2) & "A" & "', "
                                
                                Else
                                    
                                    If Left(rs1.Fields(intSyanaiSansyoNoField).Value, 2) = "SH" Then
                                        '先頭２バイトが『SH』は、固定でSH
                                        strValuesKu = strValuesKu & "'" & Left(rs1.Fields(intSyanaiSansyoNoField).Value, 2) & "', "
                                    
                                    Else
                                        'その他は、先頭３バイトを使用する
                                        strValuesKu = strValuesKu & "'" & Left(rs1.Fields(intSyanaiSansyoNoField).Value, 3) & "', "
                                    End If
                                End If
                            End If
                            '伝票タイプ＝NBは、購買発注、伝票タイプ＝UBは、出荷伝票があれば出荷伝票、無ければ在庫転送
                            '在庫転送と出荷伝票は一緒の場合があるので、明細に識別をセットする
                            If LTrim(RTrim(rs1.Fields(intDenpyoTypeField).Value)) = "UB" Then
                                If Len(LTrim(RTrim(rs1.Fields(intSyukkaDenpyoField).Value))) > 0 And Not IsNull(rs1.Fields(intSyukkaDenpyoField).Value) Then
                                    
                                    If intRef <> 5 Then
                                        If intRef = 2 Then
                                            intRef = 5          'この後の処理用の区分に在庫転送、出荷伝票混合をセット
                                        Else
                                            intRef = 3          'この後の処理用の区分に出荷伝票をセット
                                        End If
                                    End If
                                    strValuesKu = strValuesKu & "3, "
                                Else
                                    If intRef <> 5 Then
                                        If intRef = 3 Then
                                            intRef = 5          'この後の処理用の区分に在庫転送、出荷伝票混合をセット
                                        Else
                                            intRef = 2          'この後の処理用の区分に在庫転送をセット
                                        End If
                                    End If
                                    strValuesKu = strValuesKu & "2, "
                                End If
                            Else
                                strValuesKu = strValuesKu & "1, "
                            End If
                            '社内参照番号ORIGINAL
                            strValuesKu = strValuesKu & "'" & rs1.Fields(intSyanaiSansyoNoField).Value & "', "
                            
                            '価格適用月
                            intKakakuMatchFlg = 0
                            'If Me![frm_text_select] = 5 Then
                                If Len(strShiireCode) > 0 Then
                                    For t = 0 To intKakakuKikanCount - 1
                                        If strShiireCode = "0000133128" Then
                                            If datDenpyoHiduke >= datShanghaiTajimaKaishiYmd(t) And datDenpyoHiduke <= datShanghaiTajimaSyuryoYmd(t) Then
                                                strValuesKu = strValuesKu & dblShanghaiTajimaYyyyMm(t)
                                                intKakakuMatchFlg = 1
                                                Exit For
                                            End If
                                        Else
                                            If datDenpyoHiduke >= datTekiyoKaishiYmd(t) And datDenpyoHiduke <= datTekiyoSyuryoYmd(t) Then
                                                strValuesKu = strValuesKu & intTaisyoYyyyMM(t) & ", "
                                                intKakakuMatchFlg = 1
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If
                                If intKakakuMatchFlg = 0 Then
                                    strValuesKu = strValuesKu & "0, "
                                End If
                            'Else
                            '    strValuesKu = strValuesKu & "0 ) "
                            'End If
                            strValuesKu = strValuesKu & "0, null, 0 ) "
                                                        
                        Case 4
                            '抜取検査品の取り込みは、チョコビットだし、フォーマットも変わらないだろうから、個別フィールドで
                            If IsNumeric(rs1.Fields(0).Value) Then
                                strValuesKu = "'" & Right("000000000000000000" & rs1.Fields(0).Value, 18) & "', "
                            Else
                                strValuesKu = "'" & rs1.Fields(0).Value & "', "
                            End If
                            strValuesKu = strValuesKu & rs1.Fields(1).Value & " ) "
                    End Select
                    
                    strSql = strInsertKu & strValuesKu
                    
                    '実行
                    With cmd
                        .CommandText = strSql
                        .CommandType = adCmdText
                        .Execute
                    End With
                    
                End If
                rs1.MoveNext
            Loop
        End If
        rs1.Close
    End If
    strSyori = "format_err_end"
    
    On Error GoTo 0
    
    '何を取り込もうとしているかを表示（全部ZT053なので、間違えるかもしれないから）
    If Me![shanghai_chotatsu] <> 2 Then
        Select Case intRef
            Case 1
                Me![REF_NO] = "購買発注の取り込みです。"
            Case 2
                Me![REF_NO] = "在庫転送の取り込みです。"
            Case 3
                Me![REF_NO] = "在庫転送と出荷伝票の取り込みです。"
            Case 5
                Me![REF_NO] = "在庫転送と出荷伝票の混合取り込みです。"
            'Case 8
            '    Me![REF_NO] = "抜取検査品の取り込みです。"
            'Case 9
            '    Me![REF_NO] = "支給包装材の取り込みです。"
            Case 10
                Me![REF_NO] = "太田2F発注の取り込みです。"
                
        End Select
        Me![REF_NO].Requery
    End If
    If Len(strShiireCode) > 0 Then
        With cmd
            .CommandText = "usp_shiiresaki_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = strShiireCode
            .Execute
        End With
        If IsNull(cmd.Parameters(2)) Then
            ReturnValue = MsgBox("仕入先の登録がありません。" & vbLf & "このまま取り込みますか？", vbYesNo)
            If ReturnValue = vbNo Then
                GoTo Exit_btn_import_Click
            End If
        End If
            
    End If
       
    Call btn_saijikko_Click
    Exit Sub

Exit_btn_import_Click:
    
    Set cmd = Nothing
    Exit Sub

Err_btn_import_Click:

    'エラーフラグの内容でメッセージを表示する
    Select Case strSyori
        Case "tempdelete"
            Resume Next
        Case "format_err"
            MsgBox ("フォーマットが違います。")
        Case Else
            MsgBox ("取り込み処理のエラーです。" & vbLf & "テキスト名、ファイルの場所等を確認してください。")
    End Select
    
End Sub

Private Sub btn_invoice_meisai_Click()

    'INVOICE内容確認へ
    strFormName = "F_1_上海輸入_INVOICE_MAIN"
    intMenuNo = 102
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_invoice_header_Click()

    'INVOICEヘッダ入力へ
    If Me![shanghai_chotatsu] = 2 Then
        strFormName = "F_2_調達輸入_INVOICE_HEADER"
        intMenuNo = 203
    Else
        strFormName = "F_1_上海輸入_INVOICE_HEADER"
        intMenuNo = 103
    End If
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_kakunin_Click()

    strFormName = "F_1_取込内容確認_MAIN"
    Select Case Me![frm_text_select]
        Case 1
            strSubFormName = "F_1_取込内容確認_SUB_購買発注"
        Case 4
            strSubFormName = "F_1_取込内容確認_SUB_抜取検査"
        Case 5
            strSubFormName = "F_1_取込内容確認_SUB_購買発注"
    End Select
    Call next_form_open
            
End Sub

Private Sub btn_ohta_2f_Click()

    '大田２F発注台帳確認へ
    strFormName = "F_1_太田2F発注_台帳照会_MAIN"
    intMenuNo = 109
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_saijikko_Click()

    DoCmd.GoToControl "btn_import"
    
    Dim intMitouroku As Integer
    Dim strOKFlg As String
    Dim intOkikae As Integer

    Dim strKeyOrderNo As String
    Dim intLen As Integer
    Dim intMidLen As Integer
    Dim intKahenLen As Integer
    
    Dim strTempTableName As String
    Dim strDisplayOrder As String
    
    Dim intHosozaiCount As Integer
    
    Dim i As Integer
    Dim intLoopFlg As Integer
    Dim intLoopCount As Integer
    Dim intSyanaiSansyoNoCount As Double
    Dim intCountPacking As Integer
    Dim intErrNo As Integer
    
    Dim strTblPONo() As String
    
    Dim intInvoiceDoubleCount As Integer
    
    Set cmd.ActiveConnection = conn
    
    intSyanaiSansyoNoCount = 0
    intHosozaiCount = 0
    
    '作業中テンポラリのエラーステータスフィールドのクリア
    With cmd
        .CommandText = "usp_R3_koubai_hacchu_import_temp_clear"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![frm_text_select]
        .Execute
    End With
    
    If Me![frm_text_select] = 1 Then
    
        '在庫転送/出荷伝票の取り込みの場合は、購買発注が取込済みで無いとだめ。
        With cmd
            .CommandText = "usp_R3_koubai_hacchu_import_count"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = 0
            .Execute
        End With
        
        '発注件数
        If Not IsNull(cmd.Parameters(3)) Then
            intSyanaiSansyoNoCount = cmd.Parameters(3)
        End If
        
    End If
        
    'If intRef = 3 Or intRef = 5 Then
        '出荷伝票の取り込みの場合は、在庫転送が取込み済みでないとだめ
    '    strSql = "SELECT Count(tbl_t_R3_koubai_hacchu_meisai.pk_koubai_hacchu_no) AS PO_count FROM tbl_t_R3_koubai_hacchu_meisai "
    '    strSql = strSql & "INNER JOIN tbl_t_R3_koubai_hacchu_import_check_temp ON tbl_t_R3_koubai_hacchu_meisai.[zaiko_tenso_no] = "
    '    strSql = strSql & "tbl_t_R3_koubai_hacchu_import_check_temp.pk_koubai_hacchu_no WHERE tbl_t_R3_koubai_hacchu_import_check_temp.[伝票区分] = 3 "
    '    'Set rs1 = db.OpenRecordset(strSql)
    '    If rs1.RecordCount <> 0 Then
    '        rs1.MoveFirst
    '        intSyanaiSansyoNoCount = intSyanaiSansyoNoCount + rs1![po_count].Value
    '    End If
    '    rs1.Close
    'End If
            
    If Me![frm_text_select] = 4 Then
        '抜取検査品の場合は、元の購買発注が取込み済みでないとだめ
        With cmd
            .CommandText = "usp_R3_koubai_hacchu_import_nukitori_count"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Execute
        End With
        '発注データのカウント
        If Not IsNull(cmd.Parameters(2)) Then
            intSyanaiSansyoNoCount = intSyanaiSansyoNoCount + cmd.Parameters(2)
        End If
    End If
    
    '在庫転送、出荷伝票、支給包装材、抜取検査品は購買発注が無いとだめというメッセージ
    If (intRef <> 1 And intRef <> 9 And intRef <> 10) And Me![frm_text_select] <> 6 And intSyanaiSansyoNoCount = 0 Then
        MsgBox ("購買発注、または在庫転送の取り込みがありません。")
        GoTo Exit_btn_saijikko_Click
    End If
    
    '購買発注登録済みデータのバックアップテーブルのクリア
    With cmd
        .CommandText = "DELETE tbl_t_R3_koubai_hacchu_meisai_temp FROM tbl_t_R3_koubai_hacchu_meisai_temp WHERE pk_login_code = '" & Me![login_code] & "' "
        .CommandType = adCmdText
        .Execute
    End With
    
    '既に取込済みのデータがあるか
    intSyanaiSansyoNoCount = 0
    With cmd
        .CommandText = "usp_R3_koubai_hacchu_import_count"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = intRef
        .Execute
    End With

    If Not IsNull(cmd.Parameters(3)) Then
        intSyanaiSansyoNoCount = cmd.Parameters(3)
    End If

    '在庫転送、出荷伝票の場合、出荷伝票の有無も調べる
    If intRef = 3 Or intRef = 5 Then
        If Not IsNull(cmd.Parameters(4)) Then
            intSyanaiSansyoNoCount = intSyanaiSansyoNoCount + cmd.Parameters(4)
        End If
    End If

    '登録済みの場合、置き換えるかどうか
    intOkikae = 0
    If intSyanaiSansyoNoCount <> 0 And intRef <> 9 And Me![frm_text_select] <> 6 Then
        Select Case intRef
            Case 1
                ReturnValue = MsgBox("登録済みの購買発注があります。" & vbLf & "置換（はい）／追加（いいえ）／キャンセル（キャンセル）", vbYesNoCancel)
            Case 8
                ReturnValue = MsgBox("登録済みの抜取検査品があります。" & vbLf & "置換（はい）／追加（いいえ）／キャンセル（キャンセル）", vbYesNoCancel)
            'Case 9
            '    ReturnValue = MsgBox("前日以降登録済みの支給包装材があります。" & vbLf & "置換（はい）／追加（いいえ）／キャンセル（キャンセル）", vbYesNoCancel)
            Case 10
                ReturnValue = MsgBox("登録済みの購買発注があります。" & vbLf & "置換（はい）／追加（いいえ）／キャンセル（キャンセル）", vbYesNoCancel)
            Case Else
                ReturnValue = MsgBox("登録済みの在庫転送／出荷伝票があります。" & vbLf & "置換（はい）／キャンセル（いいえ）", vbYesNo)
        End Select
        If ReturnValue = vbCancel Or (intRef <> 1 And intRef <> 9 And intRef <> 8 And intRef <> 10 And ReturnValue = vbNo) Then
            GoTo Exit_btn_saijikko_Click
        Else
            '置換えの場合、出荷伝票取込み以外は一旦BACKUPに取っておいて（後で戻す項目があるため）、
            '登録済みデータに削除フラグを立ててから、削除する（一発クエリーで削除できないから・・・）
            If ReturnValue = vbYes Then
                intOkikae = 1
            End If
        End If
    End If
    
    '取込データを読み込んで、各種チェックを行う
    Me![btn_saijikko].Visible = True
    
    strOKFlg = "OK"
    intErrNo = 0
    
    intInvoiceDoubleCount = 0
    '複数INVOICEに分かれているPOがあるか。あったら、INVOICEに合わせてPOを変更する
    If Me![shanghai_chotatsu] <> 2 And Me![frm_text_select] = 1 Then
        With cmd
            .CommandText = "usp_R3_koubai_hacchu_import_sansyo_no_change"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Execute
        End With
        If Not IsNull(cmd.Parameters(2)) And cmd.Parameters(2) <> 0 Then
            MsgBox ("複数INVOICEに分かれているPOを変更しました。" & vbLf & "変更したPOを確認後、再度、実行してください。")
            GoTo Exit_btn_saijikko_Click
        End If
    End If
    
    '選択別、作業テーブル、フォーム名、内容チェック
    Select Case Me![frm_text_select]
        Case 1, 5
            '重複品目のチェック
            With cmd
                .CommandText = "usp_R3_koubai_hacchu_import_jyufuku_check_hinmoku"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
            
            If Me![frm_text_select] = 5 Then
                With cmd
                    .CommandText = "usp_R3_koubai_hacchu_ohta_2F_price_check"
                    .CommandType = adCmdStoredProc
                    .Parameters.Refresh
                    .Parameters(1) = Me![login_code]
                    .Execute
                End With
            End If
            
            '社内参照番号の重複チェック
            With cmd
                .CommandText = "usp_R3_koubai_hacchu_import_jyufuku_check"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
            
            'いろいろエラーチェック
            With cmd
                .CommandText = "usp_R3_koubai_hacchu_import_err_check"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
            
            strTempTableName = "tbl_t_R3_koubai_hacchu_import_check_temp"
        Case 4
            '重複品目のチェック
            With cmd
                .CommandText = "usp_nukitori_kensa_import_jyujuku_check"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
            'いろいろエラーチェック
            With cmd
                .CommandText = "usp_nukitori_kensa_import_err_check"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
            strTempTableName = "tbl_t_nukitori_kensa_import_check_temp"
            
        Case 6
            With cmd
                .CommandText = "usp_R3_koubai_hacchu_import_misyukka_check"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
            strTempTableName = "tbl_t_R3_koubai_hacchu_import_check_temp"
            
    End Select
    
    If Not IsNull(cmd.Parameters(3)) And cmd.Parameters(3) <> 0 Then
        intErrNo = 2
    End If
    If Not IsNull(cmd.Parameters(2)) And cmd.Parameters(3) <> 0 Then
        intErrNo = 1
    End If
    
    'メッセージの表示
    Select Case intErrNo
        Case 1
            MsgBox ("エラーデータがあります。" & vbLf & "EXCELを確認してください。")
            strOKFlg = "NG"
        Case 2
            If Me![frm_text_select] = 6 Then
                ReturnValue = MsgBox("WARNINGデータがあります。" & vbLf & "確認する（はい）／終了する（いいえ）", 4)
                If ReturnValue = vbYes Then
                    strOKFlg = "NG"
                End If
            Else
                If Me![frm_text_select] = 6 Then
                    ReturnValue = MsgBox("WARNINGデータがあります。確認（はい）／終了（いいえ）", 4 + vbDefaultButton2)
                Else
                    ReturnValue = MsgBox("WARNINGデータがあります、このまま取り込みますか？" & vbLf & "取り込み（はい）／キャンセル（いいえ）", 4)
                End If
                If ReturnValue = vbNo Then
                    strOKFlg = "NG"
                End If
            End If
        Case Else
            If Me![frm_text_select] = 6 Then
                ReturnValue = MsgBox("内容の確認を行いますか？　確認（はい）／終了（いいえ）", 4 + vbDefaultButton2)
            Else
                ReturnValue = MsgBox("内容の確認を行いますか？　確認（はい）／取込（いいえ）", 4 + vbDefaultButton2)
            End If
            If ReturnValue = vbYes Then
                strOKFlg = "NG"
            End If
    End Select
            
    'どこかで内容確認をすると選んでいたら、このフラグに「NG」がセットされている。
    '内容チェック用フォームのオープン
    If strOKFlg = "NG" Then
        strFormName = "F_1_取込内容確認_MAIN"
        Select Case Me![frm_text_select]
            Case 1
                strSubFormName = "F_1_取込内容確認_SUB_購買発注"
            Case 4
                strSubFormName = "F_1_取込内容確認_SUB_抜取検査"
            Case 5
                strSubFormName = "F_1_取込内容確認_SUB_購買発注"
            Case 6
                strSubFormName = "F_1_取込内容確認_SUB_購買発注"
        End Select
        Call next_form_open
        GoTo Exit_btn_saijikko_Click
    End If
        
    '取り込むことにしちゃったら、再実行ボタンはもう要らない。
    'Me![btn_saijikko].Visible = False
    
    If Me![frm_text_select] = 6 Then
        MsgBox ("処理終了しました。")
        GoTo Exit_btn_saijikko_Click
    End If
    
    '置換えるを選択していた場合、既存データの削除やクリア
    If intOkikae = 1 Then
        With cmd
            .CommandText = "usp_R3_koubai_hacchu_import_tourokuzumi_clear"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = Me![frm_text_select]
            .Execute
        End With
    End If
    
    Select Case Me![frm_text_select]
        Case 1, 5
            '通常の取込データ追加、番号更新
            With cmd
                .CommandText = "usp_R3_koubai_hacchu_import_hacchu_update"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
            
            '支給包装材のもろもろ追加、更新
            If Me![shanghai_chotatsu] <> 2 Then
                With cmd
                    .CommandText = "usp_R3_koubai_hacchu_import_hosozai_update"
                    .CommandType = adCmdStoredProc
                    .Parameters.Refresh
                    .Parameters(1) = Me![login_code]
                    .Execute
                End With
            Else
                If Len(strInvoiceNo) > 0 Then
                    '上海以外の購買発注のINVOICE_NO更新
                    With cmd
                        .CommandText = "usp_R3_koubai_hacchu_import_hacchu_update_invoice_no"
                        .CommandType = adCmdStoredProc
                        .Parameters.Refresh
                        .Parameters(1) = Me![login_code]
                        .Parameters(2) = strInvoiceNo
                        .Execute
                    End With
                End If
            End If

        Case 4
            '抜き取り検査情報更新
            With cmd
                .CommandText = "usp_R3_koubai_hacchu_import_nukitori_update"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
        Case 5
            '大田２F発注更新
            'With cmd
            '    .CommandText = "usp_R3_koubai_hacchu_import_ohta2F_update"
            '    .CommandType = adCmdStoredProc
            '    .Parameters.Refresh
            '    .Parameters(1) = Me![login_code]
            '    .Execute
            'End With

    End Select
    
    MsgBox ("処理終了しました。")
    
    Set cmd = Nothing
    
    On Error Resume Next
    Call syori_log_create(intMenuNo, intMenuNoSeq, Me![REF_NO], "処理選択=" & Me![frm_text_select] & " 取り込み")
    On Error GoTo 0
                        
    Exit Sub
    
Exit_btn_saijikko_Click:
    
    Set cmd = Nothing
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_1_R3購買発注取込")
    If intShanghaiChotatsu = 2 Then
        intMenuNo = 201
    Else
        intMenuNo = 104
    End If
    intMenuNoSeq = 0

End Sub

Private Sub btn_shanghai_chotatsu_Click()

    'If Me![btn_shanghai_chotatsu].Caption = "上海" Then
    '    Me![btn_shanghai_chotatsu].Caption = "調達"
    '    intShanghaiChotatsu = 2
    '    intMenuNo = 201
    'Else
    '    If Me![btn_shanghai_chotatsu].Caption = "調達" Then
    '        Me![btn_shanghai_chotatsu].Caption = "貿易"
    '        intShanghaiChotatsu = 3
    '        intMenuNo = 104
    '    Else
    '        Me![btn_shanghai_chotatsu].Caption = "上海"
    '        intShanghaiChotatsu = 1
    '        intMenuNo = 104
    '    End If
    'End If
    
    'strFormName = "F_1_R3購買発注取込"
    
    'Call next_form_open

End Sub

Private Sub Form_Open(Cancel As Integer)
    
    DoCmd.Maximize
    
    If intShanghaiChotatsu = 2 Then
        intMenuNo = 201
    Else
        intMenuNo = 104
    End If
    intMenuNoSeq = 0
    
    'Set cmd.ActiveConnection = conn
    
    'With cmd
    '    .CommandText = "usp_koumoku_contents_get"
    '    .CommandType = adCmdStoredProc
    '    .Parameters.Refresh
    '    .Parameters(1) = "FIL"
    '    .Parameters(2) = 2
    '    .Execute
    'End With
    
    'Me![PASS] = cmd.Parameters(3)
    'Me![text_name] = cmd.Parameters(4)
    strInvoiceNo = ""
    
    'Set cmd = Nothing
    
    If intShanghaiChotatsu = 2 Then
        Me![btn_invoice_meisai].Enabled = False
        Me![btn_ohta_2f].Enabled = False
        Me![opt_nukitori_kensa].Enabled = False
        Me![opt_ohta_2F_hacchu].Enabled = False
        Me![btn_invoice_header].Caption = "配送手配"
        Me![btn_invoice_header].ForeColor = "10040115"
        Me![REF_NO].BackColor = "16777215"
        Me![lbl_ref_no].Caption = "INVOICE_NO"
    Else
        Me![btn_invoice_meisai].Enabled = True
        Me![btn_ohta_2f].Enabled = True
        Me![opt_nukitori_kensa].Enabled = True
        Me![opt_ohta_2F_hacchu].Enabled = True
        Me![btn_invoice_header].Caption = "INVヘッダ"
        Me![btn_invoice_header].ForeColor = "8388672"
        Me![REF_NO].BackColor = "12632256"
        Me![lbl_ref_no].Caption = "REF_NO"
    End If
    'Me![btn_saijikko].Visible = False
    
End Sub

Private Sub frm_text_select_Click()

    Select Case Me![frm_text_select]
        Case 1, 5
            Me![text_name] = "ZT053.txt"
            Me![frm_text_type] = 2
            Me![haiso_yotei_ymd].Visible = False
            Me![lbl_ref_no].Caption = "REF_NO"
            Me![REF_NO].Enabled = False
            Me![REF_NO].BackColor = "12632256"
        Case 4
            Me![text_name] = "抜取検査品.xls"
            Me![frm_text_type] = 1
            Me![haiso_yotei_ymd].Visible = False
            Me![lbl_ref_no].Caption = "INVOICE_NO"
            Me![REF_NO].Enabled = True
            Me![REF_NO].BackColor = "16777215"
    End Select
    
End Sub

Private Sub shanghai_chotatsu_AfterUpdate()

    Select Case Me![shanghai_chotatsu]
        Case 1
            intShanghaiChotatsu = 1
            intMenuNo = 104
        Case 2
            intShanghaiChotatsu = 2
            intMenuNo = 201
        Case 3
            intShanghaiChotatsu = 3
            intMenuNo = 104
        Case Else
            intShanghaiChotatsu = 1
            intMenuNo = 104
    End Select
    
    strFormName = "F_1_R3購買発注取込"
    Call next_form_open

End Sub


Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Select Case Me![shanghai_chotatsu]
        Case 1
            intShanghaiChotatsu = 1
        Case 2
            intShanghaiChotatsu = 2
        Case 3
            intShanghaiChotatsu = 3
        Case Else
            intShanghaiChotatsu = 1
    End Select
    
    strInvoiceNo = ""
    If strFormName = "F_1_R3購買発注取込" Then
        strOpenMode = "O"
    Else
        strOpenMode = "OC"
    End If
    
    Call form_open_close("F_1_R3購買発注取込", strFormName, strOpenMode, "NULL")
    
    If strFormName = "F_1_取込内容確認_MAIN" Then
        If Len(strSubFormName) > 0 Then
            Forms(strFormName)![F_1_取込内容確認_SUB].SourceObject = strSubFormName
        End If
    End If
    
End Sub

