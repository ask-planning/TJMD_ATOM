Attribute VB_Name = "Form_F_1_上海輸入_INVOICE取込"
Attribute VB_Base = "0{BBC826F7-780F-41CD-9B66-EB4675D64100}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub bnt_menu_Click()

    'メインメニューに戻る
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_invoice_header_Click()

    'INVOICEヘッダ入力へ
    strFormName = "F_1_上海輸入_INVOICE_HEADER"
    intMenuNo = 103
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_invoice_meisai_Click()

    'INVOICE内容確認へ
    strFormName = "F_1_上海輸入_INVOICE_MAIN"
    intMenuNo = 102
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_import_click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![PASS]) Then
        MsgBox ("取り込み元のパスを入力してください。")
        Exit Sub
    End If
    If IsNull(Me![text_name]) Then
        MsgBox ("取り込み元のテキスト名を入力してください。")
        Exit Sub
    End If
    If Not IsNull(Me![text_name]) Then
        If Right(Me![text_name], 4) <> ".xls" Then
            Me![text_name] = Me![text_name] & ".xls"
        End If
    End If
    
    If IsNull(Me![chk_customer_kbn]) Then
        MsgBox ("発注先を選択してください。")
        Exit Sub
    End If
    
    If IsNull(Me![currency_type]) Or Me![currency_type] = 0 Then
        MsgBox ("通貨を選択してください。")
        Exit Sub
    End If
    
    'T_CHOTATSU-39 上海輸入INV+PL取込み画面：通貨との紐づけ
    If Left(Me![text_name], 5) = "SJ-TJ" Then
        If Me![currency_type] <> 2 Then
            MsgBox ("INVOICEと通貨が一致しません。")
            Exit Sub
        End If
    End If
    
    If Left(Me![text_name], 5) = "SJ-TR" Then
        If Me![currency_type] <> 1 Then
            MsgBox ("INVOICEと通貨が一致しません。")
            Exit Sub
        End If
    End If
    
    If Left(Me![text_name], 5) = "YJ-TR" Then
        If Me![currency_type] <> 1 Then
            MsgBox ("INVOICEと通貨が一致しません。")
            Exit Sub
        End If
    End If
    
    Dim strChkTableName As String
    Dim strImportTableName As String
    Dim strSyori As String
    Dim strOKFlg As String
    Dim intRCount As Integer
    Dim intLoopFlg As Integer
    Dim intLoopCount As Integer
    Dim strCartonNoFrom As String
    Dim strCartonNoTo As String
    
    Set cmd.ActiveConnection = conn
    
    DoCmd.Hourglass True
    
    '取込シートにより､ループ処理の開始､終了を決める
    'If Me![frm_text_select] = 0 Then
        intLoopFlg = 1
        intLoopCount = 2
    'Else
    '    intLoopFlg = Me![frm_text_select]
    '    intLoopCount = Me![frm_text_select]
    'End If

    'LOOPフラグ＝０はINVOICE、LOOPフラグ＝１はパッキング
    Do Until intLoopFlg > intLoopCount
        strSyori = "tempdelete"
    
        'IMPORT用テンポラリファイルの削除（エラーは無視）と、テーブル名を変数に取っておく
        On Error Resume Next
        
        t = Len(Me![text_name]) - 4
        
        If intLoopFlg = 1 Then
            strChkTableName = "tbl_t_shanghai_invoice_import_check_temp"
            strImportTableName = "tbl_t_shanghai_invoice_import_temp_" & Me![login_code]
            strFileName = Me![PASS] & "\" & Mid(Me![text_name], 1, t) & "_INV.xls"
        Else
            strChkTableName = "tbl_t_shanghai_packing_import_check_temp"
            'strImportTableName = "tbl_t_shanghai_packing_import_temp"
            strImportTableName = "tbl_t_shanghai_invoice_import_temp"
            strFileName = Me![PASS] & "\" & Mid(Me![text_name], 1, t) & "_PAC.xls"
        End If
        
        '取込み用テンポラリの削除　⇒　SQL SERVER2014では出来ないので止め
        'DoCmd.DeleteObject acTable, strImportTableName
        
        '取込み用テンポラリのクリア
        With cmd
            .CommandText = "DELETE " & strImportTableName & " FROM " & strImportTableName & " "
            .CommandType = adCmdText
            .Execute
        End With
        
        On Error GoTo 0
        
        strSyori = "tempdelete_end"
    
        'EXCELの取り込み、エラーのときは多分フォーマット違いの為、メッセージを表示する
        'On Error GoTo Err_btn_import_Click
        
        If intLoopFlg = 1 Then
            strSyori = "invoice_import"
        Else
            strSyori = "packing_import"
        End If
        
        'ヘッダーあり
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, strImportTableName, strFileName, True, "A:R"
        
        On Error GoTo 0
    
        With cmd
            .CommandText = "DELETE " & strChkTableName & " FROM " & strChkTableName & " WHERE pk_login_code = '" & Me![login_code] & "' "
            .CommandType = adCmdText
            .Execute
        End With
        
    
        'IMPORTテーブルの読み込みと、チェック用テンポラリのクリアとレコードセット取得
        strSql = "SELECT * FROM " & strImportTableName
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
        End With
        Set rs1 = cmd.Execute(intRcnt)
            
        'On Error GoTo Err_btn_import_Click
        'IMPORTテーブルに中身があるか？
        'IMPORTテーブルはフィールド名が当てにならないので、フィールド番号でチェック
        
        If intRcnt <> 0 Then
            rs1.MoveFirst
            
            'If rs1.Fields.Count <> 14 Then
            '    rs1.Close
            '    If intLoopFlg = 1 Then
            '        strSyori = "invoice_field_count_err"
            '    Else
            '        strSyori = "packing_field_count_err"
            '    End If
            '    GoTo Err_btn_import_Click
            'End If
            
            
            If intLoopFlg = 1 Then
                strSyori = "invoice_format_err"
            Else
                strSyori = "packing_format_err"
            End If
            
            'カートン番号が前と同じ場合は空白にする、まず比較用キー
            If intLoopFlg = 2 Then
                '開始／終了カートン番号
                strCartonNoFrom = ""
                strCartonNoTo = ""
            End If
                
            Do Until rs1.EOF
            
                'まず、『NG』をセットして、OKだったら『OK』をセットする
                strOKFlg = "NG"
                If Len(rs1.Fields(0).Value) > 0 Then
                        strOKFlg = "OK"
                End If
                
                If intLoopFlg = 1 Then
                    strSql = "INSERT INTO tbl_t_shanghai_invoice_import_check_temp ( pk_login_code, line_seq, so_no, so_no_seq, "
                    strSql = strSql & "shanghai_code, shanghai_code_text, kikaku, part_sort, suryo_tani, suryo, unit_price, amount, "
                    strSql = strSql & "dummy_no1, dummy_no2, tajima_po_no, R3_koubai_denpyo_no, invoice_no, bl_date, syukka_hoho, "
                    strSql = strSql & "R3_hinmoku_code, tsuka, tajima_po_no_original, amount_text, err_status ) "
                Else
                    strSql = "INSERT INTO tbl_t_shanghai_packing_import_check_temp ( pk_login_code, true_or_false, invoice_no, "
                    strSql = strSql & "tajima_po_no, shanghai_code, suryo, unit_price, amount, carton_no_from, carton_no_to, net_weight, "
                    strSql = strSql & "gross_weight, m3, bl_date, syukka_hoho , R3_hinmoku_code, tajima_po_no_original, err_status ) "
                End If
                strSql = strSql & "Values ( '" & Me![login_code] & "'"
                
                '『OK』だったらチェック用テンポラリに取り込み
                If strOKFlg = "OK" Then
                    
                    For n = 0 To rs1.Fields.Count - 1
                        'If n <> 0 Then
                            strSql = strSql & ", "
                        'End If
                        
                        If n = 3 Then
                            '上海コード編集
                            If IsNumeric(rs1.Fields(3).Value) Then
                                strSql = strSql & "'" & Right("00000" & rs1.Fields(3).Value, 5) & "'"
                            Else
                                strSql = strSql & "'" & LTrim(RTrim(rs1.Fields(3).Value)) & "'"
                            End If
                        Else
                        
                            'PACKING LISTのカートンNOのFROM～TO
                            If intLoopFlg = 2 And (n = 7 Or n = 8) Then
                                If n = 7 Then
                                    If IsNull(rs1.Fields(7).Value) Or Len(LTrim(RTrim(rs1.Fields(7).Value))) <= 0 Then
                                        If Len(strCartonNoFrom) > 0 Then
                                            strSql = strSql & "'" & strCartonNoFrom & "'"
                                        Else
                                            strSql = strSql & "null"
                                        End If
                                    Else
                                        strSql = strSql & "'" & LTrim(RTrim(rs1.Fields(7).Value)) & "'"
                                    End If
                                End If
                                If n = 8 Then
                                    If Not IsNull(rs1.Fields(8).Value) And Len(LTrim(RTrim(rs1.Fields(8).Value))) > 0 Then
                                        If strCartonNoTo = LTrim(RTrim(rs1.Fields(8).Value)) Then
                                            strSql = strSql & "null"
                                        Else
                                            strSql = strSql & "'" & LTrim(RTrim(rs1.Fields(8).Value)) & "'"
                                        End If
                                    Else
                                        strSql = strSql & "null"
                                    End If
                                End If
                            Else
                                If n = 14 And intLoopFlg = 1 Then
                                'Debug.Print rs1.Fields(14).Value
                                    If IsNull(rs1.Fields(14).Value) Or Len(rs1.Fields(14).Value) <= 0 Then
                                        strSql = strSql & "'9999999999'"
                                    Else
                                        strSql = strSql & "'" & rs1.Fields(14).Value & "'"
                                    End If
                                Else
                                    If Len(rs1.Fields(n).Value) > 0 And Len(LTrim(RTrim(rs1.Fields(n).Value))) > 0 Then
                                        If intLoopFlg = 1 And n = 0 Then
                                            strSql = strSql & LTrim(RTrim(rs1.Fields(n).Value))
                                        Else
                                            strSql = strSql & "'" & LTrim(RTrim(rs1.Fields(n).Value)) & "'"
                                        End If
                                    Else
                                        strSql = strSql & "null"
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    'R3品目コード
                    If IsNumeric(rs1.Fields(3).Value) Then
                        strSql = strSql & ", '" & Right("000000000000000000" & LTrim(RTrim(rs1.Fields(3).Value)), 18) & "'"
                    Else
                        strSql = strSql & ", '" & LTrim(RTrim(rs1.Fields(3).Value)) & "'"
                    End If
                    
                    'PO_ORIGINAL
                    If intLoopFlg = 1 Then
                        If Me![currency_type] = 2 Then
                            strSql = strSql & ", 'JPY'"
                        Else
                            strSql = strSql & ", 'RMB'"
                        End If
                        
                        If Len(rs1.Fields(13).Value) > 0 And Len(LTrim(RTrim(rs1.Fields(13).Value))) > 0 Then
                            strSql = strSql & ", '" & LTrim(RTrim(rs1.Fields(13).Value)) & "'"
                        Else
                            strSql = strSql & ", null"
                        End If
                        strSql = strSql & ", '" & rs1.Fields(10) & "'"
                    Else
                        If Len(rs1.Fields(2).Value) > 0 And Len(LTrim(RTrim(rs1.Fields(2).Value))) > 0 Then
                            strSql = strSql & ", '" & LTrim(RTrim(rs1.Fields(2).Value)) & "'"
                        Else
                            strSql = strSql & ", null"
                        End If
                    End If
                    strSql = strSql & ", 0 ) "
                    
                    Debug.Print strSql
                    With cmd
                        .CommandText = strSql
                        .CommandType = adCmdText
                        .Execute
                    End With
                    
                    'PACKING_LISTのFROM TOの空白の場合用
                    If intLoopFlg = 2 Then
                        If Not IsNull(rs1.Fields(7).Value) And Len(LTrim(RTrim(rs1.Fields(7).Value))) > 0 Then
                            strCartonNoFrom = LTrim(RTrim(rs1.Fields(7).Value))
                        End If
                        If Not IsNull(rs1.Fields(8).Value) And Len(LTrim(RTrim(rs1.Fields(8).Value))) > 0 Then
                            strCartonNoTo = LTrim(RTrim(rs1.Fields(8).Value))
                        End If
                    End If
                        
                End If
                rs1.MoveNext
            Loop
        End If
        strSyori = "format_err_end"
        rs1.Close
        intLoopFlg = intLoopFlg + 1
        
    Loop
    
    
    'If Me![frm_text_select] = 2 Then
    '    ReturnValue = MsgBox("INVOICEデータも作成しますか？" & vbLf & "作成（はい）／PACKINGだけ（いいえ）", 4)
    '    If ReturnValue = vbYes Then
    '        Me![frm_text_select] = 0
            
    '        With cmd
    '            .CommandText = "usp_shanghai_import_invoice_check_temp_create"
    '            .CommandType = adCmdStoredProc
    '            .Parameters.Refresh
    '            .Parameters(1) = Me![login_code]
    '            .Execute
    '        End With
    '    End If
    'End If
    
    Set rs1 = Nothing
    Set cmd = Nothing
    
    On Error GoTo 0
    
    DoCmd.Hourglass False
    Call btn_saijikko_Click
    Exit Sub

Exit_btn_import_Click:
    
    DoCmd.Hourglass False
    Set rs1 = Nothing
    Set cmd = Nothing
    Exit Sub

Err_btn_import_Click:

    DoCmd.Hourglass False
    
    'エラーフラグの内容でメッセージを表示する
    Select Case strSyori
    
        Case "tempdelete"
            Resume Next
            
        Case "invoice_field_count_err"
            MsgBox ("INVOICEのS列以降にデータが存在します。Ｏ列以降を削除してください。")
            
        Case "packing_field_count_err"
            MsgBox ("PACKINGのＯ列以降にデータが存在します。Ｏ列以降を削除してください。")
            
        Case "invoice_format_err"
            MsgBox ("INVOICEのフォーマットが違います。")
            
        Case "packing_format_err"
            MsgBox ("PACKINGのフォーマットが違います。")
            
        Case "invoice_import"
            MsgBox ("INVOICEのEXCELの取り込み処理のエラーです。テキスト名、EXCELの内容等を確認してください。")
            
        Case "packing_import"
            MsgBox ("PACKINGのEXCELの取り込み処理のエラーです。テキスト名、EXCELの内容等を確認してください。")
            
        Case Else
            MsgBox ("取り込み処理のエラーです。テキスト名等を確認してください。")
            
    End Select
    
End Sub

Private Sub btn_kakunin_Click()

    strFormName = "F_1_取込内容確認_MAIN"
    'If Me![frm_text_select] = 2 Then
    '    strSubFormName = "F_1_取込内容確認_SUB_PACKING"
    'Else
        strSubFormName = "F_1_取込内容確認_SUB_INVOICE"
    'End If
    Call next_form_open
            
End Sub

Private Sub btn_saijikko_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![chk_customer_kbn]) Then
        MsgBox ("発注先を選択してください。")
        Exit Sub
    End If
    
    If IsNull(Me![currency_type]) Or Me![currency_type] = 0 Then
        MsgBox ("通貨を選択してください。")
        Exit Sub
    End If
    
    DoCmd.GoToControl "btn_import"
    
    Dim intMitouroku As Integer
    Dim strOKFlg As String
    
    Dim intLen As Integer
    Dim intMidLen As Integer
    Dim intKahenLen As Integer
    Dim intKaishiIchi As Integer
    
    Dim strTempTableName As String
    Dim strDisplayOrder As String
    Dim strKeyOrder As String
    
    Dim i As Integer
    Dim intLoopFlg As Integer
    Dim intLoopCount As Integer
    Dim intCountInvoice As Integer
    Dim intCountPacking As Integer
    
    Dim strTblInvoiceNo() As String
    Dim intTblInvoiceKbn() As Integer
    
    Dim intInvoiceDoubleCount As Integer
    
    Dim intOkikaeFlg As Integer
    Dim intCartonQty As Integer
    Dim intSyukkaHoho As Integer
    Dim datBLDate As Date
    
    Set cmd.ActiveConnection = conn
    
    'R3得意先コードGET(parameters(1) ⇒　1：得意先コードから、2：customer_kbnから）
    intCustomerKbn = Me![chk_customer_kbn]
    If Me![chk_customer_kbn] = 1 Then
        If Me![currency_type] = 1 Then
            intCustomerKbn = 3
        End If
    End If
    
    strShiireCode = ""
    With cmd
        .CommandText = "usp_POEM_tokuisaki_master_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = 2
        .Parameters(2) = "NON"
        .Parameters(3) = intCustomerKbn
        .Execute
    End With
    If Not IsNull(cmd.Parameters(14)) Then
        strShiireCode = cmd.Parameters(14)
    End If
        
    
    strSubFormName = ""
    
    '来月以降のINVOICEがあるか
    intCountInvoice = 0
    'If Me![frm_text_select] <> 1 Then
        With cmd
            .CommandText = "usp_shanghai_import_count_get_next_month"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = DateSerial(Year(Date), Month(Date) + 1, 1)
            .Execute
        End With
            
        'INVOICEのみ以外（PACKING、または同時取り込みの場合）、PACKINGデータがあるか
        If Not IsNull(cmd.Parameters(3)) Then
            intCountInvoice = cmd.Parameters(3)
        End If
    'End If
    If intCountInvoice > 0 Then
        ReturnValue = MsgBox("今月以降のBL DATEのINVOICEです。" & vbLf & "取り込みを中断しますか？　中断（はい）／続行（いいえ）", vbYesNo)
        If ReturnValue = vbYes Then
            GoTo Exit_btn_saijikko_Click
        End If
    End If
    
    '登録済みの有無を調べるだけなので、カウントだったり、あったら１足すだけだったり
    intCountInvoice = 0
    intCountPacking = 0
    'If Me![frm_text_select] = 1 Then
    '    strInvoiceNo = Me![INVOICE_NO]
    'Else
    '    strInvoiceNo = "NULL"
    'End If
    '取込み済みINVOICE/PACKINGがあるか
    With cmd
        .CommandText = "usp_shanghai_import_tourokuzumi_count"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With

    '同時取り込み、またはPACKING取り込みの場合、INVOICEとPACKINの両方
    If Not IsNull(cmd.Parameters(2)) Then
        intCountInvoice = cmd.Parameters(2)
    End If
    If Not IsNull(cmd.Parameters(3)) Then
        intCountPacking = cmd.Parameters(3)
    End If

    '登録済みがあったら、配列を再定義して登録済み情報を取り込む
    If intCountInvoice + intCountPacking <> 0 Then
    
        ReDim strTblInvoiceNo(intCountInvoice + intCountPacking) As String
        ReDim intTblInvoiceKbn(intCountInvoice + intCountPacking) As Integer
        
        n = 0
            
        'If Me![frm_text_select] = 0 Then            '同時取込み
            strSql = "SELECT dbo.View_shanghai_invoice_no_invoice.pk_invoice_no FROM dbo.View_shanghai_invoice_no_invoice "
            strSql = strSql & "INNER JOIN dbo.View_shanghai_import_invoice_no ON dbo.View_shanghai_invoice_no_invoice."
            strSql = strSql & "pk_invoice_no = dbo.View_shanghai_import_invoice_no.INVOICE_NO "
            strSql = strSql & "WHERE View_shanghai_import_invoice_no.pk_login_code = '" & Me![login_code] & "' "
            strSql = strSql & "GROUP BY dbo.View_shanghai_invoice_no_invoice.pk_invoice_no "
            
            'Debug.Print strSql
            
            With cmd
                .CommandText = strSql
                .CommandType = adCmdText
            End With
            Set rs1 = cmd.Execute(intRcnt)
            If intRcnt <> 0 Then
                rs1.MoveFirst
                Do Until rs1.EOF
                    strTblInvoiceNo(n) = rs1![pk_invoice_no].Value
                    intTblInvoiceKbn(n) = 1
                    n = n + 1
                    rs1.MoveNext
                Loop
            End If
            rs1.Close
        'End If
        
        'If Me![frm_text_select] = 1 Then            'INVOICEのみ取込み
        '    strSql = "SELECT dbo.View_shanghai_invoice_no_invoice.pk_invoice_no FROM dbo.View_shanghai_invoice_no_invoice "
        '    strSql = strSql & "WHERE dbo.View_shanghai_invoice_no_invoice.pk_invoice_no = '" & Me![INVOICE_NO] & "' "
            
        '    With cmd
        '        .CommandText = strSql
        '        .CommandType = adCmdText
        '    End With
        '    Set rs1 = cmd.Execute(intRcnt)
        '    If intRcnt <> 0 Then
        '        rs1.MoveFirst
        '        Do Until rs1.EOF
        '            strTblInvoiceNo(n) = rs1![pk_invoice_no].Value
        '            intTblInvoiceKbn(n) = 1
        '            n = n + 1
        '            rs1.MoveNext
        '        Loop
        '    End If
        '    rs1.Close
        'End If
            
        'If Me![frm_text_select] <> 1 Then           '同時、またはPACKING取込み
            strSql = "SELECT dbo.View_shanghai_invoice_no_packing.pk_invoice_no FROM dbo.View_shanghai_invoice_no_packing "
            strSql = strSql & "INNER JOIN dbo.View_shanghai_import_invoice_no_packing ON dbo.View_shanghai_invoice_no_packing."
            strSql = strSql & "pk_invoice_no = dbo.View_shanghai_import_invoice_no_packing.invoice_no "
            strSql = strSql & "WHERE dbo.View_shanghai_import_invoice_no_packing.pk_login_code = '" & Me![login_code] & "' "
            
            With cmd
                .CommandText = strSql
                .CommandType = adCmdText
            End With
            Set rs1 = cmd.Execute(intRcnt)
            If intRcnt <> 0 Then
                rs1.MoveFirst
                Do Until rs1.EOF
                    strTblInvoiceNo(n) = rs1![pk_invoice_no].Value
                    intTblInvoiceKbn(n) = 2
                    n = n + 1
                    rs1.MoveNext
                Loop
            End If
            rs1.Close
        'End If
    End If
    
    '取り込むテキストの表示用INVOICE_NOの編集（INVOICEのみの取り込みの場合はやらない）
    strDisplayOrder = ""
    'If Me![frm_text_select] <> 1 Then
        strSql = "SELECT INVOICE_NO FROM tbl_t_shanghai_invoice_import_check_temp WHERE [INVOICE_NO] is not null "
        strSql = strSql & "AND pk_login_code = '" & Me![login_code] & "' GROUP BY INVOICE_NO ORDER BY INVOICE_NO "
        'intLen = 8
        
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
        End With
        Set rs1 = cmd.Execute(intRcnt)
        
        strInvoiceNoMain = ""
        If intRcnt <> 0 Then
            rs1.MoveFirst
            '貿易用INVOICEはINVOICE_NOの商品群が/で始まり、物流は番号の後に/無しで商品群
            If Me![shanghai_chotatsu] = 3 Then
                intKaishiIchi = InStr(rs1![INVOICE_NO].Value, "/")
                strInvoiceNoMain = Left(rs1![INVOICE_NO].Value, intKaishiIchi - 1)
                intLen = intKaishiIchi - 1
            Else
                'If IsNumeric(Mid(rs1![INVOICE_NO].Value, 10, 1)) Then
                '    strInvoiceNoMain = Left(rs1![INVOICE_NO].Value, 10)
                'Else
                '    strInvoiceNoMain = Left(rs1![INVOICE_NO].Value, 9)
                'End If
                '通貨記号があったら、先頭から９バイト、無かったら８バイト
                If IsNumeric(Mid(rs1![INVOICE_NO].Value, 5, 1)) Then
                    strInvoiceNoMain = Left(rs1![INVOICE_NO].Value, 8)
                    intLen = 8
                Else
                    strInvoiceNoMain = Left(rs1![INVOICE_NO].Value, 9)
                    intLen = 9
                End If
            End If
            strKeyOrder = Left(rs1.Fields(0).Value, intLen)
            Do Until rs1.EOF
                'Debug.Print rs1![ORDER_NO].Value
                If Len(strDisplayOrder) <= 0 Then
                    strDisplayOrder = rs1.Fields(0).Value
                Else
                    If Left(rs1.Fields(0).Value, intLen) <> strKeyOrder Then
                        strDisplayOrder = strDisplayOrder & "/" & rs1.Fields(0).Value
                        strKeyOrder = Left(rs1.Fields(0).Value, intLen)
                    Else
                        intKahenLen = Len(rs1.Fields(0).Value) - intLen
                        If Me![shanghai_chotatsu] = 3 Then
                            strDisplayOrder = strDisplayOrder & Mid(rs1.Fields(0), intLen + 1, intKahenLen)
                        Else
                            strDisplayOrder = strDisplayOrder & "/" & Mid(rs1.Fields(0), intLen + 1, intKahenLen)
                        End If
                    End If
                End If
                rs1.MoveNext
            Loop
        End If
        rs1.Close
    'End If
    
    '登録済みの場合、置き換えるかどうか
    'このIFは、（同時取り込みまたはPACKING取り込みでPACKINGがある）、か、（同時取り込みまたはINVOICE取り込みでINVOICEがある）
    intOkikaeFlg = 0
    'If (Me![frm_text_select] <> 1 And intCountPacking <> 0) Or (Me![frm_text_select] <> 2 And intCountInvoice <> 0) Then
    If (intCountPacking <> 0) Or (intCountInvoice <> 0) Then
        ReturnValue = MsgBox("登録済みのINVOICEです。" & vbLf & "置換（はい）／追加（いいえ）／キャンセル（キャンセル）", vbYesNoCancel)
        If ReturnValue = vbCancel Then
            GoTo Exit_btn_saijikko_Click
        Else
            If ReturnValue = vbYes Then
            
                intOkikaeFlg = 1
                
                '置換えの場合、登録済みINVOICE、PACKINGデータを削除する
                
                'HEADERは１つ(ただし、INVOICEのみの取り込みの場合は削除しない）
                'If Me![frm_text_select] <> 1 Then
                    strSql = "DELETE tbl_t_shanghai_invoice_header FROM tbl_t_shanghai_invoice_header "
                    strSql = strSql & "WHERE pk_invoice_no_main = '" & strInvoiceNoMain & "' "
                    With cmd
                        .CommandText = strSql
                        .CommandType = adCmdText
                        .Execute
                    End With
                'End If
                        
                For n = 0 To intCountInvoice + intCountPacking - 1
                    If Len(strTblInvoiceNo(n)) > 0 Then
                        
                        'サブヘッダ(ただし、INVOICEのみの取り込みの場合は削除しない）
                        'If Me![frm_text_select] <> 1 Then
                            strSql = "DELETE tbl_t_shanghai_invoice_header_sub FROM tbl_t_shanghai_invoice_header_sub "
                            strSql = strSql & "WHERE pk_invoice_no = '" & strTblInvoiceNo(n) & "' "
                            With cmd
                                .CommandText = strSql
                                .CommandType = adCmdText
                                .Execute
                            End With
                        'End If
                        
                        If intTblInvoiceKbn(n) = 1 Then
                            strSql = "DELETE tbl_t_shanghai_invoice_meisai FROM tbl_t_shanghai_invoice_meisai "
                            strSql = strSql & " WHERE pk_invoice_no = '" & strTblInvoiceNo(n) & "' "
                            With cmd
                                .CommandText = strSql
                                .CommandType = adCmdText
                                .Execute
                            End With
                        Else
                            'If Me![frm_text_select] <> 1 Then
                                strSql = "DELETE tbl_t_shanghai_packing_meisai FROM tbl_t_shanghai_packing_meisai "
                                strSql = strSql & " WHERE pk_invoice_no = '" & strTblInvoiceNo(n) & "' "
                                With cmd
                                    .CommandText = strSql
                                    .CommandType = adCmdText
                                    .Execute
                                End With
                            'End If
                        End If
                    End If
                Next
            Else
                If ReturnValue = vbNo Then
                    intOkikaeFlg = 2
                End If
            End If
        End If
    End If
    
    Me![REF_NO] = strDisplayOrder
    
    'INVOICEのみの場合、登録済みのPACKINGが無かったらエラー、あったらテンポラリに入れておく
    'If Me![frm_text_select] = 1 Then
        
    '    If intCountPacking = 0 Then
    '        MsgBox ("入力のINVOICE_NOでは取込済みPACKINGデータがありません。" & vbLf & "PACKINGデータを先に取り込んでください。")
    '        GoTo Exit_btn_saijikko_Click
    '    Else
    '        With cmd
    '            .CommandText = "usp_shanghai_import_packing_check_temp_create"
    '            .CommandType = adCmdStoredProc
    '            .Parameters.Refresh
    '            .Parameters(1) = Me![login_code]
    '            .Parameters(2) = Me![INVOICE_NO]
    '            .Execute
    '        End With
    '    End If
        
    'End If
        
    
    '取込データを読み込んで、各種チェックを行う
    Me![btn_saijikko].Visible = True
    
    '取込シートにより､ループ処理の開始､終了を決める
    'If Me![frm_text_select] = 0 Then
        intLoopFlg = 1
        intLoopCount = 2
    'Else
    '    intLoopFlg = Me![frm_text_select]
    '    intLoopCount = Me![frm_text_select]
    'End If
    
    strOKFlg = "OK"
    intErrNo = 0
    With cmd
        .CommandText = "usp_shanghai_import_check_err_clear"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    '１POで複数のINVOICE_NOがあるか？(貿易のINVOICE以外）
    If Me![shanghai_chotatsu] <> 3 Then
        intInvoiceDoubleCount = 0
        With cmd
            .CommandText = "usp_shanghai_import_jyufuku_check"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Execute
        End With
        If Not IsNull(cmd.Parameters(2)) Then
            intInvoiceDoubleCount = cmd.Parameters(2)
        End If
        
        If intInvoiceDoubleCount <> 0 Then
            strOKFlg = "NG"
            MsgBox ("複数INVOICEに分かれているPOがあります。" & vbLf & "POを変更し、再度実行してください。")
            strSubFormName = "F_1_取込内容確認_SUB_INVOICE"
            GoTo Err_btn_saijikko_Click
        End If
    End If
    
    Do Until intLoopFlg > intLoopCount
               
        If intLoopFlg = 1 Then
            With cmd
                .CommandText = "usp_shanghai_import_err_check_invoice"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
            'warnning
            If Not IsNull(cmd.Parameters(3)) And cmd.Parameters(3) <> 0 Then
                If intErrNo <> 1 Then
                    intErrNo = 2
                End If
            End If
            'error
            If Not IsNull(cmd.Parameters(2)) And cmd.Parameters(2) <> 0 Then
                intErrNo = 1
            End If
        Else
            With cmd
                .CommandText = "usp_shanghai_import_err_check_packing"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Execute
            End With
            'warnning
            If Not IsNull(cmd.Parameters(3)) And cmd.Parameters(3) <> 0 Then
                If intErrNo <> 1 Then
                    intErrNo = 2
                End If
            End If
            'error
            If Not IsNull(cmd.Parameters(2)) And cmd.Parameters(2) <> 0 Then
                intErrNo = 1
            End If
        End If
                
        intLoopFlg = intLoopFlg + 1
        
    Loop
    
    'エラーがない、またはWARNNINGだけの場合、金額のチェックへ
    If intErrNo <> 1 Then
        With cmd
            .CommandText = "usp_shanghai_import_check_amount"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Execute
        End With
        'warnning
        If Not IsNull(cmd.Parameters(3)) And cmd.Parameters(3) <> 0 Then
            If intErrNo <> 1 Then
                intErrNo = 2
            End If
        End If
        'error
        If Not IsNull(cmd.Parameters(2)) And cmd.Parameters(2) <> 0 Then
            intErrNo = 1
        End If
    End If
    
    'エラーがあったらメッセージを表示する
    Select Case intErrNo
        Case 1
            MsgBox ("エラーデータがあります。" & vbLf & "EXCELを確認してください。")
            strOKFlg = "NG"
        Case 2
            ReturnValue = MsgBox("WARNINGデータがあります、このまま取り込みますか？" & vbLf & "取り込み（はい）／キャンセル（いいえ）", 4)
            If ReturnValue = vbNo Then
                strOKFlg = "NG"
            End If
        Case Else
            ReturnValue = MsgBox("内容の確認を行いますか？　確認（はい）／取込（いいえ）", 4 + vbDefaultButton2)
            If ReturnValue = vbYes Then
                strOKFlg = "NG"
            End If
    End Select
            
    'どこかで内容確認をすると選んでいたら、このフラグに「NG」がセットされている。
    '内容チェック用フォームのオープン
    If strOKFlg = "NG" Then
        'If Me![frm_text_select] = 2 Then
        '    strSubFormName = "F_1_取込内容確認_SUB_PACKING"
        'Else
            strSubFormName = "F_1_取込内容確認_SUB_INVOICE"
        'End If
        strFormName = "F_1_取込内容確認_MAIN"
        GoTo Err_btn_saijikko_Click
    End If
        
    '取り込むことにしちゃったら、再実行ボタンはもう要らない。
    'Me![btn_saijikko].Visible = False
        
    
    '保存テーブルに書き込み
    '明細
    With cmd
        .CommandText = "usp_shanghai_import_invoice_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Execute
    End With
    If IsNull(cmd.Parameters(3)) Or cmd.Parameters(3) = 0 Or IsNull(cmd.Parameters(4)) Or cmd.Parameters(4) = 0 Then
        MsgBox ("INVOICE、PACKINGの取込に失敗しました。" & vbLf & "EXCELデータご確認下さい。")
    End If
    
    '明細から合計情報を抽出してヘッダ情報作成
    intCartonQty = 0
    intSyukkaHoho = 0
    With cmd
        .CommandText = "usp_shanghai_invoice_gokei_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = strInvoiceNoMain
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(4)) Then
        intCartonQty = cmd.Parameters(4)
    End If
    If Not IsNull(cmd.Parameters(2)) Then
        datBLDate = cmd.Parameters(2)
    End If
    If Not IsNull(cmd.Parameters(3)) Then
        intSyukkaHoho = cmd.Parameters(3)
    End If
    If Not IsNull(cmd.Parameters(8)) And cmd.Parameters(8) <> 0 Then
        If intOkikaeFlg = 0 Then
            intOkikaeFlg = 2
        End If
    End If
        
    'ヘッダ部作成
    'If Me![frm_text_select] <> 1 And (intOkikaeFlg <> 2) Then
    If (intOkikaeFlg <> 2) Then
            
        strSql = "INSERT INTO tbl_t_shanghai_invoice_header ( pk_invoice_no_main, REF_NO, bl_date, vessel, syukka_hoho, carton_qty, "
        strSql = strSql & "return_pallet, otsunaka_no, shiiresaki_code, update_ymd, update_login ) VALUES ( "
        strSql = strSql & "'" & strInvoiceNoMain & "', null, '" & datBLDate & "', null, " & intSyukkaHoho & ", " & intCartonQty & ", "
        strSql = strSql & "0, 0, '" & strShiireCode & "', '" & Now() & "', '" & Me![login_code] & "' ) "
    
        'Debug.Print strSql
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
        
    End If
    
    
    '進捗状況更新
    Call schedule_update(5, Date, strInvoiceNoMain, 1)

    
    '取込み時のエラーファイル削除（空白セルのエラーが出るため）
    On Error Resume Next
    'DoCmd.DeleteObject acTable, "Sheet1$_インポート エラー"
    'DoCmd.DeleteObject acTable, "Sheet1$_インポート エラー1"
    'DoCmd.DeleteObject acTable, "Sheet1$_インポート エラー2"
    'DoCmd.DeleteObject acTable, "Sheet1$_インポート エラー3"
    On Error GoTo 0
    
    MsgBox ("処理終了しました。")
                        
    
Exit_btn_saijikko_Click:
    
    Set cmd = Nothing
    Exit Sub
    
Err_btn_saijikko_Click:
    
    Set cmd = Nothing
    
    strFormName = "F_1_取込内容確認_MAIN"
    Call next_form_open
    
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_1_上海輸入_INVOICE取込")
    intMenuNo = 101
    intMenuNoSeq = 0

End Sub

Private Sub btn_shanghai_chotatsu_Click()

    'If Me![btn_shanghai_chotatsu].Caption = "上海" Then
    '    Me![btn_shanghai_chotatsu].Caption = "貿易"
    '    intShanghaiChotatsu = 3
    'Else
    '    Me![btn_shanghai_chotatsu].Caption = "上海"
    '    intShanghaiChotatsu = 1
    'End If
    
    'intMenuNo = 101
    'Me![INVOICE_NO] = Null
    
    'strFormName = "F_1_上海輸入_INVOICE取込"
    'Call next_form_open

End Sub

Private Sub Form_Open(Cancel As Integer)
    
    DoCmd.Maximize
    
    intMenuNo = 101
    intMenuNoSeq = 0
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_koumoku_contents_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "FIL"
        .Parameters(2) = intMenuNo * 10 + intShanghaiChotatsu
        .Execute
    End With
    
    Me![PASS] = cmd.Parameters(3)
    Me![text_name] = cmd.Parameters(4)
    
    
    With cmd
        .CommandText = "usp_koumoku_master_pickup"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "COS"
    End With
    
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
    
        rs1.MoveFirst
        
        On Error Resume Next
        Do Until rs1.EOF
        
            On Error Resume Next
            Me("customer_kbn_text" & rs1![pk_seq_no].Value).Caption = rs1![contents2].Value
            
            rs1.MoveNext
            
        Loop
        On Error GoTo 0
        
    End If
    rs1.Close
    
    Set rs1 = Nothing
    Set cmd = Nothing
    'Me![btn_saijikko].Visible = False
    
End Sub

Private Sub frm_text_select_Click()

    If Me![frm_text_select] = 1 Then
        Me![INVOICE_NO].Visible = True
        Me![lbl_INVOICE_NO].Visible = True
    Else
        Me![INVOICE_NO].Visible = False
        Me![lbl_INVOICE_NO].Visible = False
        Me![INVOICE_NO] = Null
    End If
    
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
    
    Dim strRefNo As String
    Dim intCurrencyType As Integer
    
    strInvoiceNo = ""
    If Not IsNull(Me![INVOICE_NO]) Then
        strInvoiceNo = Me![INVOICE_NO]
    Else
        strInvoiceNo = "NULL"
    End If
    strOpenMode = "OC"
    
    strRefNo = ""
    strTextName = ""
    intCurrencyType = 0
    If Not IsNull(Me![REF_NO]) Then
        strRefNo = Me![REF_NO]
    End If
    If Not IsNull(Me![text_name]) Then
        strTextName = Me![text_name]
    End If
    If Not IsNull(Me![currency_type]) Then
        intCurrencyType = Me![currency_type]
    End If
    
    
    Call form_open_close("F_1_上海輸入_INVOICE取込", strFormName, strOpenMode, strInvoiceNo)
    
    If strFormName = "F_1_取込内容確認_MAIN" Then
        Forms(strFormName)![REF_NO] = strRefNo
        Forms(strFormName)![text_name] = strTextName
        Forms(strFormName)![currency_type] = intCurrencyType
    End If
        
    
End Sub

Private Sub shanghai_chotatsu_AfterUpdate()

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
    
    intMenuNo = 101
    Me![INVOICE_NO] = Null
    
    strFormName = "F_1_上海輸入_INVOICE取込"
    Call next_form_open

End Sub
