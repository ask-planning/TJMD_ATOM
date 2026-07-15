Attribute VB_Name = "Form_F_1_R3購買発注_MAIN"
Attribute VB_Base = "0{72C37F1B-DEC0-4ABB-85C6-1EC6B5DFFA3B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_chichibu_print_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If

    Dim strSoshinInvoiceNo As String
    Dim datLastUpdateYmd As Date
    Dim intMeisaiCount As Integer
    
    Dim strNounyuPlant As String
    Dim strInvoiceNoSH As String
    Dim strInvoiceNoYS As String

    Set cmd.ActiveConnection = conn
    
    
    strKoubaiHacchuNo = ""
    strSyanaiSansyoNo = ""
    strNounyuPlant = ""
    If Not IsNull(Me![koubai_hacchu_no]) Then
        strKoubaiHacchuNo = "%" & Me![koubai_hacchu_no] & "%"
    Else
        strKoubaiHacchuNo = "%%"
    End If
    If Not IsNull(Me![syanai_sansyo_no]) Then
        strSyanaiSansyoNo = "%" & Me![syanai_sansyo_no] & "%"
    Else
        strSyanaiSansyoNo = "%%"
    End If
    If Not IsNull(Me![nounyu_plant]) Then
        strNounyuPlant = "%" & Me![nounyu_plant] & "%"
    Else
        strNounyuPlant = "%%"
    End If
    
    
    intMeisaiCount = 0
    strSoshinInvoiceNo = ""
    With cmd
        .CommandText = "usp_shanghai_nyuka_meisai_kensu"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![shanghai_chotatsu]
        .Parameters(2) = Me![INVOICE_NO_MAIN]
        .Parameters(3) = strSyanaiSansyoNo
        .Parameters(4) = strNounyuPlant
        .Execute
    End With
    If Not IsNull(cmd.Parameters(5)) Then
        datLastUpdateYmd = cmd.Parameters(5)
    End If
    If Not IsNull(cmd.Parameters(6)) Then
        intMeisaiCount = cmd.Parameters(6)
    End If
    
    If intMeisaiCount <> 0 Then
        If intModeAtom = 0 Or intModeAtom = 9 Then
            ReturnValue = InputBox("現在作成済みのINVOICEの最終更新日は" & datLastUpdateYmd & "です。" & vbLf & "2：内容を確認する　3：印刷する", , 3)
        Else
            ReturnValue = InputBox("現在作成済みのINVOICEの最終更新日は" & datLastUpdateYmd & "です。" & vbLf & "1:作り直す　2：内容を確認する　3：印刷する", , 3)
        End If
    Else
        If intModeAtom <> 0 And intModeAtom <> 9 Then
            ReturnValue = MsgBox("作成しますか？", vbYesNo)
            If ReturnValue = vbYes Then
                ReturnValue = "1"
            Else
                ReturnValue = "0"
            End If
        Else
            MsgBox ("作成済みの入荷明細がありません。")
            ReturnValue = "0"
            GoTo Exit_btn_chichibu_print
        End If
    End If
    If intModeAtom <> 0 And intModeAtom <> 9 And (Len(ReturnValue) <= 0 Or (ReturnValue <= 1 And ReturnValue > 3)) Then
        MsgBox ("１～３を選択してください。")
        Exit Sub
    Else
        If (intModeAtom = 0 Or intModeAtom = 9) And (Len(ReturnValue) <= 0 Or (ReturnValue <= 2 And ReturnValue > 3)) Then
            MsgBox ("２、または、３を選択してください。")
        End If
    End If
    
    DoCmd.Hourglass True
    If ReturnValue = 1 Then
    
        '作り直し
        With cmd
            .CommandText = "usp_shanghai_nyuka_meisai_create"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = Me![shanghai_chotatsu]
            .Parameters(3) = Me![INVOICE_NO_MAIN]
            .Parameters(4) = strSyanaiSansyoNo
            .Parameters(5) = strNounyuPlant
            .Execute
        End With
        
        ReturnValue = 3
        
    End If
    
    'テンポラリに出力
    With cmd
        .CommandText = "usp_shanghai_nyuka_meisai_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Parameters(3) = Me![INVOICE_NO_MAIN]
        .Parameters(4) = strSyanaiSansyoNo
        .Parameters(5) = strNounyuPlant
        .Execute
    End With
    
Exit_btn_chichibu_print:
    
    DoCmd.Hourglass False
    
    Set cmd = Nothing
    
    If ReturnValue = 3 Then

        '進捗状況更新
        Call schedule_update(11, Date, Me![INVOICE_NO_MAIN], 1)

        DoCmd.OpenReport "R_上海タジマ入荷明細_秩父用", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' "
    
    End If
    
    Me![chk_display_kbn] = 3
    Call chk_display_kbn_Click
    
End Sub

Private Sub btn_hacchu_pickup_Click()

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    '社内参照番号抽出
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 202
    Else
        intMenuNo = 105
    End If
    intMenuNoSeq = 6
    strInputCode = Me![INVOICE_NO_MAIN]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_invoice_meisai_Click()

    'INVOICE明細照会フォームへ
    strFormName = "F_1_上海輸入_INVOICE_MAIN"
    intShanghaiChotatsu = 1
    intMenuNo = 102
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_kaigaikoubainyuko_Click()
    Dim URL As String
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
        
    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("元INV_NOを入力して下さい。")
        Exit Sub
    End If
    
    URL = "http://172.20.136.216:7700/dataspider/trigger/KaigaiKoubaiNyuko?InvoiceNo=" & Me![INVOICE_NO_MAIN]
    
    If MsgBox("入庫処理を実行しますか？", vbYesNo) = vbYes Then
        
    Else
        Exit Sub
    End If
    
    With IE
        .Navigate (URL)
        .Visible = False
    End With

    Set IE = Nothing
    MsgBox ("実行しました。R3入庫処理されるまで2,3分程かかります。")

End Sub

Private Sub btn_koushin_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If Me![chk_display_kbn] <> 2 And Me![chk_display_kbn] <> 3 Then
        Exit Sub
    End If
    
    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE NOを入力してください。")
        Exit Sub
    End If
    
    '購買発注データ更新（行削除もあるので、登録済みを削除して追加する）
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_R3_koubai_hacchu_temp_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    If IsNull(cmd.Parameters(2)) Or cmd.Parameters(2) = 0 Then
        MsgBox ("データが抽出されていません")
        GoTo Exit_btn_koushin_Click
    Else
        ReturnValue = MsgBox("更新しますか？", vbYesNo)
        If ReturnValue = vbNo Then
            GoTo Exit_btn_koushin_Click
        End If
    End If
    
    '登録済みの購買発注データの削除
    If Me![chk_display_kbn] = 2 Then
        strSql = "DELETE tbl_t_R3_koubai_hacchu_meisai FROM tbl_t_R3_koubai_hacchu_meisai WHERE invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
        If Not IsNull(Me![koubai_hacchu_no]) Then
            strSql = strSql & "AND pk_koubai_hacchu_no = '" & Me![koubai_hacchu_no] & "' "
        End If
        If Not IsNull(Me![syanai_sansyo_no]) Then
            strSql = strSql & "AND syanai_sansyo_no = '" & Me![syanai_sansyo_no] & "' "
        End If
        If Not IsNull(Me![nounyu_plant]) Then
            strSql = strSql & "AND nounyu_plant = '" & Me![nounyu_plant] & "' "
        End If
        If Me![shanghai_chotatsu] = 3 Then
            strSql = strSql & "AND koubai_group = '002' "
        Else
            strSql = strSql & "AND koubai_group <> '002' "
        End If
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
    
        '購買発注データ更新
        With cmd
            .CommandText = "usp_R3_koubai_hacchu_meisai_update"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Execute
        End With
    Else
        
        strSql = "DELETE tbl_t_shanghai_nyuka_meisai FROM tbl_t_shanghai_nyuka_meisai WHERE pk_invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
        If Not IsNull(Me![koubai_hacchu_no]) Then
            strSql = strSql & "AND koubai_hacchu_no = '" & Me![koubai_hacchu_no] & "' "
        End If
        If Not IsNull(Me![syanai_sansyo_no]) Then
            strSql = strSql & "AND syanai_sansyo_no = '" & Me![syanai_sansyo_no] & "' "
        End If
        If Not IsNull(Me![nounyu_plant]) Then
            strSql = strSql & "AND nounyu_plant = '" & Me![nounyu_plant] & "' "
        End If
        If Me![shanghai_chotatsu] = 3 Then
            strSql = strSql & "AND boueki_flg = 1 "
        Else
            strSql = strSql & "AND boueki_flg <> 1 "
        End If
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
    
        '入荷明細データ更新
        With cmd
            .CommandText = "usp_shanghai_nyuka_meisai_update"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Execute
        End With
    End If
    
    Set cmd = Nothing
    
    Call syori_log_create(intMenuNo, intMenuNoSeq, Me![INVOICE_NO_MAIN], "btn_koushin")
    
    MsgBox ("更新完了しました。")
    
    Exit Sub
    
Exit_btn_koushin_Click:
    
    Set cmd = Nothing

End Sub


Private Sub btn_matome_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE NOを入力してください。")
        Exit Sub
    End If

    If Me![chk_display_kbn] = 3 Then
        ReturnValue = MsgBox("入荷明細を作成します。" & vbLf & "既存デーは置換えになります。実行（はい）／キャンセル（いいえ）", 4)
    Else
        ReturnValue = MsgBox("購買発注のまとめデータを作成します。" & vbLf & "既存デーは置換えになります。実行（はい）／キャンセル（いいえ）", 4)
    End If
    If ReturnValue = vbNo Then
        Exit Sub
    End If
    
    '購買発注データ更新（行削除もあるので、登録済みを削除して追加する）
    
    Dim strKeyKoubaiNo As String
    Dim strKeyTensoNo As String
    Dim strKeySyukkaNo As String
    Dim strEditSyukkaNo As String
    Dim strKeyPlant As String
    
    Dim strNounyuPlant As String
    
    
    Set cmd.ActiveConnection = conn
    
    
    If Me![chk_display_kbn] = 3 Then
    
        strKoubaiHacchuNo = ""
        strSyanaiSansyoNo = ""
        strNounyuPlant = ""
        If Not IsNull(Me![koubai_hacchu_no]) Then
            strKoubaiHacchuNo = "%" & Me![koubai_hacchu_no] & "%"
        Else
            strKoubaiHacchuNo = "%%"
        End If
        If Not IsNull(Me![syanai_sansyo_no]) Then
            strSyanaiSansyoNo = "%" & Me![syanai_sansyo_no] & "%"
        Else
            strSyanaiSansyoNo = "%%"
        End If
        If Not IsNull(Me![nounyu_plant]) Then
            strNounyuPlant = "%" & Me![nounyu_plant] & "%"
        Else
            strNounyuPlant = "%%"
        End If
        
        '登録済みの入荷明細データの削除
        strSql = "DELETE tbl_t_shanghai_nyuka_meisai FROM tbl_t_shanghai_nyuka_meisai WHERE pk_invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
        If Not IsNull(Me![koubai_hacchu_no]) Then
            strSql = strSql & "AND koubai_hacchu_no LIKE '%" & Me![koubai_hacchu_no] & "%' "
        End If
        If Not IsNull(Me![syanai_sansyo_no]) Then
            strSql = strSql & "AND syanai_sansyo_no LIKE '%" & Me![syanai_sansyo_no] & "%' "
        End If
        If Not IsNull(Me![nounyu_plant]) Then
            strSql = strSql & "AND nounyu_plant LIKE '%" & Me![nounyu_plant] & "%' "
        End If
        If Me![shanghai_chotatsu] = 3 Then
            strSql = strSql & "AND boueki_flg = 1 "
        Else
            strSql = strSql & "AND boueki_flg <> 1 "
        End If
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
        
        '作り直し
        With cmd
            .CommandText = "usp_shanghai_nyuka_meisai_recreate"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = Me![shanghai_chotatsu]
            .Parameters(3) = Me![INVOICE_NO_MAIN]
            .Parameters(4) = strSyanaiSansyoNo
            .Parameters(5) = strNounyuPlant
            .Execute
        End With
        
        'テンポラリに出力
        With cmd
            .CommandText = "usp_shanghai_nyuka_meisai_temp_create"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = Me![shanghai_chotatsu]
            .Parameters(3) = Me![INVOICE_NO_MAIN]
            .Parameters(4) = strSyanaiSansyoNo
            .Parameters(5) = strNounyuPlant
            .Execute
        End With
    
    Else
    
        '登録済みの購買発注合計データの削除
        strSql = "DELETE tbl_t_R3_koubai_hacchu_amount FROM tbl_t_R3_koubai_hacchu_amount WHERE invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
        If Not IsNull(Me![koubai_hacchu_no]) Then
            strSql = strSql & "AND koubai_hacchu_no = '" & Me![koubai_hacchu_no] & "' "
        End If
        If Not IsNull(Me![syanai_sansyo_no]) Then
            strSql = strSql & "AND pk_syanai_sansyo_no = '" & Me![syanai_sansyo_no] & "' "
        End If
        If Not IsNull(Me![nounyu_plant]) Then
            strSql = strSql & "AND nounyu_plant = '" & Me![nounyu_plant] & "' "
        End If
        If Me![shanghai_chotatsu] = 3 Then
            strSql = strSql & "AND koubai_group = '002' "
        Else
            strSql = strSql & "AND koubai_group <> '002' "
        End If
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
        
        '購買発注データ更新
        With cmd
            .CommandText = "usp_R3_koubai_hacchu_gokei_update"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Execute
        End With

        '合計データの出荷伝票NOの編集(購買発注、在庫転送の番号で出荷伝票をまとめる）
        strSql = "SELECT pk_koubai_hacchu_no, syanai_sansyo_no, zaiko_tenso_no, syukka_denpyo_no, invoice_no_main "
        strSql = strSql & "FROM tbl_t_R3_koubai_hacchu_meisai_temp WHERE pk_login_code = '" & Me![login_code] & "' "
        strSql = strSql & "AND syukka_denpyo_no is not null "
        strSql = strSql & "GROUP BY pk_koubai_hacchu_no, syanai_sansyo_no, zaiko_tenso_no, syukka_denpyo_no, invoice_no_main "
        strSql = strSql & "ORDER BY invoice_no_main, pk_koubai_hacchu_no, zaiko_tenso_no, syukka_denpyo_no "
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
        End With
    
        'Debug.Print strSql
        Set rs1 = cmd.Execute(intRcnt)
        If intRcnt <> 0 Then
            rs1.MoveFirst
            strKeyKoubaiNo = rs1![pk_koubai_hacchu_no].Value
            strKeyTensoNo = rs1![zaiko_tenso_no].Value
            strKeySyukkaNo = rs1![syukka_denpyo_no].Value
            strEditSyukkaNo = rs1![syukka_denpyo_no].Value
            Do Until rs1.EOF
                If strKeyKoubaiNo <> rs1![pk_koubai_hacchu_no].Value Then
                    strSql = "UPDATE tbl_t_R3_koubai_hacchu_amount SET syukka_denpyo_no = '" & strEditSyukkaNo & "' WHERE koubai_hacchu_no = '"
                    strSql = strSql & strKeyKoubaiNo & "' AND zaiko_tenso_no = '" & strKeyTensoNo & "' "
                    With cmd
                        .CommandText = strSql
                        .CommandType = adCmdText
                        .Execute
                    End With
                    
                    strKeyKoubaiNo = rs1![pk_koubai_hacchu_no].Value
                    strKeyTensoNo = rs1![zaiko_tenso_no].Value
                    strKeySyukkaNo = rs1![syukka_denpyo_no].Value
                    strEditSyukkaNo = rs1![syukka_denpyo_no].Value
                End If
                If strKeyTensoNo <> rs1![zaiko_tenso_no].Value Then
                    strSql = "UPDATE tbl_t_R3_koubai_hacchu_amount SET syukka_denpyo_no = '" & strEditSyukkaNo & "' WHERE koubai_hacchu_no = '"
                    strSql = strSql & strKeyKoubaiNo & "' AND zaiko_tenso_no = '" & strKeyTensoNo & "' "
                    With cmd
                        .CommandText = strSql
                        .CommandType = adCmdText
                        .Execute
                    End With
                    
                    strKeyTensoNo = rs1![zaiko_tenso_no].Value
                    strKeySyukkaNo = rs1![syukka_denpyo_no].Value
                    strEditSyukkaNo = rs1![syukka_denpyo_no].Value
                End If
                
                For n = 1 To Len(rs1![syukka_denpyo_no].Value)
                    If Mid(strKeySyukkaNo, n, 1) <> Mid(rs1![syukka_denpyo_no].Value, n, 1) Then
                        strEditSyukkaNo = strEditSyukkaNo & "/" & Right(rs1![syukka_denpyo_no].Value, Len(rs1![syukka_denpyo_no].Value) - n + 1)
                        n = Len(rs1![syukka_denpyo_no].Value)
                    End If
                Next
                
                rs1.MoveNext
            Loop
            
            strSql = "UPDATE tbl_t_R3_koubai_hacchu_amount SET syukka_denpyo_no = '" & strEditSyukkaNo & "' WHERE koubai_hacchu_no = '"
            strSql = strSql & strKeyKoubaiNo & "' AND zaiko_tenso_no = '" & strKeyTensoNo & "' "
            With cmd
                .CommandText = strSql
                .CommandType = adCmdText
                .Execute
            End With
                
        End If
        rs1.Close
        
    End If
    
    Set cmd = Nothing
    
    Me![F_1_R3購買発注_SUB].Requery
    
    Call syori_log_create(intMenuNo, intMenuNoSeq, Me![INVOICE_NO_MAIN], "btn_matome")
    
    MsgBox ("作成完了しました。")
    
End Sub

Private Sub btn_menu_Click()

    'メインメニューへ戻る
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_pickup_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    Dim intMeisaiCount As Integer
    
    Dim strNounyuPlant As String
    Dim strInvoiceNoSH As String
    Dim strInvoiceNoYS As String
    
    
    strKoubaiHacchuNo = ""
    strSyanaiSansyoNo = ""
    strNounyuPlant = ""
    If Not IsNull(Me![koubai_hacchu_no]) Then
        strKoubaiHacchuNo = "%" & Me![koubai_hacchu_no] & "%"
    Else
        strKoubaiHacchuNo = "%%"
    End If
    If Not IsNull(Me![syanai_sansyo_no]) Then
        strSyanaiSansyoNo = "%" & Me![syanai_sansyo_no] & "%"
    Else
        strSyanaiSansyoNo = "%%"
    End If
    If Not IsNull(Me![nounyu_plant]) Then
        strNounyuPlant = "%" & Me![nounyu_plant] & "%"
    Else
        strNounyuPlant = "%%"
    End If
    
    If Me![chk_display_kbn] = 4 Then
        If Me![shanghai_chotatsu] <> 2 Then
            If IsNull(Me![invoice_no_sh]) And IsNull(Me![invoice_no_ys]) Then
                MsgBox ("INVOICE NOを入力してください。")
                Exit Sub
            Else
                If Len(Me![invoice_no_sh]) > 0 And Len(Me![invoice_no_sh]) <= 7 Then
                    MsgBox ("上海タジマのINVOICE NOの桁数が違います。")
                    Exit Sub
                Else
                    If Len(Me![invoice_no_sh]) > 0 Then
                        strInvoiceNoSH = Me![invoice_no_sh]
                    End If
                End If
                If Len(Me![invoice_no_ys]) > 0 And Len(Me![invoice_no_ys]) <= 7 Then
                    MsgBox ("庸助貿易のINVOICE NOの桁数が違います。")
                    Exit Sub
                Else
                    If Len(Me![invoice_no_ys]) > 0 Then
                        strInvoiceNoYS = Me![invoice_no_ys]
                    End If
                End If
                ReturnValue = MsgBox("ロジの検数用、入庫処理用データを作成します。" & vbLf & "INVOICEは、" & strInvoiceNoSH & "と" & strInvoiceNoYS & "でよろしいですか？" & vbLf & "実行（はい）／キャンセル（いいえ）", vbYesNo)
                If ReturnValue = vbNo Then
                    Exit Sub
                End If
            End If
        Else
            strInvoiceNoSH = Me![INVOICE_NO_MAIN]
            strInvoiceNoYS = Me![INVOICE_NO_MAIN]
        End If
        
    Else
        If IsNull(Me![INVOICE_NO_MAIN]) Then
            MsgBox ("INVOICE NOを入力してください。")
            Exit Sub
        End If
    End If

        
    Set cmd.ActiveConnection = conn

    
    intMeisaiCount = 0
    If Me![chk_display_kbn] = 4 Then
    
            'ロジ送信用データ作成
            With cmd
                .CommandText = "usp_R3_koubai_hacchu_logistics_temp_create"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Parameters(2) = strInvoiceNoSH
                .Parameters(3) = strInvoiceNoYS
                .Execute
            End With
            
            If Not IsNull(cmd.Parameters(4)) Then
                intMeisaiCount = cmd.Parameters(4)
            End If
            
            'Set cmd = Nothing
            
            'intMenuNoSeq = 3
            'strInputCode = strInvoiceNoSH & "/" & strInvoiceNoYS
            
            'strFormName = "F_9_メール送信_MAIN"
            'Call next_form_open
            
            'Exit Sub
            
    Else
        DoCmd.Hourglass True
        
        If Me![chk_display_kbn] = 3 Then
        
            '入荷明細の有無チェック、無かったら作る
            If Me![shanghai_chotatsu] <> 2 Then
                
                With cmd
                    .CommandText = "usp_shanghai_nyuka_meisai_kensu"
                    .CommandType = adCmdStoredProc
                    .Parameters.Refresh
                    .Parameters(1) = Me![shanghai_chotatsu]
                    .Parameters(2) = Me![INVOICE_NO_MAIN]
                    .Parameters(3) = strSyanaiSansyoNo
                    .Parameters(4) = strNounyuPlant
                    .Execute
                End With
                If Not IsNull(cmd.Parameters(6)) Then
                    intMeisaiCount = cmd.Parameters(6)
                End If
                
                If intMeisaiCount = 0 Then
                
                    ReturnValue = MsgBox("新規作成しますか？", vbYesNo)
                    If ReturnValue = vbYes Then
                    
                        '作り直し
                        With cmd
                            .CommandText = "usp_shanghai_nyuka_meisai_create"
                            .CommandType = adCmdStoredProc
                            .CommandTimeout = 300
                            .Parameters.Refresh
                            .Parameters(1) = Me![login_code]
                            .Parameters(2) = Me![shanghai_chotatsu]
                            .Parameters(3) = Me![INVOICE_NO_MAIN]
                            .Parameters(4) = strSyanaiSansyoNo
                            .Parameters(5) = strNounyuPlant
                            .Execute
                        End With
                        
                        'Debug.Print cmd.Parameters(6)
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
        
        '購買発注データ抽出
        With cmd
            .CommandText = "usp_R3_koubai_hacchu_temp_create"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = Me![shanghai_chotatsu]
            .Parameters(3) = Me![INVOICE_NO_MAIN]
            .Parameters(4) = strKoubaiHacchuNo
            .Parameters(5) = strSyanaiSansyoNo
            .Parameters(6) = strNounyuPlant
            .Execute
        End With
        If Me![chk_display_kbn] = 1 Or Me![chk_display_kbn] = 2 Then
            If Not IsNull(cmd.Parameters(7)) And cmd.Parameters(7) <> 0 Then
                intMeisaiCount = cmd.Parameters(7)
            End If
        End If
        
        
        '入荷明細データ抽出
        If Me![shanghai_chotatsu] <> 2 Then
                   
            'テンポラリに出力
            With cmd
                .CommandText = "usp_shanghai_nyuka_meisai_temp_create"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
                .Parameters(2) = Me![shanghai_chotatsu]
                .Parameters(3) = Me![INVOICE_NO_MAIN]
                .Parameters(4) = strSyanaiSansyoNo
                .Parameters(5) = strNounyuPlant
                .Execute
            End With
            If Me![chk_display_kbn] = 3 Then
                If Not IsNull(cmd.Parameters(6)) And cmd.Parameters(6) <> 0 Then
                    intMeisaiCount = cmd.Parameters(6)
                Else
                    intMeisaiCount = 0
                End If
            End If
            
        End If
        
    End If
    
    DoCmd.Hourglass False
    
    If intMeisaiCount = 0 Then
        MsgBox ("対象データはありませんでした。")
    End If
        
    Set cmd = Nothing
    
    Me![F_1_R3購買発注_SUB].Requery
    
End Sub

Private Sub btn_print_Click()

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If

    Dim intReportNo As Integer
    
    If Me![chk_display_kbn] = 4 Then
        ReturnValue = InputBox("1:ロジ_検数用明細　2:ロジ_入庫処理用明細", , "1")
        If Len(ReturnValue) <= 0 Then
            Exit Sub
        Else
            If Not IsNull(ReturnValue) And Len(ReturnValue) > 0 And IsNumeric(ReturnValue) _
                And (ReturnValue = "1" Or ReturnValue <= "2") Then
                intReportNo = ReturnValue
            Else
                Exit Sub
            End If
        End If
    Else
        intReportNo = 3
    End If
    
    Select Case intReportNo
        Case 1
            strReportName = "R_購買発注_ロジ送信_検数用"
        Case 2
            strReportName = "R_購買発注_ロジ送信_入庫処理用"
        Case 3
            strReportName = "R_R3購買発注確認"
        Case Else
            strReportName = "R_購買発注_ロジ送信_検数用"
    End Select

    DoCmd.OpenReport strReportName, acViewPreview, , "[pk_login_code] = '" & Me![login_code] & "' "
    
End Sub

Private Sub btn_sap_amount_Click()

    '金額チェックへ
    strFormName = "F_1_SAP_AMOUNT_MAIN"
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 205
    Else
        intMenuNo = 107
    End If
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_1_R3購買発注_MAIN")
    If intShanghaiChotatsu = 2 Then
        intMenuNo = 202
    Else
        intMenuNo = 105
    End If
    intMenuNoSeq = 0

End Sub

Private Sub btn_shanghai_chotatsu_Click()

    'If Me![btn_shanghai_chotatsu].Caption = "上海" Then
    '    Me![btn_shanghai_chotatsu].Caption = "調達"
    '    intShanghaiChotatsu = 2
    '    intMenuNo = 202
    'Else
    '    If Me![btn_shanghai_chotatsu].Caption = "調達" Then
    '        Me![btn_shanghai_chotatsu].Caption = "貿易"
    '        intShanghaiChotatsu = 3
    '        intMenuNo = 105
    '    Else
    '        Me![btn_shanghai_chotatsu].Caption = "上海"
    '        intShanghaiChotatsu = 1
    '        intMenuNo = 105
    '    End If
    'End If
    
    'Me![INVOICE_NO_MAIN] = Null
    
    'strFormName = "F_1_R3購買発注_MAIN"
    'Call next_form_open

End Sub

Private Sub btn_syukko_pickup_Click()

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    '出荷伝票番号抽出
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 202
    Else
        intMenuNo = 105
    End If
    intMenuNoSeq = 8
    strInputCode = Me![INVOICE_NO_MAIN]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_tehacchu_zaiko_pickup_Click()

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    '在庫転送番号抽出
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 202
    Else
        intMenuNo = 105
    End If
    intMenuNoSeq = 9
    strInputCode = Me![INVOICE_NO_MAIN]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_torikomi_Click()

    '購買発注取込処理へ
    strFormName = "F_1_R3購買発注取込"
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 201
    Else
        intMenuNo = 104
    End If
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_zaiko_pickup_Click()

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    '在庫転送番号抽出
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 202
    Else
        intMenuNo = 105
    End If
    intMenuNoSeq = 7
    strInputCode = Me![INVOICE_NO_MAIN]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_zaiko_tenso_Click()

    '在庫転送作成処理へ
    strFormName = "F_1_R3在庫転送_MAIN"
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 204
    Else
        intMenuNo = 106
    End If
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub chk_display_kbn_Click()

    intMenuNoSeq = Me![chk_display_kbn]

    Select Case Me![chk_display_kbn]
        Case 1
            'まとめ
            Me![F_1_R3購買発注_SUB].SourceObject = "F_1_R3購買発注_SUB1"
            Me![btn_matome].Caption = "まとめ作成"
            Me![btn_matome].Enabled = False
            Me![btn_print].Enabled = True
            Me![btn_koushin].Enabled = False
            Me![INVOICE_NO_MAIN].Enabled = True
            Me![invoice_no_sh] = Null
            Me![invoice_no_ys] = Null
            Me![lbl_invoice_no_ys].Visible = False
            Me![lbl_invoice_no_sh].Visible = False
            Me![invoice_no_sh].Visible = False
            Me![invoice_no_ys].Visible = False
            Me![koubai_hacchu_no].Enabled = True
        Case 2
            '明細
            Me![F_1_R3購買発注_SUB].SourceObject = "F_1_R3購買発注_SUB2"
            Me![btn_print].Enabled = False
            Me![btn_matome].Caption = "まとめ作成"
            If intModeAtom = 0 Or intModeAtom = 9 Then
                Me![btn_matome].Enabled = False
                Me![btn_koushin].Enabled = False
            Else
                Me![btn_matome].Enabled = True
                Me![btn_koushin].Enabled = True
            End If
            Me![INVOICE_NO_MAIN].Enabled = True
            Me![invoice_no_sh] = Null
            Me![invoice_no_ys] = Null
            Me![lbl_invoice_no_ys].Visible = False
            Me![lbl_invoice_no_sh].Visible = False
            Me![invoice_no_sh].Visible = False
            Me![invoice_no_ys].Visible = False
            Me![koubai_hacchu_no].Enabled = True
        Case 4
            'ロジ用
            Me![F_1_R3購買発注_SUB].SourceObject = "F_1_R3購買発注_SUB3"
            Me![btn_matome].Caption = "まとめ作成"
            Me![btn_matome].Enabled = False
            If Me![shanghai_chotatsu] = 2 Then
                Me![INVOICE_NO_MAIN].Enabled = True
                Me![lbl_invoice_no_ys].Visible = False
                Me![lbl_invoice_no_sh].Visible = False
                Me![invoice_no_sh].Visible = False
                Me![invoice_no_ys].Visible = False
            Else
                Me![INVOICE_NO_MAIN].Enabled = False
                Me![lbl_invoice_no_ys].Visible = True
                Me![lbl_invoice_no_sh].Visible = True
                Me![invoice_no_sh].Visible = True
                Me![invoice_no_ys].Visible = True
            End If
            Me![btn_print].Enabled = True
            Me![btn_koushin].Enabled = False
            Me![koubai_hacchu_no].Enabled = True
        Case 3
            '入荷明細
            Me![F_1_R3購買発注_SUB].SourceObject = "F_1_R3購買発注_SUB4"
            Me![btn_matome].Caption = "(再)作成"
            Me![btn_matome].Enabled = True
            Me![btn_print].Enabled = False
            If intModeAtom = 0 Or intModeAtom = 9 Then
                Me![btn_koushin].Enabled = False
            Else
                Me![btn_koushin].Enabled = True
            End If
            Me![INVOICE_NO_MAIN].Enabled = True
            Me![invoice_no_sh] = Null
            Me![invoice_no_ys] = Null
            Me![lbl_invoice_no_ys].Visible = False
            Me![lbl_invoice_no_sh].Visible = False
            Me![invoice_no_sh].Visible = False
            Me![invoice_no_ys].Visible = False
            Me![koubai_hacchu_no].Enabled = False
    End Select
    
    '候補リスト
    If Me![chk_display_kbn] <> 4 Then
        If Not IsNull(Me![INVOICE_NO_MAIN]) Then
            strSql = "SELECT pk_koubai_hacchu_no FROM tbl_t_R3_koubai_hacchu_meisai WHERE invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
            strSql = strSql & "GROUP BY pk_koubai_hacchu_no ORDER BY pk_koubai_hacchu_no DESC "
            Me![koubai_hacchu_no].RowSource = strSql
        
            If Not IsNull(Me![koubai_hacchu_no]) Then
                strSql = "SELECT syanai_sansyo_no FROM tbl_t_R3_koubai_hacchu_meisai WHERE invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
                strSql = strSql & "AND pk_koubai_hacchu_no = '" & Me![koubai_hacchu_no] & "' GROUP BY syanai_sansyo_no "
                strSql = strSql & "ORDER BY syanai_sansyo_no "
            Else
                If Me![shanghai_chotatsu] <> 2 Then
                    strSql = "SELECT syanai_sansyo_no FROM tbl_t_shanghai_invoice_meisai WHERE pk_invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
                    strSql = strSql & "AND LEFT(syanai_sansyo_no, 3) <> 'NSN' AND LEFT(syanai_sansyo_no, 3) <> 'HON' "
                    strSql = strSql & "GROUP BY syanai_sansyo_no ORDER BY syanai_sansyo_no "
                Else
                    strSql = "SELECT syanai_sansyo_no FROM tbl_t_R3_koubai_hacchu_meisai WHERE invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
                    strSql = strSql & "GROUP BY syanai_sansyo_no ORDER BY syanai_sansyo_no "
                End If
            End If
            Me![syanai_sansyo_no].RowSource = strSql
        Else
            If Not IsNull(Me![koubai_hacchu_no]) Then
                strSql = "SELECT syanai_sansyo_no FROM tbl_t_R3_koubai_hacchu_meisai WHERE pk_koubai_hacchu_no = '" & Me![koubai_hacchu_no] & "' "
                strSql = strSql & " GROUP BY syanai_sansyo_no ORDER BY syanai_sansyo_no "
            Else
                If Me![shanghai_chotatsu] <> 2 Then
                    strSql = "SELECT syanai_sansyo_no FROM tbl_t_shanghai_invoice_meisai GROUP BY syanai_sansyo_no ORDER BY syanai_sansyo_no "
                Else
                    strSql = "SELECT syanai_sansyo_no FROM tbl_t_R3_koubai_hacchu_meisai GROUP BY syanai_sansyo_no ORDER BY syanai_sansyo_no "
                End If
            End If
            Me![syanai_sansyo_no].RowSource = strSql
        End If
    End If

    Me![INVOICE_NO_MAIN].Requery
    Me![koubai_hacchu_no].Requery
    Me![syanai_sansyo_no].Requery
    Me![F_1_R3購買発注_SUB].Requery
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    DoCmd.Maximize
    
    Me![chk_display_kbn] = 2
    Me![F_1_R3購買発注_SUB].SourceObject = "F_1_R3購買発注_SUB2"
    Me![btn_print].Enabled = False
    Me![btn_matome].Caption = "まとめ作成"
    If intModeAtom = 0 Or intModeAtom = 9 Then
        Me![btn_matome].Enabled = False
        Me![btn_koushin].Enabled = False
    Else
        Me![btn_matome].Enabled = True
        Me![btn_koushin].Enabled = True
    End If
    Me![INVOICE_NO_MAIN].Enabled = True
    Me![invoice_no_sh] = Null
    Me![invoice_no_ys] = Null
    Me![lbl_invoice_no_ys].Visible = False
    Me![lbl_invoice_no_sh].Visible = False
    Me![invoice_no_sh].Visible = False
    Me![invoice_no_ys].Visible = False
    Me![koubai_hacchu_no].Enabled = True
    
    Select Case intShanghaiChotatsu
        Case 3
            'Me![INVOICE_NO_MAIN].RowSource = "View__invoice_no_main_boueki"
            strSql = "SELECT dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main, dbo.tbl_t_shanghai_invoice_header.BL_DATE "
            strSql = strSql & "FROM dbo.tbl_t_shanghai_invoice_meisai INNER JOIN dbo.tbl_t_shanghai_invoice_header ON "
            strSql = strSql & "dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main = dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main "
            strSql = strSql & "WHERE (dbo.tbl_t_shanghai_invoice_meisai.boueki_flg = 1) "
            strSql = strSql & "GROUP BY dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main, dbo.tbl_t_shanghai_invoice_header.BL_DATE "
            strSql = strSql & "ORDER BY dbo.tbl_t_shanghai_invoice_header.bl_date DESC, dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main DESC "
            Me![INVOICE_NO_MAIN].RowSource = strSql
            Me![btn_sap_amount].Enabled = True
            Me![btn_invoice_meisai].Enabled = True
            Me![btn_chichibu_print].Enabled = True
            Me![chk_soshin].Enabled = True
            intMenuNo = 105
            Me![btn_invoice_header].Caption = "INVヘッダ"
            Me![btn_invoice_header].ForeColor = "8388672"
        Case 2
            'Me![INVOICE_NO_MAIN].RowSource = "View__invoice_no_main_chotatsu"
            strSql = "SELECT pk_invoice_no_main, bl_date FROM dbo.tbl_t_chotatsu_invoice_header ORDER BY bl_date DESC "
            Me![INVOICE_NO_MAIN].RowSource = strSql
            Me![btn_sap_amount].Enabled = False
            Me![btn_invoice_meisai].Enabled = False
            Me![btn_chichibu_print].Enabled = False
            Me![btn_invoice_header].Caption = "配送手配"
            Me![btn_invoice_header].ForeColor = "10040115"
            Me![chk_soshin].Enabled = False
            intMenuNo = 202
        Case Else
            'Me![INVOICE_NO_MAIN].RowSource = "View__invoice_no_main"
            strSql = "SELECT dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main, dbo.tbl_t_shanghai_invoice_header.bl_date "
            strSql = strSql & "FROM dbo.tbl_t_shanghai_invoice_meisai INNER JOIN dbo.tbl_t_shanghai_invoice_header ON "
            strSql = strSql & "dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main = dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main "
            strSql = strSql & "WHERE (dbo.tbl_t_shanghai_invoice_meisai.boueki_flg = 0) OR (dbo.tbl_t_shanghai_invoice_meisai.boueki_flg IS NULL) "
            strSql = strSql & "GROUP BY dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main, dbo.tbl_t_shanghai_invoice_header.bl_date "
            strSql = strSql & "ORDER BY dbo.tbl_t_shanghai_invoice_header.bl_date DESC, dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main DESC "
            Me![INVOICE_NO_MAIN].RowSource = strSql
            Me![btn_sap_amount].Enabled = True
            Me![btn_invoice_meisai].Enabled = True
            Me![btn_chichibu_print].Enabled = True
            Me![chk_soshin].Enabled = True
            intMenuNo = 105
            Me![btn_invoice_header].Caption = "INVヘッダ"
            Me![btn_invoice_header].ForeColor = "8388672"
    End Select
    Me![INVOICE_NO_MAIN].Requery
    
    intMenuNoSeq = 0
    
End Sub

Private Sub invoice_no_main_AfterUpdate()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_R3_koubai_hacchu_temp_clear"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    Set cmd = Nothing
    
    Call chk_display_kbn_Click
    
End Sub

Private Sub plant_AfterUpdate()

    Call chk_display_kbn_Click

End Sub

Private Sub nounyu_plant_AfterUpdate()

    Call chk_display_kbn_Click

End Sub

Private Sub shanghai_chotatsu_AfterUpdate()

    Select Case Me![shanghai_chotatsu]
        Case 1
            intShanghaiChotatsu = 1
            intMenuNo = 105
        Case 2
            intShanghaiChotatsu = 2
            intMenuNo = 202
        Case 3
            intShanghaiChotatsu = 3
            intMenuNo = 105
        Case Else
            intShanghaiChotatsu = 1
            intMenuNo = 105
    End Select
    
    Me![INVOICE_NO_MAIN] = Null
    
    strFormName = "F_1_R3購買発注_MAIN"
    Call next_form_open

End Sub

Private Sub to_excel_Click()

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    If Me![chk_display_kbn] = 4 Then
        'ReturnValue = InputBox("1:購買発注まとめ　2:購買発注明細　3:ロジ検数用リスト　4：ロジ入庫処理用リスト", , "3")
        ReturnValue = InputBox("1:ロジ検数用リスト　2：ロジ入庫処理用リスト", , "1")
        If Len(ReturnValue) <= 0 Then
            MsgBox ("送信する明細が選択されていません。")
            Exit Sub
        Else
            If ReturnValue <> "1" And ReturnValue <> "2" Then
                MsgBox ("１または２を選択してください。")
                Exit Sub
            Else
                If ReturnValue = "2" Then
                    intMenuNoSeq = 5
                Else
                    intMenuNoSeq = 4
                End If
            End If
        End If
    Else
        intMenuNoSeq = Me![chk_display_kbn]
    End If

    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 202
    Else
        intMenuNo = 105
    End If
    strInputCode = Me![INVOICE_NO_MAIN]
    
    strFormName = "F_9_メール送信_MAIN"
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

Private Sub koubai_hacchu_no_AfterUpdate()

    Call chk_display_kbn_Click

End Sub

Private Sub syanai_sansyo_no_AfterUpdate()

    Call chk_display_kbn_Click

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
    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        strInvoiceNo = Me![INVOICE_NO_MAIN]
    Else
        strInvoiceNo = "NULL"
    End If
    If strFormName = "F_1_R3購買発注_MAIN" Then
        strOpenMode = "O"
    Else
        strOpenMode = "OC"
    End If
    
    Call form_open_close("F_1_R3購買発注_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub
