Attribute VB_Name = "Form_F_1_SAP_AMOUNT_MAIN"
Attribute VB_Base = "0{FF88E3C6-1631-4F10-8521-FDFCD0762AD9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_create_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    '仕入れ明細作成

    If IsNull(Me![taisyo_yyyy]) Or IsNull(Me![taisyo_mm]) Or Len(Me![taisyo_yyyy]) <= 0 Or Len(Me![taisyo_mm]) <= 0 Then
        MsgBox ("対象年月を入力してください。")
        Exit Sub
    End If
    
    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE NOを入力してください。")
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    strShiireCode = "%%"
    Call shiiresaki_get
    
    intShinkiFlg = 0
    '登録済みがあるか
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![shanghai_chotatsu]
        .Parameters(2) = Me![taisyo_yyyymm]
        .Parameters(3) = strInvoiceNo
        .Parameters(4) = strShiireCode
        .Execute
    End With
    If Not IsNull(cmd.Parameters(5)) And cmd.Parameters(5) <> 0 Then
        ReturnValue = MsgBox("登録済みを置き換えますか？", vbYesNo)
        If ReturnValue = vbNo Then
            GoTo Exit_btn_create_Click
        Else
            intShinkiFlg = 1
        End If
    End If
    
    '既存データの削除（抽出の条件と同じ条件のデータを削除する）
    If intShinkiFlg = 0 Then
        strSql = "DELETE tbl_t_kaigai_shiire_jisseki_ruikei FROM tbl_t_kaigai_shiire_jisseki_ruikei "
        strSql = strSql & "WHERE pk_yyyymm = " & Me![taisyo_yyyymm] & " "
        Select Case Me![shanghai_chotatsu]
            Case 0, 1
                '上海担当
                'If Len(strShiireCode) > 0 Then
                '    strSql = strSql & "AND ( shiiresaki_code = '" & strShiireCode & "') "
                'End If
                strSql = strSql & "AND ( boueki_flg = 0 ) "
            Case 2
                '調達担当
                strSql = strSql & "AND ( boueki_flg = 2 ) "
            Case 3
                '貿易
                'If Len(strShiireCode) > 0 Then
                '    strSql = strSql & "AND ( shiiresaki_code = '" & strShiireCode & "') "
                'End If
                strSql = strSql & "AND ( boueki_flg = 1 ) "
            Case Else
                MsgBox ("担当が選択されていません。" & vbLf & "ログイン右横のボタンをクリックして担当を選択してください。")
                Set cmd = Nothing
                Exit Sub
        End Select
        If Not IsNull(Me![INVOICE_NO_MAIN]) Then
            strSql = strSql & "AND invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
        End If
        'If Not IsNull(Me![shiiresaki_code]) Then
        '    strSql = strSql & "AND shiiresaki_code = '" & Me![shiiresaki_code] & "' "
        'End If
        
        'Debug.Print strSql
        
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
    End If
    
    '追加
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Execute
    End With
    
    Set cmd = Nothing
    
    Call chk_display_kbn_Click
    
    MsgBox ("集計完了しました、内容をご確認ください。")
    
    Exit Sub
    
Exit_btn_create_Click:
    
    Set cmd = Nothing
    
End Sub

Private Sub btn_delete_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![INVOICE_NO_MAIN]) Or Len(Me![INVOICE_NO_MAIN]) <= 0 Then
        MsgBox ("削除はINVOICE単位です。元INV_NOを入力してください。")
        Exit Sub
    End If
    
    ReturnValue = MsgBox("削除しますか？", vbYesNo)
    If ReturnValue = vbNo Then
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_delete"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![INVOICE_NO_MAIN]
        .Execute
    End With
    
    If IsNull(cmd.Parameters(2)) Or cmd.Parameters(2) = 0 Then
        MsgBox ("対象データはありませんでした。")
    Else
        If IsNull(cmd.Parameters(3)) Or cmd.Parameters(3) = 0 Then
            MsgBox ("削除完了しました。")
        Else
            MsgBox ("削除出来ませんでした。" & vbLf & "IT室に問い合わせてください。")
        End If
    End If
    
    Set cmd = Nothing

End Sub

Private Sub btn_excel_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![taisyo_yyyy]) Or IsNull(Me![taisyo_mm]) Or Len(Me![taisyo_yyyy]) <= 0 Or Len(Me![taisyo_mm]) <= 0 Then
        MsgBox ("対象年月を入力してください。")
        Exit Sub
    End If
    
    If IsNull(Me![chk_display_kbn]) Or Me![chk_display_kbn] = 0 Then
        Me![chk_display_kbn] = 1
    End If
    intMenuNoSeq = chk_display_kbn

    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 205
    Else
        intMenuNo = 107
    End If
    strInputCode = Me![taisyo_yyyymm]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_hacchu_check_Click()

    '購買発注明細確認へ
    strFormName = "F_1_R3購買発注_MAIN"
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 202
    Else
        intMenuNo = 105
    End If
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_invoice_amount_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![taisyo_yyyy]) Or IsNull(Me![taisyo_mm]) Or Len(Me![taisyo_yyyy]) <= 0 Or Len(Me![taisyo_mm]) <= 0 Then
        MsgBox ("対象年月を入力してください。")
        Exit Sub
    End If

    '海上保険依頼用明細抽出

    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        '進捗状況更新
        Call schedule_update(15, Date, Me![INVOICE_NO_MAIN], Me![shanghai_chotatsu])
    End If
    
    Select Case Me![shanghai_chotatsu]
        Case 3
            intMenuNo = 107
            intMenuNoSeq = 4
        Case 2
            intMenuNo = 205
            intMenuNoSeq = 3
        Case Else
            intMenuNo = 107
            intMenuNoSeq = 3
    End Select
    strInputCode = Me![taisyo_yyyymm]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_invoice_meisai_Click()

    'INVOICE明細照会フォームへ
    strFormName = "F_1_上海輸入_INVOICE_MAIN"
    intMenuNo = 102
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_menu_Click()

    'メインメニューへ戻る
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_pickup_Click()
    
    '仕入れ明細抽出

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![taisyo_yyyy]) Or IsNull(Me![taisyo_mm]) Or Len(Me![taisyo_yyyy]) <= 0 Or Len(Me![taisyo_mm]) <= 0 Then
        MsgBox ("対象年月を入力してください。")
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        strInvoiceNo = Me![INVOICE_NO_MAIN]
    Else
        strInvoiceNo = "%%"
    End If
    
    strShiireCode = "%%"
    Call shiiresaki_get
    
    '登録済みがあるか
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![shanghai_chotatsu]
        .Parameters(2) = Me![taisyo_yyyymm]
        .Parameters(3) = strInvoiceNo
        .Parameters(4) = strShiireCode
        .Execute
    End With
    If Not IsNull(cmd.Parameters(5)) And cmd.Parameters(5) <> 0 Then
        intShinkiFlg = 2
    Else
        intShinkiFlg = 1
    End If
    
    Me![total_amount] = 0
    Me![total_amount_usd] = 0
    Me![total_amount_rmb] = 0
    '抽出
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_temp_create"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 300
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Parameters(3) = intShinkiFlg
        .Parameters(4) = Me![taisyo_yyyymm]
        .Parameters(5) = strInvoiceNo
        .Parameters(6) = strShiireCode
        .Execute
    End With
    
    If IsNull(cmd.Parameters(7)) Or cmd.Parameters(7) = 0 Then
        MsgBox ("対象データはありませんでした。")
    Else
        If intShinkiFlg = 1 Then
            MsgBox ("新規作成です。" & vbLf & "登録する場合、作成ボタンをクリックしてください。")
        Else
            MsgBox ("内容を確認してください。")
        End If
    End If
    
    If Not IsNull(cmd.Parameters(8)) Then
        Me![total_amount] = cmd.Parameters(8)
    End If
    If Not IsNull(cmd.Parameters(9)) And cmd.Parameters(9) <> 0 Then
        Me![total_amount_usd] = cmd.Parameters(9)
    End If
    If Not IsNull(cmd.Parameters(10)) And cmd.Parameters(10) <> 0 Then
        Me![total_amount_rmb] = cmd.Parameters(10)
    End If
    'オーダー件数表示
    If Not IsNull(cmd.Parameters(7)) Then
        Me![order_cnt] = cmd.Parameters(7)
    End If
    
    Set cmd = Nothing
    
    Call chk_display_kbn_Click
    
End Sub

Private Sub btn_print_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![taisyo_yyyy]) Or IsNull(Me![taisyo_mm]) Or Len(Me![taisyo_yyyy]) <= 0 Or Len(Me![taisyo_mm]) <= 0 Then
        MsgBox ("対象年月を入力してください。")
        Exit Sub
    End If

    '印刷

    Dim intOthersCount As Integer
    Dim intShiiresakiKbn As Integer
    
    If IsNull(Me![chk_shiiresaki_kbn]) Or Me![chk_shiiresaki_kbn] = 0 Then
        Me![chk_shiiresaki_kbn] = 1
    End If
    
    Set cmd.ActiveConnection = conn
    
    '品目マスタ外の品目があるか？
    intOthersCount = 0
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_master_gai"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Execute
    End With
        
    If Not IsNull(cmd.Parameters(3)) Then
        intOthersCount = cmd.Parameters(3)
    End If
    
    Set cmd = Nothing
    
    'マスタ外品目が合ったら｢その他明細｣の印刷
    If intOthersCount <> 0 Then
        DoCmd.OpenReport "R_仕入明細書_上海輸入_その他", acViewPreview, , "pk_login_code = " & Me![login_code]
    End If

    '仕入れ明細の印刷
    If Me![shanghai_chotatsu] <> 2 Then
        DoCmd.OpenReport "R_仕入明細書_上海輸入", acViewPreview, , "pk_login_code = " & Me![login_code]
        DoCmd.OpenReport "R_仕入明細書_上海輸入_INVOICE別", acViewPreview, , "pk_login_code = " & Me![login_code]
    Else
        DoCmd.OpenReport "R_仕入明細書_調達輸入", acViewPreview, , "pk_login_code = " & Me![login_code]
    End If
    
End Sub

Private Sub btn_recreate_Click()

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("元INV NOを入力して下さい。")
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    strSql = "DELETE tbl_t_kaigai_shiire_jisseki_ruikei FROM tbl_t_kaigai_shiire_jisseki_ruikei "
    strSql = strSql & "WHERE pk_yyyymm = " & Me![taisyo_yyyymm] & " "
    strSql = strSql & "AND invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
        
    'Debug.Print strSql
    
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    Me![total_amount] = 0
    Me![total_amount_usd] = 0
    Me![total_amount_rmb] = 0
    '抽出
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Parameters(3) = 1
        .Parameters(4) = Me![taisyo_yyyymm]
        .Parameters(5) = Me![INVOICE_NO_MAIN]
        .Parameters(6) = "%%"
        .Execute
    End With
    
    If IsNull(cmd.Parameters(7)) Or cmd.Parameters(7) = 0 Then
        MsgBox ("対象データはありませんでした。")
    Else
        MsgBox ("内容を確認して、作成ボタンをクリックしてください。")
    End If
    
    If Not IsNull(cmd.Parameters(8)) Then
        Me![total_amount] = cmd.Parameters(8)
    End If
    If Not IsNull(cmd.Parameters(9)) And cmd.Parameters(9) <> 0 Then
        Me![total_amount_usd] = cmd.Parameters(9)
    End If
    If Not IsNull(cmd.Parameters(10)) And cmd.Parameters(10) <> 0 Then
        Me![total_amount_rmb] = cmd.Parameters(10)
    End If
    
    Set cmd = Nothing
    
    Call chk_display_kbn_Click
    

End Sub

Private Sub btn_sagaku_keisan_Click()

    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_sagaku_keisan"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    Set cmd = Nothing
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_1_SAP_AMOUNT_MAIN")
    If intShanghaiChotatsu = 2 Then
        intMenuNo = 205
    Else
        intMenuNo = 107
    End If
    intMenuNoSeq = 0

End Sub

Private Sub btn_shanghai_chotatsu_Click()

    'If Me![btn_shanghai_chotatsu].Caption = "上海" Then
    '    Me![btn_shanghai_chotatsu].Caption = "調達"
    '    intShanghaiChotatsu = 2
    '    intMenuNo = 205
    'Else
    '    If Me![btn_shanghai_chotatsu].Caption = "調達" Then
    '        Me![btn_shanghai_chotatsu].Caption = "貿易"
    '        intShanghaiChotatsu = 3
    '        intMenuNo = 107
    '    Else
    '        Me![btn_shanghai_chotatsu].Caption = "上海"
    '        intShanghaiChotatsu = 1
    '        intMenuNo = 107
    '    End If
    'End If
    
    'Me![INVOICE_NO_MAIN] = Null
    
    'strFormName = "F_1_SAP_AMOUNT_MAIN"
    'Call next_form_open

End Sub

Private Sub btn_zaiko_tenso_Click()

    '在庫転送処理へ
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

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    '表示の切替　INVOICE⇒購買発注、購買発注⇒INVOICE、購買予定⇒INVOICE
    If IsNull(Me![taisyo_yyyy]) Or IsNull(Me![taisyo_mm]) Or Len(Me![taisyo_yyyy]) <= 0 Or Len(Me![taisyo_mm]) <= 0 Then
        Exit Sub
    End If

    If Not IsNull(Me![taisyo_yyyy]) And Not IsNull(Me![taisyo_mm]) Then
        intTaisyoYyyy = Me![taisyo_yyyy]
        intTaisyoMm = Me![taisyo_mm]
        Me![taisyo_yyyymm] = Me![taisyo_yyyy] * 100 + Me![taisyo_mm]
    End If
    
    Select Case Me![shanghai_chotatsu]
        Case 3
            '貿易
            strSql = "SELECT dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main FROM dbo.tbl_t_shanghai_invoice_header "
            strSql = strSql & "INNER JOIN dbo.tbl_t_shanghai_invoice_meisai ON dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main = "
            strSql = strSql & "dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main "
            strSql = strSql & "WHERE bl_date >= '" & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm], 1) & "' "
            strSql = strSql & "AND dbo.tbl_t_shanghai_invoice_header.bl_date < '" & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm] + 1, 1) - 1 & "' "
            strSql = strSql & "AND dbo.tbl_t_shanghai_invoice_meisai.boueki_flg = 1 "
            strSql = strSql & "GROUP BY dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main ORDER BY dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main DESC "
        Case 2
            '調達
            strSql = "SELECT pk_invoice_no_main FROM dbo.tbl_t_chotatsu_invoice_header WHERE bl_date >= '"
            strSql = strSql & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm], 1) & "' AND bl_date < '"
            strSql = strSql & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm] + 1, 1) - 1 & "' "
            strSql = strSql & "GROUP BY pk_invoice_no_main ORDER BY pk_invoice_no_main DESC "
        Case Else
            '上海担当
            strSql = "SELECT dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main FROM dbo.tbl_t_shanghai_invoice_header "
            strSql = strSql & "INNER JOIN dbo.tbl_t_shanghai_invoice_meisai ON dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main = "
            strSql = strSql & "dbo.tbl_t_shanghai_invoice_meisai.pk_invoice_no_main "
            strSql = strSql & "WHERE dbo.tbl_t_shanghai_invoice_header.bl_date >= '" & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm], 1) & "' "
            strSql = strSql & "AND dbo.tbl_t_shanghai_invoice_header.bl_date < '" & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm] + 1, 1) - 1 & "' "
            strSql = strSql & "AND dbo.tbl_t_shanghai_invoice_meisai.boueki_flg = 0 "
            strSql = strSql & "GROUP BY dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main ORDER BY dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main DESC "
    End Select
    Me![INVOICE_NO_MAIN].RowSource = strSql
    Me![INVOICE_NO_MAIN].Requery

    If Me![chk_display_kbn] = 2 Then
        Me![F_1_SAP_AMOUNT_SUB].SourceObject = "F_1_SAP_AMOUNT_SUB2"
        Me![btn_create].Enabled = False
    Else
        Me![F_1_SAP_AMOUNT_SUB].SourceObject = "F_1_SAP_AMOUNT_SUB1"
        Me![btn_create].Enabled = True
    End If

    Me![F_1_SAP_AMOUNT_SUB].Requery
    Me.Requery
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    If intShanghaiChotatsu = 2 Then
        intMenuNo = 205
    Else
        intMenuNo = 107
    End If
    intMenuNoSeq = 0
    
    If Len(intTaisyoYyyy) <= 0 Or intTaisyoYyyy = 0 Then
        intTaisyoYyyy = Year(Date)
    End If
    If Len(intTaisyoMm) <= 0 Or intTaisyoMm = 0 Then
        intTaisyoMm = Month(Date)
    End If
    Me![taisyo_yyyy] = intTaisyoYyyy
    Me![taisyo_mm] = intTaisyoMm
    Me![taisyo_yyyymm] = intTaisyoYyyy * 100 + intTaisyoMm
    
    
    
    Set cmd.ActiveConnection = conn
    
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
    
    
    'Call chk_display_kbn_Click
    Me![F_1_SAP_AMOUNT_SUB].Requery
    Me.Requery
    
End Sub

Private Sub invoice_no_main_AfterUpdate()
    
    Call chk_display_kbn_Click

End Sub

Private Sub shanghai_chotatsu_AfterUpdate()

    Select Case Me![shanghai_chotatsu]
        Case 1
            intShanghaiChotatsu = 1
            intMenuNo = 107
        Case 2
            intShanghaiChotatsu = 2
            intMenuNo = 205
        Case 3
            intShanghaiChotatsu = 3
            intMenuNo = 107
        Case Else
            intShanghaiChotatsu = 1
            intMenuNo = 107
    End Select
    
    Me![INVOICE_NO_MAIN] = Null
    
    strFormName = "F_1_SAP_AMOUNT_MAIN"
    Call next_form_open

End Sub

Private Sub taisyo_mm_AfterUpdate()

    Call chk_display_kbn_Click

End Sub

Private Sub taisyo_yyyy_AfterUpdate()

    Call chk_display_kbn_Click

End Sub

Private Sub shiiresaki_get()

    'If Me![btn_shanghai_chotatsu].Caption = "調達" Or Me![shanghai_chotatsu] = 2 Then
    If Me![shanghai_chotatsu] = 2 Then
        If Not IsNull(Me![shiiresaki_code]) Then
            strShiireCode = Me![shiiresaki_code]
        End If
    Else
    
        If IsNull(Me![chk_shiiresaki_kbn]) Then
            intCustomerKbn = 0
        Else
            intCustomerKbn = Me![chk_shiiresaki_kbn]
        End If
    
        If intCustomerKbn <> 0 Then
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
                strShiireCode = "%" & cmd.Parameters(14) & "%"
            End If
        End If
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
    
    strInvoiceNo = ""
    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        strInvoiceNo = Me![INVOICE_NO_MAIN]
    Else
        strInvoiceNo = "NULL"
    End If
    If strFormName = "F_1_SAP_AMOUNT_MAIN" Then
        strOpenMode = "O"
    Else
    
        'If strFormName = "F_9_メール送信_MAIN" Then
        '    Set cmd.ActiveConnection = conn
            
        '    With cmd
        '        .CommandText = "usp_mail_soshin_invoice_no_temp_create"
        '        .CommandType = adCmdStoredProc
        '        .Parameters.Refresh
        '        .Parameters(1) = Me![login_code]
        '        .Parameters(2) = intShanghaiChotatsu
        '        .Parameters(3) = Me![taisyo_yyyy] * 100 + Me![taisyo_mm]
        '        .Execute
        '    End With
            
        '    Set cmd = Nothing
        'End If
            
        strOpenMode = "OC"
    End If
    
    Call form_open_close("F_1_SAP_AMOUNT_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

