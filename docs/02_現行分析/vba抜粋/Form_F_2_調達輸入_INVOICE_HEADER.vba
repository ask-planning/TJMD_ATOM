Attribute VB_Name = "Form_F_2_調達輸入_INVOICE_HEADER"
Attribute VB_Base = "0{8D198739-E3AF-4878-8F1D-F545E8422D6F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

    Dim intHeaderFlg As Integer


Private Sub bnt_menu_Click()

    'メニューに戻る
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    strOpenMode = "OC"
    Call next_form_open
    
End Sub


Private Sub btn_excel_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![pk_invoice_no_main]) Then
        Exit Sub
    End If

    intMenuNo = 203
    intMenuNoSeq = 1
    strOpenMode = "OC"
    strInputCode = ""
    strInvoiceNo = Me![pk_invoice_no_main]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open
    

End Sub

Private Sub btn_hacchu_kakunin_Click()

    'R3購買発注確認フォームへ
    strFormName = "F_1_R3購買発注_MAIN"
    intMenuNo = 202
    intMenuNoSeq = 0
    strOpenMode = "OC"
    Call next_form_open
    
End Sub

Private Sub btn_haiso_tehai_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If

    '配送手配書の印刷
    ReturnValue = MsgBox("配送タイプ、配送予定日の入力は済んでいますか。", 4)
    If ReturnValue = vbNo Then
        Exit Sub
    End If
    
    'Set cmd.ActiveConnection = conn
    
    'With cmd
    '    .CommandText = "usp_chotatsu_syukka_denpyo_temp_create"
    '    .CommandType = adCmdStoredProc
    '    .Parameters.Refresh
    '    .Parameters(1) = Me![login_code]
    '    .Parameters(2) = Me![pk_invoice_no_main]
    '    .Parameters(3) = "NULL"
    '    .Execute
    'End With
    
    'Set cmd = Nothing

    '進捗状況更新
    'Call schedule_update(7, Date, Me![pk_invoice_no_main], 2)
    
    If Me![syukka_hoho] = 0 Then
        If IsNull(Me![hasso_saki_2]) Or Me![hasso_saki_2] = 0 Or Len(Me![hasso_saki_2]) <= 0 Then
            DoCmd.OpenReport "R_調達輸入_配送手配_船便", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' AND [haiso_saki2] is null "
        Else
            DoCmd.OpenReport "R_調達輸入_配送手配_船便2", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' AND [haiso_saki2] is not null "
        End If
    Else
        DoCmd.OpenReport "R_調達輸入_配送手配_AIR", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' "
    End If
    
End Sub

Private Sub btn_invoice_change_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Dim strAfterInvNo As String

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    ReturnValue = InputBox("新しいINVOICE_NOを入力してください。", , Me![pk_invoice_no_main])
    If Len(ReturnValue) <= 0 Then
        MsgBox ("INVOICE_NOがありません。")
        Exit Sub
    End If
    
    strAfterInvNo = ReturnValue
    
    Set cmd.ActiveConnection = conn
    
    '関連テーブルのINVOICE_NOをすべて変更
    With cmd
        .CommandText = "usp_chotatsu_invoice_no_change"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_invoice_no_main]
        .Parameters(3) = strAfterInvNo
        .Execute
    End With
    
    
    Me![pk_invoice_no_main] = strAfterInvNo
    Me![pk_invoice_no_main].Requery
    
    Call pk_invoice_no_main_AfterUpdate
    
    MsgBox ("INVOICE_NOの変更完了しました。")
    
    
End Sub

Private Sub btn_kaijyo_hoken_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If

    '海上保険用の経費入力フォームへ
    
    Set cmd.ActiveConnection = conn
    
    '未作成の経費データ作成（製品金額）
    With cmd
        .CommandText = "usp_chotatsu_hoken_goods_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    Set cmd = Nothing
    
    strShiireCode = Me![shiiresaki_code]
    strFormName = "F_3_輸入諸経費入力_MAIN"
    strOpenMode = "OC"
    Call next_form_open
    
End Sub

Private Sub btn_nyuka_meisai_Click()

    '入荷明細作成と印刷
    
    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    Dim intSyoriKbn As Integer

    Dim intKeyHaisoType As Integer
    Dim strEditDenpyoNo As String
    
    Set cmd.ActiveConnection = conn
    
    intSyoriKbn = 0
    ReturnValue = InputBox("作成／編集：1　印刷：2", , 1)
    If Len(ReturnValue) <= 0 Then
        GoTo Exit_btn_nyuka_meisai_Click
    End If
    
    intSyoriKbn = ReturnValue
    
    '作成済みがあるか、あったら作り直すか
    If intSyoriKbn = 1 Then
        strSql = "SELECT COUNT(hinmoku_code) as meisai_count FROM dbo.tbl_t_chotatsu_nyuka_meisai WHERE invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
        End With
        Set rs1 = cmd.Execute(intRcnt)
        If intRcnt <> 0 Then
            rs1.MoveFirst
            If rs1![meisai_count].Value <> 0 Then
                ReturnValue = InputBox("登録済みの入荷明細があります。" & vbLf & vbLf & "(作り直し：1　編集：2)", , 2)
                If Len(ReturnValue) <= 0 Then
                    rs1.Close
                    GoTo Exit_btn_nyuka_meisai_Click
                Else
                    If ReturnValue = 1 Then
                        strSql = "DELETE dbo.tbl_t_chotatsu_nyuka_meisai FROM dbo.tbl_t_chotatsu_nyuka_meisai WHERE invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
                        With cmd
                            .CommandText = strSql
                            .CommandType = adCmdText
                            .Execute
                        End With
                    End If
                End If
            End If
        End If
        rs1.Close
    End If
    
    '作業用テンポラリ作成
    With cmd
        .CommandText = "usp_chotatsu_nyuka_meisai_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If IsNull(cmd.Parameters(3)) Or cmd.Parameters(3) = 0 Then
        MsgBox ("対象データはありませんでした")
        GoTo Exit_btn_nyuka_meisai_Click
    End If
    
    '編集画面へ
    If intSyoriKbn = 1 Then
        
        Dim strInvoiceNo As String
        
        strInvoiceNo = Me![pk_invoice_no_main]
        
        strFormName = "F_2_調達輸入_入荷明細編集_MAIN"
        intMenuNo = 206
        intMenuNoSeq = 0
        strOpenMode = "O"
        
        Call next_form_open
        
        Forms![F_2_調達輸入_入荷明細編集_MAIN]![pk_invoice_no_main] = strInvoiceNo
        Forms![F_2_調達輸入_入荷明細編集_MAIN]![F_2_調達輸入_入荷明細編集_SUB].Requery
        
        GoTo Exit_btn_nyuka_meisai_Click
    End If
        
    
    '印刷用テンポラリの削除
    strSql = "DELETE tbl_t_chotatsu_syukka_denpyo_temp FROM tbl_t_chotatsu_syukka_denpyo_temp WHERE pk_login_code = '" & Me![login_code] & "' "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With

    '出荷伝票番号の編集
    strSql = "SELECT pk_syukka_denpyo_no, haiso_type FROM tbl_t_chotatsu_syukka_denpyo WHERE [pk_invoice_no_main] = '" & Me![pk_invoice_no_main] & "' "
    strSql = strSql & "ORDER BY haiso_type, pk_syukka_denpyo_no "
    
'Debug.Print strSql
    strEditDenpyoNo = ""
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
        rs1.MoveFirst
        intKeyHaisoType = rs1![haiso_type].Value
        Do Until rs1.EOF
        
            '配送タイプが変わったら、まとめた出荷伝票番号をテンポラリに書き込む
            If intKeyHaisoType <> rs1![haiso_type].Value Then
                strSql = "INSERT tbl_t_chotatsu_syukka_denpyo_temp ( pk_login_code, pk_invoice_no_main, pk_syukka_denpyo_no, haiso_type ) "
                strSql = strSql & "VALUES ( '" & Me![login_code] & "', '" & Me![pk_invoice_no_main] & "', '" & strEditDenpyoNo & "', " & intKeyHaisoType & " ) "
                With cmd
                    .CommandText = strSql
                    .CommandType = adCmdText
                    .Execute
                End With
                
                strEditDenpyoNo = ""
                intKeyHaisoType = rs1![haiso_type].Value
            End If
            If Len(strEditDenpyoNo) > 0 Then
                strEditDenpyoNo = strEditDenpyoNo & "/"
            End If
            strEditDenpyoNo = strEditDenpyoNo & rs1![pk_syukka_denpyo_no].Value
            rs1.MoveNext
        Loop
        '最後のレコードの書き込み
        strSql = "INSERT tbl_t_chotatsu_syukka_denpyo_temp ( pk_login_code, pk_invoice_no_main, pk_syukka_denpyo_no, haiso_type ) "
        strSql = strSql & "VALUES ( '" & Me![login_code] & "', '" & Me![pk_invoice_no_main] & "', '" & strEditDenpyoNo & "', " & intKeyHaisoType & " ) "
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
    End If
    rs1.Close
    
    'With cmd
    '    .CommandText = "usp_chotatsu_syukka_denpyo_temp_create"
    '    .CommandType = adCmdStoredProc
    '    .Parameters.Refresh
    '    .Parameters(1) = Me![login_code]
    '    .Parameters(2) = Me![pk_invoice_no_main]
    '    .Parameters(3) = strEditDenpyoNo
    '    .Execute
    'End With
    
    '進捗状況更新
    intScheduleSyoriNo = 0
    With cmd
        .CommandText = "usp_koumoku_contents_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "HKB"
        .Parameters(2) = Me![haiso_type]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(4)) Or Len(cmd.Parameters(4)) > 0 Then
        intScheduleSyoriNo = cmd.Parameters(4)
    End If
    If Not IsNull(cmd.Parameters(5)) Or Len(cmd.Parameters(5)) > 0 Then
        intScheduleSyoriNo = cmd.Parameters(5)
    End If
        
    Select Case intScheduleSyoriNo
        Case 1, 12
            intScheduleSyoriNo = 10
        Case 3, 13
            intScheduleSyoriNo = 12
        Case 4, 5, 8
            intScheduleSyoriNo = 11
        Case 11, 15
            intScheduleSyoriNo = 14
    End Select
    
    Set cmd = Nothing

    'Call schedule_update(intScheduleSyoriNo, Date, Me![pk_invoice_no_main], 2)

    DoCmd.OpenReport "R_調達輸入_入荷明細", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' "
    
    Exit Sub
    
Exit_btn_nyuka_meisai_Click:
    
    Set cmd = Nothing
    
End Sub

Private Sub btn_nyuka_schedule_Click()
    
    'Set cmd.ActiveConnection = conn
    
    'With cmd
    '    .CommandText = "usp_chotatsu_syukka_denpyo_temp_create"
    '    .CommandType = adCmdStoredProc
    '    .Parameters.Refresh
    '    .Parameters(1) = Me![login_code]
    '    .Parameters(2) = Me![pk_invoice_no_main]
    '    .Parameters(3) = "NULL"
    '    .Execute
    'End With
    
    'Set cmd = Nothing

    DoCmd.OpenReport "R_調達輸入_日新入荷", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' "

End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_2_調達輸入_INVOICE_HEADER")
    intMenuNo = 203
    intMenuNoSeq = 0

End Sub

Private Sub btn_syohin_name_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    strInvoiceNo = Me![pk_invoice_no_main]
    strShiireCode = Me![shiiresaki_code]
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_chotatsu_syohingun_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shiiresaki_code]
        .Execute
    End With
    
    Set cmd = Nothing
    
    strFormName = "F_M_調達輸入_商品名称入力_MAIN"
    strOpenMode = "OC"
    Call next_form_open
    

End Sub

Private Sub btn_syukka_denpyo_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        Exit Sub
    End If

    Dim strInvoiceNo As String
    
    strInvoiceNo = Me![pk_invoice_no_main]
    
    strFormName = "F_2_調達輸入_出荷伝票登録_MAIN"
    intMenuNo = 206
    intMenuNoSeq = 0
    strOpenMode = "O"
    
    Call next_form_open
    
End Sub

Private Sub btn_torikomi_Click()

    '購買発注取込処理へ
    strFormName = "F_1_R3購買発注取込"
    intMenuNo = 201
    intMenuNoSeq = 0
    strOpenMode = "OC"
    Call next_form_open
    
End Sub

Private Sub btn_touroku_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    DoCmd.GoToControl "pk_invoice_no_main"
    
    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    If IsNull(Me![haiso_type]) Then
        MsgBox ("配送タイプを入力してください。")
        Exit Sub
    End If
    
    If IsNull(Me![insurance_seikyu_ymd]) And Not IsNull(Me![BL_DATE]) Then
        Me![insurance_seikyu_ymd] = Me![BL_DATE]
    End If
    
    Dim strHaisoPlant As String
    
    
    Set cmd.ActiveConnection = conn
    
    'INVOICEの存在チェック
    intHeaderFlg = 0
    With cmd
        .CommandText = "usp_chotatsu_invoice_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(2)) Then
        intHeaderFlg = cmd.Parameters(2)
    End If
    
    'ヘッダ部作成
    intFieldCount = 26
    
    ReDim strFieldName(intFieldCount) As String
    ReDim intFieldType(intFieldCount) As Integer
    
    strFieldName(0) = "pk_invoice_no_main":      intFieldType(0) = 0
    strFieldName(1) = "shiiresaki_code":         intFieldType(1) = 0
    strFieldName(2) = "bl_date":                 intFieldType(2) = 0
    strFieldName(3) = "vessel":                  intFieldType(3) = 0
    strFieldName(4) = "syukka_hoho":             intFieldType(4) = 1
    strFieldName(5) = "ETD":                     intFieldType(5) = 0
    strFieldName(6) = "ETA":                     intFieldType(6) = 0
    strFieldName(7) = "carton_qty":              intFieldType(7) = 1
    strFieldName(8) = "haiso_yotei_ymd":         intFieldType(8) = 0
    strFieldName(9) = "haiso_type":              intFieldType(9) = 1
    strFieldName(10) = "otsunaka_no":            intFieldType(10) = 1
    'strFieldName(12) = "syukka_denpyo_no":       intFieldType(12) = 0
    strFieldName(11) = "invoice_maisu":          intFieldType(11) = 1
    strFieldName(12) = "packing_maisu":          intFieldType(12) = 1
    strFieldName(13) = "bl_maisu":               intFieldType(13) = 1
    strFieldName(14) = "arrival_notice_maisu":   intFieldType(14) = 1
    strFieldName(15) = "insurance_maisu":        intFieldType(15) = 1
    strFieldName(16) = "chk_arrival_notice":     intFieldType(16) = 1
    strFieldName(17) = "chk_insurance":          intFieldType(17) = 1
    strFieldName(18) = "chk_bl":                 intFieldType(18) = 1
    strFieldName(19) = "bl_option":              intFieldType(19) = 1
    strFieldName(20) = "insurance_seikyu_ymd":   intFieldType(20) = 0
    strFieldName(21) = "container_type":         intFieldType(21) = 1
    strFieldName(22) = "container_no":           intFieldType(22) = 0
    strFieldName(23) = "seal_no":                intFieldType(23) = 0
    strFieldName(24) = "update_ymd":             intFieldType(24) = 3
    strFieldName(25) = "update_login":           intFieldType(25) = 3
    
    If intHeaderFlg = 0 Then
    
        strSql = "INSERT INTO tbl_t_chotatsu_invoice_header ("
        For n = 0 To intFieldCount - 1
            If n <> 0 Then
                strSql = strSql & ","
            End If
            strSql = strSql & strFieldName(n)
        Next
        strSql = strSql & ") VALUES ("
        For n = 0 To intFieldCount - 1
            If n <> 0 Then
                strSql = strSql & ", "
            End If
            Select Case intFieldType(n)
                Case 0
                    If IsNull(Me(strFieldName(n))) Or Len(Me(strFieldName(n))) <= 0 Then
                        strSql = strSql & "null"
                    Else
                        strSql = strSql & "'" & Me(strFieldName(n)) & "'"
                    End If
                Case 1
                    If IsNull(Me(strFieldName(n))) Or Len(Me(strFieldName(n))) <= 0 Then
                        strSql = strSql & "null"
                    Else
                        strSql = strSql & Me(strFieldName(n))
                    End If
                Case 3
                    If strFieldName(n) = "update_ymd" Then
                        strSql = strSql & "'" & Now() & "'"
                    End If
                    If strFieldName(n) = "update_login" Then
                        strSql = strSql & "'" & Me![login_code] & "'"
                    End If
            End Select
        Next
        strSql = strSql & " )"
                
    Else
    
        strSql = "UPDATE tbl_t_chotatsu_invoice_header SET "
        For n = 1 To intFieldCount - 1
            If n <> 1 Then
                strSql = strSql & ", "
            End If
            Select Case intFieldType(n)
                Case 0
                    If IsNull(Me(strFieldName(n))) Or Len(Me(strFieldName(n))) <= 0 Then
                        strSql = strSql & strFieldName(n) & " = null"
                    Else
                        strSql = strSql & strFieldName(n) & " = '" & Me(strFieldName(n)) & "'"
                    End If
                Case 1
                    If IsNull(Me(strFieldName(n))) Or Len(Me(strFieldName(n))) <= 0 Then
                        strSql = strSql & strFieldName(n) & " = null"
                    Else
                        strSql = strSql & strFieldName(n) & " = " & Me(strFieldName(n))
                    End If
                Case 3
                    If strFieldName(n) = "update_ymd" Then
                        strSql = strSql & strFieldName(n) & " = '" & Now() & "'"
                    End If
                    If strFieldName(n) = "update_login" Then
                        strSql = strSql & strFieldName(n) & " = '" & Me![login_code] & "'"
                    End If
            End Select
        Next
        strSql = strSql & " WHERE pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
        
    End If
    
    'Debug.Print strSql
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    '社内参照番号別配送先の更新
    With cmd
        .CommandText = "usp_chotatsu_invoice_header_sub_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    
    '購買発注明細の配送先の更新　⇒　ヘッダサブUPDATEストアドの中で更新　2014/09/18
    'strHaisoPlant = ""
    'With cmd
    '    .CommandText = "usp_koumoku_contents_get"
    '    .CommandType = adCmdStoredProc
    '    .Parameters.Refresh
    '    .Parameters(1) = "TNT"
    '    .Parameters(2) = Me![haiso_type]
    '    .Execute
    'End With
    'If Not IsNull(cmd.Parameters(4)) Then
    '    strHaisoPlant = cmd.Parameters(4)
    'End If
    
    'If Len(strHaisoPlant) > 0 Then
    '    strSql = "UPDATE tbl_t_R3_koubai_hacchu_meisai SET nounyu_plant = '" & strHaisoPlant & "', "
    '    strSql = strSql & "update_ymd = '" & Now() & "', update_login = '" & Me![login_code] & "' "
    '    strSql = strSql & "WHERE invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
    '    With cmd
    '        .CommandText = strSql
    '        .CommandType = adCmdText
    '        .Execute
    '    End With
    'End If
    
    
    
    MsgBox ("登録完了しました。")
    
    Set cmd = Nothing

    '進捗状況更新
    'Call schedule_update(3, Date, Me![pk_invoice_no_main], 2)
    
    Call pk_invoice_no_main_AfterUpdate

End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
    intMenuNo = 203
    intMenuNoSeq = 0
    
    Me![btn_invoice_change].Enabled = False

End Sub

Private Sub pk_invoice_no_main_AfterUpdate()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    Call chotatsu_invoice_header_display(Me![pk_invoice_no_main], "F_2_調達輸入_INVOICE_HEADER")
    
    Me![btn_invoice_change].Enabled = True
    
Exit_pk_invoice_no_main_AfterUpdate:
    
    Set rs1 = Nothing
    Set cmd = Nothing

End Sub


Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    strInvoiceNo = ""
    If Not IsNull(Me![pk_invoice_no_main]) Then
        strInvoiceNo = Me![pk_invoice_no_main]
    Else
        strInvoiceNo = "NULL"
    End If
    
    Call form_open_close("F_2_調達輸入_INVOICE_HEADER", strFormName, strOpenMode, strInvoiceNo)
    
End Sub
