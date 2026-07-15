Attribute VB_Name = "Form_F_1_上海輸入_INVOICE_HEADER"
Attribute VB_Base = "0{808A5EDD-842F-4F68-A152-0DEF4435F95E}"
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
    Call next_form_open
    
End Sub

Private Sub btn_backNo_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Dim strEditRefNo As String
    Dim dblRefNo As Double
    
    'REF番号をひとつ戻す
    
    Set cmd.ActiveConnection = conn
    
    strEditRefNo = ""
    dblRefNo = 0
    With cmd
        .CommandText = "usp_shanghai_ref_no_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "TY05-"
        .Execute
    End With
    If Not IsNull(cmd.Parameters(2)) And Not IsNull(cmd.Parameters(3)) Then
        strEditRefNo = cmd.Parameters(2) & cmd.Parameters(3) - 1
        dblRefNo = cmd.Parameters(3) - 1
    End If
    
    Me![REF_NO] = strEditRefNo
    
    If dblRefNo <> 0 Then
        
        With cmd
            .CommandText = "usp_shanghai_ref_no_update"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = "TY05-"
            .Parameters(2) = dblRefNo
            .Parameters(3) = Me![pk_invoice_no_main]
            .Parameters(4) = strEditRefNo
            .Execute
        End With
        
    End If
    
    Set cmd = Nothing
    
End Sub


Private Sub btn_hacchu_kakunin_Click()

    'R3購買発注確認フォームへ
    strFormName = "F_1_R3購買発注_MAIN"

    intMenuNo = 105
    intMenuNoSeq = 0
    
    'Debug.Print intShanghaiChotatsu
    
    Call next_form_open
    
End Sub

Private Sub btn_haiso_tehai_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE NOを入力してください。")
        Exit Sub
    End If
    
    If Me![shanghai_chotatsu] <> 1 Then
        MsgBox ("処理権限がありません。（国内発注用）")
        Exit Sub
    End If

    '配送手配書の印刷
    ReturnValue = MsgBox("配送タイプ、コンテナ区分、コンテナNOの入力は済んでいますか。", 4)
    If ReturnValue = vbNo Then
        Exit Sub
    End If

    Dim datHaisoYotei As Date
    Dim intContainerKbn As Integer
    Dim strContainerKbn As String
    Dim intHaisoType As Integer
    Dim intHaisoTypeCount As Integer
    Dim dblCartonQty As Double
    Dim dblTotalQty As Double
    Dim intHaisosaki1 As Integer
    Dim intHaisosaki2 As Integer
    Dim intHaisosaki3 As Integer
    Dim intOtsunakaNo As Integer
    
    Dim strContainerNo As String
    Dim strSealNo As String
    Dim datBLDate As Date
    Dim strSenmei As String
    
    Dim intTblHaisoType() As Integer
    Dim strTblHaisoTypeText() As String
    Dim intTblHaisosaki() As Integer
    
    '明細行にコントロールがあるとその行が更新されないかも、なので、コントロールの移動
    DoCmd.GoToControl "pk_invoice_no_main"
    
    Set cmd.ActiveConnection = conn
    
    '配送タイプの登録数取得
    intHaisoTypeCount = 0
    With cmd
        .CommandText = "usp_koumoku_master_count"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "HKB"
        .Execute
    End With
    If Not IsNull(cmd.Parameters(2)) Then
        intHaisoTypeCount = cmd.Parameters(2)
    End If
    
    '配送タイプの配列再定義(０はありえないのでIF分は入れない）
    ReDim intTblHaisoType(intHaisoTypeCount) As Integer
    ReDim intTblHaisosaki(intHaisoTypeCount) As Integer
    ReDim strTblHaisoTypeText(intHaisoTypeCount) As String
    
    For n = 0 To intHaisoTypeCount
        intTblHaisoType(n) = 0
        intTblHaisosaki(n) = 0
        strTblHaisoTypeText(n) = ""
    Next
    
    strSql = "DELETE tbl_t_shanghai_haiso_tehai_print_temp FROM tbl_t_shanghai_haiso_tehai_print_temp WHERE pk_login_code = '" & Me![login_code] & "' "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    strSql = "SELECT * FROM View_shanghai_haiso_tehai_pickup WHERE View_shanghai_haiso_tehai_pickup.pk_invoice_no_main = '"
    strSql = strSql & Me![pk_invoice_no_main] & "' ORDER BY View_shanghai_haiso_tehai_pickup.haiso_yotei_ymd, "
    strSql = strSql & "View_shanghai_haiso_tehai_pickup.haiso_type, View_shanghai_haiso_tehai_pickup.container_type, "
    strSql = strSql & "View_shanghai_haiso_tehai_pickup.container_no, View_shanghai_haiso_tehai_pickup.pk_invoice_no "

    'Debug.Print strSql
    
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
    
        '配送パターンでグループ化して、印刷用テンポラリに吐き出し
        rs1.MoveFirst
        strInvoiceNo = rs1![pk_invoice_no].Value
        If Not IsNull(rs1![haiso_type].Value) Then
            intHaisoType = rs1![haiso_type].Value
        Else
            intHaisoType = 0
        End If
        If Not IsNull(rs1![container_type].Value) Then
            intContainerKbn = rs1![container_type].Value
            strContainerKbn = rs1![container_type_text].Value
        Else
            intContainerKbn = 0
            strContainerKbn = ""
        End If
        If Not IsNull(rs1![haiso_yotei_ymd].Value) Then
            datHaisoYotei = rs1![haiso_yotei_ymd].Value
        Else
            datHaisoYotei = Date
        End If
        intHaisosaki1 = 0
        intHaisosaki2 = 0
        intHaisosaki3 = 0
        dblCartonQty = 0
        dblTotalQty = 0
        strContainerNo = ""
        strSealNo = ""
        intOtsunakaNo = 0
        If Not IsNull(rs1![haiso_saki1].Value) Then
            intHaisosaki1 = rs1![haiso_saki1].Value
        End If
        If Not IsNull(rs1![haiso_saki2].Value) Then
            intHaisosaki2 = rs1![haiso_saki2].Value
        End If
        If Not IsNull(rs1![haiso_saki3].Value) Then
            intHaisosaki3 = rs1![haiso_saki3].Value
        End If
        If Not IsNull(rs1![container_no].Value) Then
            strContainerNo = rs1![container_no].Value
        End If
        If Not IsNull(rs1![seal_no].Value) Then
            strSealNo = rs1![seal_no].Value
        End If
        If Not IsNull(rs1![vessel].Value) And Len(rs1![vessel].Value) > 0 Then
            strSenmei = rs1![vessel].Value
        End If
        If Not IsNull(rs1![BL_DATE].Value) And Len(rs1![BL_DATE].Value) > 0 Then
            datBLDate = rs1![BL_DATE].Value
        End If
        If Not IsNull(rs1![otsunaka_no].Value) Then
            intOtsunakaNo = rs1![otsunaka_no].Value
        End If
        Do Until rs1.EOF
            If (Not IsNull(rs1![haiso_yotei_ymd].Value) And rs1![haiso_yotei_ymd].Value <> datHaisoYotei) _
                Or (Not IsNull(rs1![haiso_type].Value) And intHaisoType <> 0 And intHaisoType <> rs1![haiso_type].Value) _
                Or (Not IsNull(rs1![container_type].Value) And intContainerKbn <> 0 And intContainerKbn <> rs1![container_type].Value) _
                Or (Not IsNull(rs1![container_no].Value) And Len(strContainerNo) > 0 And strContainerNo <> rs1![container_no].Value) Then
                
                strSql = "INSERT INTO tbl_t_shanghai_haiso_tehai_print_temp (pk_login_code, ref_no, invoice_no_main, invoice_no, "
                strSql = strSql & "vessel, bl_date, haiso_type, haiso_saki1, haiso_saki2, haiso_saki3, container_type, "
                strSql = strSql & "container_no, seal_no, total_carton, total_quantity, haisyo_yotei_ymd, week_day, otsunaka_no, "
                strSql = strSql & "nisshin_plant, logistic_plant ) VALUES ('" & Me![login_code] & "', "
                If Len(Me![REF_NO]) > 0 And Not IsNull(Me![REF_NO]) Then
                    strSql = strSql & "'" & Me![REF_NO] & "','" & Me![pk_invoice_no_main] & "','" & strInvoiceNo & "'"
                Else
                    strSql = strSql & "null, '" & Me![pk_invoice_no_main] & "','" & strInvoiceNo & "'"
                End If
                If Len(strSenmei) > 0 Then
                    strSql = strSql & ",'" & strSenmei & "' "
                Else
                    strSql = strSql & ",null "
                End If
                If Len(datBLDate) > 0 Then
                    strSql = strSql & ",'" & datBLDate & "' "
                Else
                    strSql = strSql & ", null "
                End If
                strSql = strSql & "," & intHaisoType & "," & intHaisosaki1 & "," & intHaisosaki2 & "," & intHaisosaki3 & ","
                
                If Len(strContainerKbn) > 0 Then
                    strSql = strSql & "'" & strContainerKbn & "', "
                Else
                    strSql = strSql & "NULL, "
                End If
                If Len(strContainerNo) > 0 Then
                    strSql = strSql & "'" & strContainerNo & "',"
                Else
                    strSql = strSql & "null,"
                End If
                If Len(strSealNo) > 0 Then
                    strSql = strSql & "'" & strSealNo & "',"
                Else
                    strSql = strSql & "null,"
                End If
                strSql = strSql & dblCartonQty & "," & dblTotalQty & ",'" & datHaisoYotei & "', " & Weekday(datHaisoYotei) & ", "
                strSql = strSql & intOtsunakaNo & ", 3, 1 ) "
                
                With cmd
                    .CommandText = strSql
                    .CommandType = adCmdText
                    .Execute
                End With
                
                strInvoiceNo = rs1![pk_invoice_no].Value
                If Not IsNull(rs1![haiso_type].Value) Then
                    intHaisoType = rs1![haiso_type].Value
                Else
                    intHaisoType = 0
                End If
                If Not IsNull(rs1![haiso_saki1].Value) Then
                    intHaisosaki1 = rs1![haiso_saki1].Value
                Else
                    intHaisosaki1 = 0
                End If
                If Not IsNull(rs1![haiso_saki2].Value) Then
                    intHaisosaki2 = rs1![haiso_saki2].Value
                Else
                    intHaisosaki2 = 0
                End If
                If Not IsNull(rs1![haiso_saki3].Value) Then
                    intHaisosaki3 = rs1![haiso_saki3].Value
                Else
                    intHaisosaki3 = 0
                End If
                If Not IsNull(rs1![container_type].Value) Then
                    intContainerKbn = rs1![container_type].Value
                    strContainerKbn = rs1![container_type_text].Value
                Else
                    intContainerKbn = 0
                    strContainerKbn = ""
                End If
                If Not IsNull(rs1![haiso_yotei_ymd].Value) Then
                    datHaisoYotei = rs1![haiso_yotei_ymd].Value
                Else
                    datHaisoYotei = Date
                End If
                If Not IsNull(rs1![container_no].Value) Then
                    strContainerNo = rs1![container_no].Value
                Else
                    strContainerNo = ""
                End If
                If Not IsNull(rs1![seal_no].Value) Then
                    strSealNo = rs1![seal_no].Value
                Else
                    strSealNo = ""
                End If
                If Not IsNull(rs1![otsunaka_no].Value) Then
                    intOtsunakaNo = rs1![otsunaka_no].Value
                End If
                dblCartonQty = 0
                dblTotalQty = 0
            End If
            If strInvoiceNo <> rs1![pk_invoice_no].Value Then
                If IsNumeric(Mid(rs1![pk_invoice_no].Value, 9, 1)) Then
                    strInvoiceNo = strInvoiceNo & "," & Mid(rs1![pk_invoice_no], 10, Len(rs1![pk_invoice_no].Value) - 9)
                Else
                    strInvoiceNo = strInvoiceNo & "," & Mid(rs1![pk_invoice_no], 9, Len(rs1![pk_invoice_no].Value) - 8)
                End If
            End If
            dblCartonQty = dblCartonQty + rs1![carton_qty].Value
            dblTotalQty = dblTotalQty + rs1![total_suryo].Value
            
            If rs1![haiso_type].Value <> 0 And Not IsNull(rs1![haiso_type].Value) Then
                intTblHaisoType(rs1![haiso_type].Value - 1) = rs1![haiso_type].Value
            End If
            If Not IsNull(rs1![haiso_type_text].Value) Then
                strTblHaisoTypeText(rs1![haiso_type].Value - 1) = rs1![haiso_type_text].Value
            End If
            If Not IsNull(rs1![haiso_saki3].Value) And rs1![haiso_saki3].Value <> 0 Then
                intTblHaisosaki(rs1![haiso_type].Value - 1) = 3
            Else
                If Not IsNull(rs1![haiso_saki2].Value) And rs1![haiso_saki2].Value <> 0 Then
                    intTblHaisosaki(rs1![haiso_type].Value - 1) = 2
                Else
                    intTblHaisosaki(rs1![haiso_type].Value - 1) = 1
                End If
            End If
            rs1.MoveNext
        Loop
    End If
    rs1.Close
    
    'ＥＯＦになったら、キーブレイク用の最後の分を書き込み
    strSql = "INSERT INTO tbl_t_shanghai_haiso_tehai_print_temp (pk_login_code, ref_no, invoice_no_main, invoice_no, "
    strSql = strSql & "vessel, bl_date, haiso_type, haiso_saki1, haiso_saki2, haiso_saki3, container_type, "
    strSql = strSql & "container_no, seal_no, total_carton, total_quantity, haisyo_yotei_ymd, week_day, otsunaka_no, "
    strSql = strSql & "nisshin_plant, logistic_plant ) VALUES ( '" & Me![login_code] & "', "
    If Len(Me![REF_NO]) > 0 And Not IsNull(Me![REF_NO]) Then
        strSql = strSql & "'" & Me![REF_NO] & "','" & Me![pk_invoice_no_main] & "','" & strInvoiceNo & "'"
    Else
        strSql = strSql & "null, '" & Me![pk_invoice_no_main] & "','" & strInvoiceNo & "'"
    End If
    If Len(strSenmei) > 0 Then
        strSql = strSql & ",'" & strSenmei & "' "
    Else
        strSql = strSql & ",null "
    End If
    If Len(datBLDate) > 0 Then
        strSql = strSql & ",'" & datBLDate & "' "
    Else
        strSql = strSql & ", null "
    End If
    strSql = strSql & "," & intHaisoType & "," & intHaisosaki1 & "," & intHaisosaki2 & "," & intHaisosaki3 & ","
    
    If Len(strContainerKbn) > 0 Then
        strSql = strSql & "'" & strContainerKbn & "', "
    Else
        strSql = strSql & "NULL, "
    End If
    If Len(strContainerNo) > 0 Then
        strSql = strSql & "'" & strContainerNo & "',"
    Else
        strSql = strSql & "null,"
    End If
    If Len(strSealNo) > 0 Then
        strSql = strSql & "'" & strSealNo & "',"
    Else
        strSql = strSql & "null,"
    End If
    strSql = strSql & dblCartonQty & "," & dblTotalQty & ",'" & datHaisoYotei & "', " & Weekday(datHaisoYotei) & ", "
    strSql = strSql & intOtsunakaNo & ", 3, 1 ) "
    
    'Debug.Print strSql
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    Set cmd = Nothing

    '進捗状況更新
    Call schedule_update(7, Date, Me![pk_invoice_no_main], 1)

    '配列を該当パターンを印刷
    For n = 0 To intHaisoTypeCount - 1
        If intTblHaisoType(n) <> 0 Then
        
            '配送タイプが１０以上はＡＩＲ分　⇒　増えたため、テキストで判断
            'If intTblHaisoType(n) >= 20 Then
            If Left(strTblHaisoTypeText(n), 3) = "AIR" Then
                DoCmd.OpenReport "R_上海輸入_配送手配_AIR", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' AND [haiso_type] >= 10 "
            Else
                '配送タイプ別レポート
                Select Case intTblHaisosaki(n)
                    Case 3
                        DoCmd.OpenReport "R_上海輸入_配送手配_定期便3", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' AND ([haiso_saki3] is not null and [haiso_saki3] <> 0 ) "
                    Case 2
                        DoCmd.OpenReport "R_上海輸入_配送手配_定期便2", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' AND ([haiso_saki2] is not null and [haiso_saki2] <> 0 ) AND ( [haiso_saki3] is null Or [haiso_saki3] = 0 ) "
                    Case Else
                        '配送タイプ番の配送手配書の印刷　⇒　配送先が１箇所
                        DoCmd.OpenReport "R_上海輸入_配送手配_定期便", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' AND ( [haiso_saki2] is null Or [haiso_saki2] = 0 ) AND ( [haiso_saki3] is null Or [haiso_saki3] = 0 ) "
                End Select
                If n = 2 Then
                    '日新用があったら、日新入荷スケジュールも印刷
                    DoCmd.OpenReport "R_上海輸入_日新入荷", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' AND [haiso_type] = 3 "
                
                    '進捗状況更新
                    Call schedule_update(12, Date, Me![pk_invoice_no_main], 1)
                    
                End If
            End If
        End If
    Next
    
End Sub

Private Sub btn_hatsuban_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Dim strEditRefNo As String
    Dim dblRefNo As Double
    
    Set cmd.ActiveConnection = conn
    
    strEditRefNo = ""
    dblRefNo = 0
    With cmd
        .CommandText = "usp_shanghai_ref_no_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "TY05-"
        .Execute
    End With
    If Not IsNull(cmd.Parameters(2)) And Not IsNull(cmd.Parameters(3)) Then
        strEditRefNo = cmd.Parameters(2) & cmd.Parameters(3)
        dblRefNo = cmd.Parameters(3)
    End If
    
    Me![REF_NO] = strEditRefNo
    
    If dblRefNo <> 0 Then
        
        With cmd
            .CommandText = "usp_shanghai_ref_no_update"
            .CommandType = adCmdStoredProc
            .Parameters(1) = "TY05-"
            .Parameters(2) = dblRefNo
            .Parameters(3) = Me![pk_invoice_no_main]
            .Parameters(4) = strEditRefNo
            .Execute
        End With
    
    End If
    
    Set cmd = Nothing
    
End Sub

Private Sub btn_invoice_kakunin_Click()

    'INVOICE明細照会フォームへ
    strFormName = "F_1_上海輸入_INVOICE_MAIN"
    intMenuNo = 102
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_meisai_koushin_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE NOを入力してください。")
        Exit Sub
    End If

    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_shanghai_container_type_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    If Not IsNull(cmd.Parameters(2)) And cmd.Parameters(2) <> 0 Then
        MsgBox ("同じサイズのコンテナが複数ある場合、ここでの更新は出来ません。")
    Else
        With cmd
            .CommandText = "usp_shanghai_container_no_update"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = 0
            .Execute
        End With
    
        MsgBox ("更新完了しました。")
        
    End If
    
    Set cmd = Nothing
    

End Sub

Private Sub btn_nyuka_meisai_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If

    Dim strSoshinInvoiceNo As String
    Dim datLastUpdateYmd As Date
    Dim intMeisaiCount As Integer
    Dim strKeyContainerNo As String

    Set cmd.ActiveConnection = conn
    
    intMeisaiCount = 0
    intShinkiFlg = 0
    strSoshinInvoiceNo = ""
    With cmd
        .CommandText = "usp_shanghai_nyuka_meisai_kensu"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![shanghai_chotatsu]
        .Parameters(2) = Me![pk_invoice_no_main]
        .Parameters(3) = "%%"
        .Parameters(4) = "%%"
        .Execute
    End With
    If Not IsNull(cmd.Parameters(5)) Then
        datLastUpdateYmd = cmd.Parameters(5)
    End If
    If Not IsNull(cmd.Parameters(6)) Then
        intMeisaiCount = cmd.Parameters(6)
    End If
    
    If intMeisaiCount <> 0 Then
        ReturnValue = InputBox("現在作成済みのINVOICEの最終更新日は" & datLastUpdateYmd & "です。" & vbLf & "1:作り直す　2：内容を確認する　3：印刷する", , 3)
        If Len(ReturnValue) <= 0 Or (ReturnValue <= 1 And ReturnValue > 3) Then
            MsgBox ("１～３を選択してください。")
            intShinkiFlg = 0
        Else
            intShinkiFlg = ReturnValue
        End If
    Else
        ReturnValue = MsgBox("新規作成します", vbYesNo)
        If ReturnValue = vbYes Then
            intShinkiFlg = 1
        Else
            intShinkiFlg = 0
        End If
    End If
    If intShinkiFlg = 0 Then
        GoTo Exit_btn_nyuka_meisai_Click
    End If
    
    DoCmd.Hourglass True
    If intShinkiFlg = 1 Then
    
        '作り直し
        With cmd
            .CommandText = "usp_shanghai_nyuka_meisai_create"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = Me![shanghai_chotatsu]
            .Parameters(3) = Me![pk_invoice_no_main]
            .Parameters(4) = "%%"
            .Parameters(5) = "%%"
            .Execute
        End With
        
        intShinkiFlg = 3
        
    End If
    
    'テンポラリに出力
    With cmd
        .CommandText = "usp_shanghai_nyuka_meisai_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Parameters(3) = Me![pk_invoice_no_main]
        .Parameters(4) = "%%"
        .Parameters(5) = "%%"
        .Execute
    End With
    
    DoCmd.Hourglass False
    
    Set cmd = Nothing
    
    If intShinkiFlg = 3 Then

        '進捗状況更新
        Call schedule_update(10, Date, Me![pk_invoice_no_main], 1)

        DoCmd.OpenReport "R_上海タジマ入荷明細", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' "
        
    End If
    
Exit_btn_nyuka_meisai_Click:
    
    Set cmd = Nothing
    
    If intShinkiFlg = 2 Then
        'MsgBox ("購買発注明細の処理に移動して、内容を確認してください。")
        Call btn_hacchu_kakunin_Click
    End If
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_1_上海輸入_INVOICE_HEADER")
    intMenuNo = 103
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
    
    'intMenuNo = 103
    'Me![pk_invoice_no_main] = Null
    
    'strFormName = "F_1_上海輸入_INVOICE_HEADER"
    'Call next_form_open

End Sub

Private Sub btn_torikomi_Click()

    '購買発注取込処理へ
    strFormName = "F_1_R3購買発注取込"
    intMenuNo = 104
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_touroku_Click()

    If Not IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    DoCmd.GoToControl "pk_invoice_no_main"
    
    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    If IsNull(Me![vessel]) Then
        MsgBox ("船名を入力してください。")
        Exit Sub
    End If
    
    If IsNull(Me![otsunaka_no]) Then
        MsgBox ("乙仲を入力してください。")
        Exit Sub
    End If
    
    Dim intContainerCount As Integer
    Dim strKeyContainerNo As String
    Dim strTblEditContainer() As String
    Dim strTblContainerNo() As String
    Dim intHaisoYoteiFlg As Integer
    
    
    Dim strKeySansyoNo As String
    Dim strKeyPlant As String
    Dim strKeyHokanBasyo As String
    Dim intSeqNo As Integer
    Dim intMeisaiNo As Integer
    
    
    Dim strPara() As String
    Dim intPara() As Integer
    
    
    Set cmd.ActiveConnection = conn
    
    'ヘッダ部作成
    t = 9
    
    ReDim strPara(t) As String
    ReDim intPara(t) As Integer
    
    strPara(0) = "pk_invoice_no_main":     intPara(0) = 0
    strPara(1) = "REF_NO":                 intPara(1) = 0
    strPara(2) = "bl_date":                intPara(2) = 0
    strPara(3) = "vessel":                 intPara(3) = 0
    strPara(4) = "syukka_hoho":            intPara(4) = 1
    strPara(5) = "carton_qty":             intPara(5) = 1
    strPara(6) = "return_pallet":          intPara(6) = 1
    strPara(7) = "otsunaka_no":            intPara(7) = 1
    strPara(8) = "update_ymd":             intPara(8) = 1
    strPara(9) = "update_login":           intPara(9) = 1
    
    With cmd
        .CommandText = "usp_shanghai_invoice_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(2)) And cmd.Parameters(2) <> 0 Then
        intHeaderFlg = 1
    Else
        intHeaderFlg = 0
    End If
    
    'ヘッダ作成
    If intHeaderFlg = 0 Then
        '新規
        strSql = "INSERT INTO tbl_t_shanghai_invoice_header ("
        For n = 0 To t
            If n <> 0 Then
                strSql = strSql & ","
            End If
            strSql = strSql & strPara(n)
        Next
        strSql = strSql & ") VALUES ("
        For n = 0 To t - 2
            If n <> 0 Then
                strSql = strSql & ","
            End If
            If IsNull(Me(strPara(n))) Or Len(Me(strPara(n))) <= 0 Then
                strSql = strSql & "null"
            Else
                If intPara(n) = 0 Then
                    strSql = strSql & "'" & Me(strPara(n)) & "'"
                Else
                    strSql = strSql & Me(strPara(n))
                End If
            End If
        Next
        strSql = strSql & ", '" & Now() & "', '" & Me![login_code] & "' )"
                
    Else
    
        '更新
        strSql = "UPDATE tbl_t_shanghai_invoice_header SET "
        For n = 0 To t - 2
            If n <> 0 Then
                strSql = strSql & ","
            End If
            If intPara(n) <> 3 Then
                If IsNull(Me(strPara(n))) Or Len(Me(strPara(n))) <= 0 Then
                    strSql = strSql & strPara(n) & " = null"
                Else
                    If intPara(n) = 0 Then
                        strSql = strSql & strPara(n) & " = '" & Me(strPara(n)) & "'"
                    Else
                        strSql = strSql & strPara(n) & " = " & Me(strPara(n))
                    End If
                End If
            End If
        Next
        strSql = strSql & ", update_ymd = '" & Now() & "', update_login = '" & Me![login_code] & "' "
        strSql = strSql & "WHERE pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
        
    End If
    
    'Debug.Print strSql
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    'サブヘッダ更新
    With cmd
        .CommandText = "usp_shanghai_invoice_header_sub_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    'ヘッダサブの印刷用のINVOICE_NOの編集
    'コンテナNOの件数
    intContainerCount = 0
    With cmd
        .CommandText = "usp_shanghai_invoice_container_count"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    If Not IsNull(cmd.Parameters(2)) Then
        intContainerCount = cmd.Parameters(2)
    End If
    
    ReDim strTblEditContainer(intContainerCount) As String
    ReDim strTblContainerNo(intContainerCount) As String
    
    For n = 0 To intContainerCount
        strTblEditContainer(n) = ""
        strTblContainerNo(n) = ""
    Next
    
    intHaisoYoteiFlg = 0
    'サブヘッダを読み込んで印刷用INVOICE_NOを編集する、また入荷予定日の有無の確認（この後の処理用）
    strSql = "SELECT pk_invoice_no, container_type, container_no, haiso_yotei_ymd FROM tbl_t_shanghai_invoice_header_sub WHERE pk_invoice_no_main = '"
    strSql = strSql & Me![pk_invoice_no_main] & "' AND haiso_type <> 3 ORDER BY container_type, container_no, pk_invoice_no "
    
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    Set rs1 = cmd.Execute(intRcnt)
    
    n = 0
    If intRcnt <> 0 Then
        rs1.MoveFirst
        If rs1![container_type].Value <> 0 And Not IsNull(rs1![container_type].Value) _
            And Not IsNull(rs1![container_no].Value) And Len(rs1![container_no].Value) > 0 Then
            strKeyInvoiceNo = rs1![pk_invoice_no].Value
            strKeyContainerNo = rs1![container_no].Value
            If IsNumeric(Mid(rs1![pk_invoice_no].Value, 9, 1)) Then
                strTblEditContainer(n) = Right(rs1![pk_invoice_no], Len(rs1![pk_invoice_no].Value) - 9)
            Else
                strTblEditContainer(n) = Right(rs1![pk_invoice_no], Len(rs1![pk_invoice_no].Value) - 8)
            End If
            strTblContainerNo(n) = rs1![container_no].Value
        End If
        Do Until rs1.EOF
            If Not IsNull(rs1![container_type].Value) And rs1![container_type].Value <> 0 _
                And Not IsNull(rs1![container_no].Value) And Len(rs1![container_no].Value) > 0 Then
                If strKeyContainerNo <> rs1![container_no].Value Then
                    n = n + 1
                    strKeyInvoiceNo = rs1![pk_invoice_no].Value
                    strKeyContainerNo = rs1![container_no].Value
                    strTblContainerNo(n) = rs1![container_no].Value
                    If IsNumeric(Mid(rs1![pk_invoice_no].Value, 9, 1)) Then
                        strTblEditContainer(n) = Right(rs1![pk_invoice_no], Len(rs1![pk_invoice_no]) - 9)
                    Else
                        strTblEditContainer(n) = Right(rs1![pk_invoice_no], Len(rs1![pk_invoice_no]) - 8)
                    End If
                End If
                If strKeyInvoiceNo <> rs1![pk_invoice_no].Value Then
                    If IsNumeric(Mid(rs1![pk_invoice_no].Value, 9, 1)) Then
                        strTblEditContainer(n) = strTblEditContainer(n) & "," & Right(rs1![pk_invoice_no], Len(rs1![pk_invoice_no]) - 9)
                    Else
                        strTblEditContainer(n) = strTblEditContainer(n) & "," & Right(rs1![pk_invoice_no], Len(rs1![pk_invoice_no]) - 8)
                    End If
                End If
            End If
            If Not IsNull(rs1![haiso_yotei_ymd].Value) Then
                intHaisoYoteiFlg = 1
            End If
            rs1.MoveNext
        Loop
    End If
    rs1.Close
    'サブヘッダ更新
    For n = 0 To intContainerCount - 1
        strSql = "UPDATE tbl_t_shanghai_invoice_header_sub SET invoice_no_print = '" & strTblEditContainer(n) & "' "
        strSql = strSql & "WHERE tbl_t_shanghai_invoice_header_sub.container_no = '" & strTblContainerNo(n) & "' "
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
    Next
    
    If intHaisoYoteiFlg <> 0 Then
    
        Dim URL As String, IE As Object
        Set IE = CreateObject("InternetExplorer.Application")
        
        URL = "http://172.20.136.216:7700/dataspider/trigger/KJH_R3_ZM510?MainInvoiceNo=" & Me![pk_invoice_no_main]
        With IE
            .Navigate (URL)
            .Visible = False
        End With
        Set IE = Nothing
    
    End If
            
    '購買発注データの支給包装材の納入期日をINVOICEヘッダの配送予定日で更新にする
    '(支給包装材は購買発注データは元々無く、システムで作成しているため、納期はここで更新しておく）
    With cmd
        .CommandText = "usp_shanghai_invoice_shikyu_hosozai_nouki_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    
    'POEMの巻き芯入出庫情報更新
    With cmd
        .CommandText = "usp_poem_makishin_nyu_syukko_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    
    MsgBox ("登録完了しました。")
    
    Set rs1 = Nothing
    Set cmd = Nothing
    
    'Debug.Print intShanghaiChotatsu
    
    Call pk_invoice_no_main_AfterUpdate
    
    'Debug.Print intShanghaiChotatsu

End Sub

Private Sub btn_zaiko_tenso_Click()

    '在庫転送作成処理へ
    strFormName = "F_1_R3在庫転送_MAIN"
    intMenuNo = 106
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 103
    intMenuNoSeq = 0
    
    'Debug.Print "header open flg=" & intShanghaiChotatsu
    
End Sub

Private Sub pk_invoice_no_main_AfterUpdate()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    'Debug.Print "before header invoice pickup /flg=" & intShanghaiChotatsu
    
    Dim strEditRefNo As String
    Dim dblRefNo As Double
    
    Me![syukka_hoho].BackColor = "8454143"
    Me![BL_DATE].BackColor = "8454143"
    Me![syukka_hoho].Requery
    Me![BL_DATE].Requery
    
    Me![vessel] = Null
    Me![otsunaka_no] = Null
    Me![total_amount] = Null
    Me![total_suryo] = Null
    Me![carton_qty] = Null
    Me![syukka_hoho] = Null
    Me![BL_DATE] = Null
    Me![return_pallet] = Null

    If IsNull(Me![pk_invoice_no_main]) Then
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    intHeaderFlg = 0
    strEditRefNo = ""
    intInvoiceHeaderCount = 0
    intInvoiceSubCount = 0
    intShinkiFlg = 0
    
    '登録済みか
    With cmd
        .CommandText = "usp_shanghai_invoice_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    'ヘッダの件数
    If Not IsNull(cmd.Parameters(2)) Then
        intInvoiceHeaderCount = cmd.Parameters(2)
    End If
    'サブヘッダの件数
    If Not IsNull(cmd.Parameters(3)) And cmd.Parameters(3) <> 0 Then
        intInvoiceSubCount = cmd.Parameters(3)
    Else
        '無かったら「新規」というフラグを立てる
        intShinkiFlg = 1
    End If
    
    'BL DATEと出荷方法が１INVOICEの中で複数あるかチェック
    If Not IsNull(cmd.Parameters(4)) And cmd.Parameters(4) > 1 Then
        MsgBox ("同じINVOICE_NOで、BL_DATE、出荷方法が異なるPACKINGがあります。" & vbLf & "PACKINGデータを修正してください。")
        GoTo Exit_pk_invoice_no_main_AfterUpdate
    End If
    
    If intInvoiceHeaderCount = 0 Then
        MsgBox ("新規登録です。")
        'MsgBox ("対象INVOICEデータはありませんでした。")
        On Error Resume Next
        For Each ct In Me
            If ct.Name <> "pk_invoice_no_main" And Left(ct.Name, 5) <> "LOGIN" And ct.Name <> "bumon_code" And ct.Name <> "return_form_name" _
                And ct.Name <> "shanghai_chotatsu" Then
                Me(ct.Name) = Null
            End If
        Next
        On Error GoTo 0
    Else
    
        'INVOICEヘッダ情報の読み込み　⇒　FORMに表示
        strSql = "SELECT * FROM tbl_t_shanghai_invoice_header WHERE pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
        End With
        Set rs1 = cmd.Execute(intRcnt)
        On Error Resume Next
        If intRcnt <> 0 Then
            intHeaderFlg = 1
            rs1.MoveFirst
            strEditRefNo = rs1![REF_NO].Value
            For Each ct In Me
                For n = 0 To rs1.Fields.Count - 1
                    If ct.Name = rs1.Fields(n).Name Then
                        Me(ct.Name) = rs1.Fields(n).Value
                        n = rs1.Fields.Count - 1
                    End If
                Next
            Next
        End If
        rs1.Close
    End If
    
    '既存SUBヘッダ情報の抽出、テンポラリ作成(parameters(3)=0）
    With cmd
        .CommandText = "usp_shanghai_invoice_header_sub_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Parameters(3) = Me![pk_invoice_no_main]
        .Parameters(4) = intShinkiFlg
        .Execute
    End With
    
    '明細から合計情報を抽出してヘッダ情報作成
    With cmd
        .CommandText = "usp_shanghai_invoice_gokei_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(4)) Then
        Me![carton_qty] = cmd.Parameters(4)
    End If
    If Not IsNull(cmd.Parameters(5)) Then
        Me![total_amount] = cmd.Parameters(5)
    End If
    If Not IsNull(cmd.Parameters(6)) Then
        Me![total_suryo] = cmd.Parameters(6)
    End If
    If intInvoiceHeaderCount = 0 Then
        If Not IsNull(cmd.Parameters(2)) Then
            Me![BL_DATE] = cmd.Parameters(2)
        End If
        If Not IsNull(cmd.Parameters(3)) Then
            Me![syukka_hoho] = cmd.Parameters(3)
        End If
    End If
    
    '初番はここではしないでおこう
    'If Len(strEditRefNo) <= 0 And Not IsNull(Me![syukka_hoho]) And Me![syukka_hoho] = 0 Then
    '    strSql = "SELECT [pk_ref_no], [SEQ_NO] FROM tbl_t_shanghai_ref_no_kanri  "
    '    'Set rs1 = db.OpenRecordset(strSql)
    '    If Not rs1.BOF And Not rs1.EOF Then
    '        rs1.MoveFirst
    '        If Not IsNull(rs1![pk_ref_no].Value) And rs1![pk_ref_no].Value = "TY05-" Then
    '            strEditRefNo = rs1![pk_ref_no].Value & rs1![SEQ_NO].Value
    '            dblRefNo = rs1![SEQ_NO].Value
    '        End If
    '    End If
    '    rs1.Close
        
    '    me![ref_no] = strEditRefNo
        
    'End If
    
    'If dblRefNo <> 0 Then
    '    DoCmd.SetWarnings False
        
    '    strSql = "UPDATE tbl_t_shanghai_ref_no_kanri SET pk_seq_no = '" & dblRefNo + 1 & "' WHERE pk_ref_no = 'TY05-' "
    '    DoCmd.RunSQL strSql
        
    '    strSql = "UPDATE tbl_t_shanghai_invoice_header SET REF_NO = '" & strEditRefNo & "' WHERE [pk_invoice_no_main] = '" & Me![pk_invoice_no_main] & "' "
    '    DoCmd.RunSQL strSql
        
    '    DoCmd.SetWarnings True
    'End If
    
    If intInvoiceHeaderCount = 0 Then
        Me![syukka_hoho].BackColor = "12632256"
        Me![BL_DATE].BackColor = "12632256"
        Me![syukka_hoho].Requery
        Me![BL_DATE].Requery
    End If
    
    Me![F_1_上海輸入_INVOICE_HEADER_SUB].Requery
    
    'Debug.Print "after header invoice pickup /flg=" & intShanghaiChotatsu
    
Exit_pk_invoice_no_main_AfterUpdate:
    
    Set rs1 = Nothing
    Set cmd = Nothing

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
    If Not IsNull(Me![pk_invoice_no_main]) Then
        strInvoiceNo = Me![pk_invoice_no_main]
    Else
        strInvoiceNo = "NULL"
    End If
    strOpenMode = "OC"
    
    Call form_open_close("F_1_上海輸入_INVOICE_HEADER", strFormName, strOpenMode, strInvoiceNo)
    
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
    
    intMenuNo = 103
    Me![pk_invoice_no_main] = Null
    
    strFormName = "F_1_上海輸入_INVOICE_HEADER"
    Call next_form_open

End Sub
