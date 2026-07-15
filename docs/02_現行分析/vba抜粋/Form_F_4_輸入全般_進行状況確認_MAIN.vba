Attribute VB_Name = "Form_F_4_輸入全般_進行状況確認_MAIN"
Attribute VB_Base = "0{2903D5DD-9D86-46FE-B8FA-1652CA22936C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

    Dim strYmdName As String
    Dim datYmdValue As Date
    Dim intYmdFlg As Integer
    Dim strFieldName As String


Private Sub access_ymd_AfterUpdate()

    strYmdName = "access_ymd"
    
    Call ymd_update
    
End Sub

Private Sub bnt_menu_Click()
    
    'メニューに戻る

    If IsNull(Me![return_form_name]) Then
        strFormName = "F_0_START_MENU"
    Else
        strFormName = Me![return_form_name]
    End If
    intMenuNo = 0
    Call next_form_open
    
End Sub


Private Sub btn_delete_Click()

    ReturnValue = MsgBox("削除しますか？", 4)
    If ReturnValue = vbNo Then
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    strSql = "DELETE tbl_t_import_progress FROM tbl_t_import_progress WHERE pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
    
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    Set cmd = Nothing
    
    Me![pk_invoice_no_main].Requery
    Me![F_4_輸入全般_進行状況確認_SUB].Requery
    
    Call shinko_jyokyo_display(Me![pk_invoice_no_main])
    
    MsgBox ("削除完了しました。")
    
    
End Sub

Private Sub btn_excel_Click()

    If IsNull(Me![taisyo_yyyymm]) Then
        MsgBox ("対象年月を入力してください。")
        Exit Sub
    End If
    
    '社内参照番号抽出
    intMenuNo = 403
    If intShanghaiChotatsu = 2 Then
        intMenuNoSeq = 2
    Else
        intMenuNoSeq = 1
    End If
    strInputCode = Me![taisyo_yyyymm]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_print_Click()

    If intShanghaiChotatsu <> 2 Then
        DoCmd.OpenReport "R_進行状況確認", acViewPreview, , "taisyo_yyyymm = '" & Me![taisyo_yyyymm] & "' AND syurui <= 2 "
    Else
        DoCmd.OpenReport "R_進行状況確認", acViewPreview, , "taisyo_yyyymm = '" & Me![taisyo_yyyymm] & "' AND ( syurui = 3 OR syurui = 4 ) "
    End If
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_4_輸入全般_進行状況確認_MAIN")
    intMenuNo = 403
    intMenuNoSeq = 0

End Sub

Private Sub chk_access_ymd_Click()

    strYmdName = "access_ymd"
    If IsNull(Me![chk_access_ymd]) Or Me![chk_access_ymd] <> -1 Then
        Me![access_ymd] = Null
    Else
        Me![access_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_chotatsu_mail_ymd_Click()

    strYmdName = "chotatsu_mail_ymd"
    If IsNull(Me![chk_chotatsu_mail_ymd]) Or Me![chk_chotatsu_mail_ymd] <> -1 Then
        Me![chotatsu_mail_ymd] = Null
    Else
        Me![chotatsu_mail_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_haiso_boueki_ymd_Click()

    strYmdName = "haiso_boueki_ymd"
    If IsNull(Me![chk_haiso_boueki_ymd]) Or Me![chk_haiso_boueki_ymd] <> -1 Then
        Me![haiso_boueki_ymd] = Null
    Else
        Me![haiso_boueki_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_haiso_chotatsu_ymd_Click()

    strYmdName = "haiso_chotatsu_ymd"
    If IsNull(Me![chk_haiso_chotatsu_ymd]) Or Me![chk_haiso_chotatsu_ymd] <> -1 Then
        Me![haiso_chotatsu_ymd] = Null
    Else
        Me![haiso_chotatsu_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_haiso_kojyo_ymd_Click()

    strYmdName = "haiso_kojyo_ymd"
    If IsNull(Me![chk_haiso_kojyo_ymd]) Or Me![chk_haiso_kojyo_ymd] <> -1 Then
        Me![haiso_kojyo_ymd] = Null
    Else
        Me![haiso_kojyo_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_haiso_logi_ymd_Click()

    strYmdName = "haiso_logi_ymd"
    If IsNull(Me![chk_haiso_logi_ymd]) Or Me![chk_haiso_logi_ymd] <> -1 Then
        Me![haiso_logi_ymd] = Null
    Else
        Me![haiso_logi_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_haiso_nisshin_ymd_Click()

    strYmdName = "haiso_nisshin_ymd"
    If IsNull(Me![chk_haiso_nisshin_ymd]) Or Me![chk_haiso_nisshin_ymd] <> -1 Then
        Me![haiso_nisshin_ymd] = Null
    Else
        Me![haiso_nisshin_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_hoken_ymd_Click()

    strYmdName = "hoken_ymd"
    If IsNull(Me![chk_hoken_ymd]) Or Me![chk_hoken_ymd] <> -1 Then
        Me![hoken_ymd] = Null
    Else
        Me![hoken_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_keihi_ymd_Click()

    strYmdName = "keihi_ymd"
    If IsNull(Me![chk_keihi_ymd]) Or Me![chk_keihi_ymd] <> -1 Then
        Me![keihi_ymd] = Null
    Else
        Me![keihi_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_kensa_ymd_Click()

    strYmdName = "kensa_ymd"
    If IsNull(Me![chk_kensa_ymd]) Or Me![chk_kensa_ymd] <> -1 Then
        Me![kensa_ymd] = Null
    Else
        Me![kensa_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_shanghai_header_ymd_Click()

    strYmdName = "shanghai_header_ymd"
    If IsNull(Me![chk_shanghai_header_ymd]) Or Me![chk_shanghai_header_ymd] <> -1 Then
        Me![shanghai_header_ymd] = Null
    Else
        Me![shanghai_header_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_shanghai_syorui_ymd_Click()

    strYmdName = "shanghai_syorui_ymd"
    If IsNull(Me![chk_shanghai_syorui_ymd]) Or Me![chk_shanghai_syorui_ymd] <> -1 Then
        Me![shanghai_syorui_ymd] = Null
    Else
        Me![shanghai_syorui_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_syukko_ymd_Click()

    strYmdName = "syukko_ymd"
    If IsNull(Me![chk_syukko_ymd]) Or Me![chk_syukko_ymd] <> -1 Then
        Me![syukko_ymd] = Null
    Else
        Me![syukko_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_tsukan_irai_ymd_Click()

    strYmdName = "tsukan_irai_ymd"
    If IsNull(Me![chk_tsukan_irai_ymd]) Or Me![chk_tsukan_irai_ymd] <> -1 Then
        Me![tsukan_irai_ymd] = Null
    Else
        Me![tsukan_irai_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chk_tsukan_kyoka_ymd_Click()

    strYmdName = "tsukan_kyoka_ymd"
    If IsNull(Me![chk_tsukan_kyoka_ymd]) Or Me![chk_tsukan_kyoka_ymd] <> -1 Then
        Me![tsukan_kyoka_ymd] = Null
    Else
        Me![tsukan_kyoka_ymd] = Date
    End If
    
    Call ymd_update
    
End Sub

Private Sub chotatsu_mail_ymd_AfterUpdate()

    strYmdName = "chotatsu_mail_ymd"
    
    Call ymd_update
    
End Sub

Private Sub chotatsu_mail_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 103
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 403
    intMenuNoSeq = 0
    
    Me![taisyo_yyyy] = Year(Date)
    Me![taisyo_mm] = Month(Date)
    
    Call sub_form_display
    
End Sub

Private Sub haiso_boueki_ymd_AfterUpdate()

    strYmdName = "haiso_boueki_ymd"
    
    Call ymd_update
    
End Sub

Private Sub haiso_boueki_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 103
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
End Sub

Private Sub haiso_chotatsu_ymd_AfterUpdate()

    strYmdName = "haiso_chotatsu_ymd"
    
    Call ymd_update
    
End Sub

Private Sub haiso_chotatsu_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 105
    Else
        intMenuNo = 202
    End If
    Call sentaku_form_open
    
End Sub

Private Sub haiso_kojyo_ymd_AfterUpdate()

    strYmdName = "haiso_kojyo_ymd"
    
    Call ymd_update
    
End Sub

Private Sub haiso_kojyo_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 105
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
End Sub

Private Sub haiso_logi_ymd_AfterUpdate()

    strYmdName = "haiso_logi_ymd"
    
    Call ymd_update
    
End Sub

Private Sub haiso_logi_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 103
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
End Sub

Private Sub haiso_nisshin_ymd_AfterUpdate()

    strYmdName = "haiso_kojyo_ymd"
    
    Call ymd_update
    
End Sub

Private Sub haiso_nisshin_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 103
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
End Sub

Private Sub hoken_ymd_AfterUpdate()

    strYmdName = "hoken_ymd"
    
    Call ymd_update
    
End Sub

Private Sub hoken_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 107
    Else
        intMenuNo = 301
    End If
    Call sentaku_form_open
    
End Sub

Private Sub pk_invoice_no_main_AfterUpdate()

    If Not IsNull(Me![pk_invoice_no_main]) And Len(Me![pk_invoice_no_main]) > 0 Then
        Call shinko_jyokyo_display(Me![pk_invoice_no_main])
    End If
    
End Sub

Private Sub keihi_ymd_AfterUpdate()

    strYmdName = "keihi_ymd"
    
    Call ymd_update
    
End Sub

Private Sub keihi_ymd_DblClick(Cancel As Integer)

    intMenuNo = 301
    Call sentaku_form_open
    
End Sub

Private Sub kensa_ymd_AfterUpdate()

    strYmdName = "kensa_ymd"
    
    Call ymd_update
    
End Sub

Private Sub shanghai_header_ymd_AfterUpdate()

    strYmdName = "shanghai_header_ymd"
    
    Call ymd_update
    
End Sub

Private Sub shanghai_header_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 103
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
End Sub

Private Sub shanghai_syorui_ymd_AfterUpdate()

    strYmdName = "shanghai_syorui_ymd"
    
    Call ymd_update
    
End Sub

Private Sub shanghai_syorui_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 103
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
End Sub

Private Sub syukko_ymd_AfterUpdate()

    strYmdName = "syukko_ymd"
    
    Call ymd_update
    
End Sub

Private Sub ymd_update()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Or Len(Me![pk_invoice_no_main]) <= 0 Then
        MsgBox ("INVOICE_NOを入力してください｡ ")
        Exit Sub
    End If

    If IsNull(Me![syurui]) Or Len(Me![syurui]) <= 0 Then
        MsgBox ("種類を入力してください｡ ")
        Call shinko_jyokyo_display(Me![pk_invoice_no_main])
        Exit Sub
    End If

    Dim strFieldName(16) As String
    Dim strShinchokuJyokyo As String
    Dim intKanryo As Integer
    
    strFieldName(0) = "pk_invoice_no_main"
    strFieldName(1) = "syurui"
    strFieldName(2) = "syukko_ymd"
    strFieldName(3) = "shanghai_syorui_ymd"
    strFieldName(4) = "shanghai_header_ymd"
    strFieldName(5) = "kensa_ymd"
    strFieldName(6) = "tsukan_irai_ymd"
    strFieldName(7) = "chotatsu_mail_ymd"
    strFieldName(8) = "tsukan_kyoka_ymd"
    strFieldName(9) = "haiso_logi_ymd"
    strFieldName(10) = "haiso_kojyo_ymd"
    strFieldName(11) = "haiso_nisshin_ymd"
    strFieldName(12) = "haiso_boueki_ymd"
    strFieldName(13) = "haiso_chotatsu_ymd"
    strFieldName(14) = "hoken_ymd"
    strFieldName(15) = "access_ymd"
    strFieldName(16) = "keihi_ymd"

    If IsNull(Me(strYmdName)) Or Len(Me(strYmdName)) <= 0 Then
        Me("chk_" & strYmdName) = Null
        intYmdFlg = 0
    Else
        Me("chk_" & strYmdName) = -1
        intYmdFlg = 1
        datYmdValue = Me(strYmdName)
    End If
    
    If intYmdFlg = 0 Then
        ReturnValue = MsgBox("日付の取消（クリア）ですか？", 4)
        If ReturnValue = vbNo Then
            Call shinko_jyokyo_display(Me![pk_invoice_no_main])
            Exit Sub
        End If
    End If
    
    Dim intShinkiFlg As Integer
    
    Set cmd.ActiveConnection = conn
    
    strShinchokuJyokyo = "11111111111111111"
    With cmd
        .CommandText = "usp_koumoku_contents_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "SRY"
        .Parameters(2) = Me![syurui]
        .Execute
    End With
    If Not IsNull(cmd.Parameters(4)) Then
        strShinchokuJyokyo = cmd.Parameters(4)
    End If
    
    For n = 17 To 2 Step -1
        If Mid(strShinchokuJyokyo, n, 1) = "1" Then
            intKanryo = n - 2
            n = 2
        End If
    Next
    
    '存在チェック
    With cmd
        .CommandText = "usp_shinchoku_jyokyo_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If IsNull(cmd.Parameters(2)) Or cmd.Parameters(2) = 0 Then
        intShinkiFlg = 1
    Else
        intShinkiFlg = 0
    End If
    
    t = 0
    For n = 2 To 16
        If Not IsNull(Me(strFieldName(n))) And Len(Me(strFieldName(n))) > 0 Then
            t = n - 1
        End If
    Next
    
    If t >= intKanryo Then
        t = 99
    End If
    
    If intShinkiFlg = 1 Then
        strSql = "INSERT INTO tbl_t_import_progress ( pk_invoice_no_main, syurui, " & strYmdName & ", shinchoku_jyokyo, update_ymd, "
        strSql = strSql & "update_login ) VALUES ( '" & Me![pk_invoice_no_main] & "', "
        If Not IsNull(Me![syurui]) And Len(Me![syurui]) > 0 Then
            strSql = strSql & Me![syurui] & ", "
        Else
            strSql = strSql & "0, "
        End If
        If intYmdFlg = 0 Then
            strSql = strSql & "null, "
        Else
            strSql = strSql & "'" & datYmdValue & "', "
        End If
        strSql = strSql & t & ", '" & Now() & "', '" & Me![login_code] & "' ) "
    Else
        strSql = "UPDATE tbl_t_import_progress SET " & strYmdName & " = "
        If intYmdFlg = 0 Then
            strSql = strSql & "null "
        Else
            strSql = strSql & "'" & datYmdValue & "' "
        End If
        strSql = strSql & ", shinchoku_jyokyo = " & t & ", update_ymd = '" & Now() & "', update_login = '" & Me![login_code] & "' "
        strSql = strSql & "WHERE pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
    End If
    
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    Set cmd = Nothing
    
    Call shinko_jyokyo_display(Me![pk_invoice_no_main])
    Me![pk_invoice_no_main].Requery
    Me![syurui].Requery
    Me![F_4_輸入全般_進行状況確認_SUB].Requery

End Sub

Private Sub sub_form_display()

    Dim datFromYmd As Date
    Dim datToYmd As Date

    If IsNull(Me![taisyo_yyyy]) Or Len(Me![taisyo_yyyy]) <= 0 Then
        Me![taisyo_yyyy] = Year(Date)
    End If

    If IsNull(Me![taisyo_mm]) Or Len(Me![taisyo_mm]) <= 0 Then
        Me![taisyo_mm] = Month(Date)
    End If
    
    Me![taisyo_yyyymm] = Me![taisyo_yyyy] * 100 + Me![taisyo_mm]
    
    datFromYmd = DateSerial(Me![taisyo_yyyy], Me![taisyo_mm], 1)
    datToYmd = DateSerial(Me![taisyo_yyyy], Me![taisyo_mm] + 1, 1) - 1
    
    If IsNull(Me![syurui]) Or Me![syurui] <= 2 Then
        strSql = "SELECT View__invoice_no_main_header.pk_invoice_no_main FROM View__invoice_no_main_header "
        strSql = strSql & "WHERE View__invoice_no_main_header.bl_date >= '" & datFromYmd & "' "
        strSql = strSql & "AND View__invoice_no_main_header.bl_date <= '" & datToYmd & "' "
        strSql = strSql & "ORDER BY View__invoice_no_main_header.bl_date DESC "
        Me![pk_invoice_no_main].RowSource = strSql
    Else
        If Me![syurui] <= 4 Then
            strSql = "SELECT View__invoice_no_main_chotatsu.pk_invoice_no_main FROM View__invoice_no_main_chotatsu "
            strSql = strSql & "WHERE View__invoice_no_main_chotatsu.bl_date >= '" & datFromYmd & "' "
            strSql = strSql & "AND View__invoice_no_main_chotatsu.bl_date <= '" & datToYmd & "' "
            strSql = strSql & "ORDER BY View__invoice_no_main_chotatsu.bl_date DESC "
            Me![pk_invoice_no_main].RowSource = strSql
        Else
            strSql = "SELECT View__invoice_no_main_makishin.pk_invoice_no_main FROM View__invoice_no_main_makishin "
            strSql = strSql & "WHERE View__invoice_no_main_makishin.bl_date >= '" & datFromYmd & "' "
            strSql = strSql & "AND View__invoice_no_main_makishin.bl_date <= '" & datToYmd & "' "
            strSql = strSql & "ORDER BY View__invoice_no_main_makishin.bl_date DESC "
            Me![pk_invoice_no_main].RowSource = strSql
        End If
    End If
    Me![pk_invoice_no_main].Requery
    
    Me![F_4_輸入全般_進行状況確認_SUB].Requery
    
End Sub

Private Sub syukko_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 101
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
End Sub

Private Sub syurui_AfterUpdate()

    Dim strFieldName(17) As String
    Dim intBoxDisplay(2) As Integer
    
    strFieldName(0) = "pk_invoice_no_main"
    strFieldName(1) = "syurui"
    strFieldName(2) = "syukko_ymd"
    strFieldName(3) = "shanghai_syorui_ymd"
    strFieldName(4) = "shanghai_header_ymd"
    strFieldName(5) = "kensa_ymd"
    strFieldName(6) = "tsukan_irai_ymd"
    strFieldName(7) = "chotatsu_mail_ymd"
    strFieldName(8) = "tsukan_kyoka_ymd"
    strFieldName(9) = "haiso_logi_ymd"
    strFieldName(10) = "haiso_kojyo_ymd"
    strFieldName(11) = "haiso_nisshin_ymd"
    strFieldName(12) = "haiso_boueki_ymd"
    strFieldName(13) = "haiso_chotatsu_ymd"
    strFieldName(14) = "hoken_ymd"
    strFieldName(15) = "access_ymd"
    strFieldName(16) = "keihi_ymd"
    
    intBoxDisplay(0) = 0
    intBoxDisplay(1) = 0

    If Not IsNull(Me![input_type]) Then
        
        On Error Resume Next
        For n = 2 To 16
            If Mid(Me![input_type], n + 1, 1) = "1" Then
                Me(strFieldName(n)).Visible = True
                Me("chk_" & strFieldName(n)).Visible = True
                Me(strFieldName(n)).BackStyle = 1
                Me("lbl_" & strFieldName(n)).ForeColor = "16711680"
                Me("lbl_" & strFieldName(n)).Visible = True
                If n >= 3 And n <= 4 Then
                    intBoxDisplay(0) = intBoxDisplay(0) + 1
                Else
                    If n >= 9 And n <= 13 Then
                        intBoxDisplay(1) = intBoxDisplay(1) + 1
                    End If
                End If
                If n < 3 Or (n > 4 And n < 9) Or n > 13 Then
                    Me("yajirushi_" & n - 1).Visible = True
                    Me("yajirushi_" & n - 1 & "R").Visible = True
                    Me("line_" & n - 1).Visible = True
                    Me("line_" & n - 1 & "R").Visible = True
                End If
            Else
                Me(strFieldName(n)).Visible = False
                Me("chk_" & strFieldName(n)).Visible = False
                Me(strFieldName(n)).BackStyle = 0
                Me("lbl_" & strFieldName(n)).ForeColor = "8421504"
                Me("lbl_" & strFieldName(n)).Visible = False
                If n < 3 Or (n > 4 And n < 9) Or n > 13 Then
                    Me("yajirushi_" & n - 1).Visible = False
                    Me("yajirushi_" & n - 1 & "R").Visible = False
                    Me("line_" & n - 1).Visible = False
                    Me("line_" & n - 1 & "R").Visible = False
                End If
            End If
        Next
        
        On Error GoTo 0
                
    End If
    
    If intBoxDisplay(0) = 0 Then
        Me![box_shanghai].Visible = False
        Me![lbl_shanghai].Visible = False
        Me![yajirushi_2].Visible = False
    Else
        Me![box_shanghai].Visible = True
        Me![lbl_shanghai].Visible = True
        Me![yajirushi_2].Visible = True
    End If
    
    If intBoxDisplay(1) = 0 Then
        Me![box_haiso_renraku].Visible = False
        Me![lbl_haiso_renraku].Visible = False
        Me![yajirushi_8].Visible = False
    Else
        Me![box_haiso_renraku].Visible = True
        Me![lbl_haiso_renraku].Visible = True
        Me![yajirushi_8].Visible = True
    End If
    
    Call sub_form_display

End Sub

Private Sub taisyo_mm_AfterUpdate()
    
    Call sub_form_display
    
End Sub

Private Sub taisyo_yyyy_AfterUpdate()
    
    Call sub_form_display
    
End Sub

Private Sub tsukan_irai_ymd_AfterUpdate()

    strYmdName = "tsukan_irai_ymd"
    
    Call ymd_update
    
End Sub

Private Sub tsukan_irai_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 103
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
End Sub

Private Sub tsukan_kyoka_ymd_AfterUpdate()

    strYmdName = "tsukan_kyoka_ymd"
    
    Call ymd_update
    
End Sub

Private Sub tsukan_kyoka_ymd_DblClick(Cancel As Integer)

    If Left(Me![pk_invoice_no_main], 3) = "SJ-" Or Left(Me![pk_invoice_no_main], 3) = "YJ-" Or IsNull(Me![pk_invoice_no_main]) Then
        intMenuNo = 103
    Else
        intMenuNo = 203
    End If
    Call sentaku_form_open
    
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
    strOpenMode = "OC"
    
    Call form_open_close("F_4_輸入全般_進行状況確認_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

Private Sub sentaku_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    'strFormName = ""
    Select Case intMenuNo
        Case 0
            strFormName = "F_0_START_MENU"
        Case 101
            strFormName = "F_1_上海輸入_INVOICE取込"
        Case 103
            strFormName = "F_1_上海輸入_INVOICE_HEADER"
        Case 105
            strFormName = "F_1_R3購買発注_MAIN"
        Case 107
            strFormName = "F_1_SAP_AMOUNT_MAIN"
        Case 202
            strFormName = "F_1_R3購買発注_MAIN"
        Case 203
            strFormName = "F_2_調達輸入_INVOICE_HEADER"
        Case 301
            strFormName = "F_3_輸入諸経費入力_MAIN"
    End Select
    intMenuNoSeq = 0
    
    strInvoiceNo = ""
    If Not IsNull(Me![pk_invoice_no_main]) Then
        strInvoiceNo = Me![pk_invoice_no_main]
    Else
        strInvoiceNo = "NULL"
    End If
    strOpenMode = "OC"
    
    Call form_open_close("F_4_輸入全般_進行状況確認_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

