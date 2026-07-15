Attribute VB_Name = "Form_F_1_SAP_AMOUNT_MAIN_backup"
Attribute VB_Base = "0{7A6BD94D-8EE3-4E57-AE13-179182D98163}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit


Private Sub btn_excel_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![chk_display_kbn]) Or Me![chk_display_kbn] = 0 Then
        Me![chk_display_kbn] = 1
    End If
    intMenuNoSeq = chk_display_kbn

    If intShanghaiChotatsu = 2 Then
        intMenuNo = 205
    Else
        intMenuNo = 107
    End If
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

Private Sub btn_create_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    '仕入れ明細作成

    If IsNull(Me![taisyo_yyyy]) Or IsNull(Me![taisyo_mm]) Or Len(Me![taisyo_yyyy]) <= 0 Or Len(Me![taisyo_mm]) <= 0 Then
        MsgBox ("対象年月を入力してください。")
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    '既存データの削除（抽出の条件と同じ条件のデータを削除する）
    strSql = "DELETE tbl_t_kaigai_shiire_jisseki_ruikei FROM tbl_t_kaigai_shiire_jisseki_ruikei "
    strSql = strSql & "WHERE pk_yyyymm = " & Me![taisyo_yyyymm] & " "
    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        strSql = strSql & "AND pk_invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
    End If
    If Not IsNull(Me![shiiresaki_code]) Then
        strSql = strSql & "AND shiiresaki_code = '" & Me![shiiresaki_code] & "' "
    End If
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    '追加
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    Set cmd = Nothing
    
    Call chk_display_kbn_Click
    
    MsgBox ("集計完了しました、内容をご確認ください。")
    
End Sub

Private Sub btn_hacchu_check_Click()

    '購買発注明細確認へ
    strFormName = "F_1_R3購買発注_MAIN"
    If intShanghaiChotatsu = 2 Then
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

    '海上保険依頼用明細抽出

    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        '進捗状況更新
        Call schedule_update(15, Date, Me![INVOICE_NO_MAIN], intShanghaiChotatsu)
    End If
    
    If intShanghaiChotatsu = 2 Then
        intMenuNo = 205
    Else
        intMenuNo = 107
    End If
    intMenuNoSeq = 3
    strInputCode = Me![taisyo_yyyymm]
    
    strFormName = "F_9_メール送信_MAIN"
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
    
    Dim strShiiresakiCode As String
    
    Set cmd.ActiveConnection = conn
    
    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        strInvoiceNo = Me![INVOICE_NO_MAIN]
    Else
        strInvoiceNo = "%%"
    End If
    
    If Not IsNull(Me![shiiresaki_code]) Then
        strShiiresakiCode = Me![shiiresaki_code]
    Else
        strShiiresakiCode = "%%"
    End If
    If Not IsNull(Me![chk_shiiresaki_kbn]) Then
        If Me![chk_shiiresaki_kbn] = 2 Then
            strShiiresakiCode = "133128"
        Else
            If Me![chk_shiiresaki_kbn] = 3 Then
                strShiiresakiCode = "143534"
            End If
        End If
    End If
    
    '登録済みがあるか
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![taisyo_yyyymm]
        .Parameters(2) = strInvoiceNo
        .Parameters(3) = strShiiresakiCode
        .Execute
    End With
    If Not IsNull(cmd.Parameters(4)) And cmd.Parameters(4) <> 0 Then
        intShinkiFlg = 2
    Else
        intShinkiFlg = 1
    End If
        
    '抽出
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = intShanghaiChotatsu
        .Parameters(3) = intShinkiFlg
        .Parameters(4) = Me![taisyo_yyyymm]
        .Parameters(5) = strInvoiceNo
        .Parameters(6) = strShiiresakiCode
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
    
    Set cmd = Nothing
    
    Call chk_display_kbn_Click
    
        

End Sub

Private Sub btn_print_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
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
        .Execute
    End With
        
    If Not IsNull(cmd.Parameters(2)) Then
        intOthersCount = cmd.Parameters(2)
    End If
    
    Set cmd = Nothing
    
    'マスタ外品目が合ったら｢その他明細｣の印刷
    If intOthersCount <> 0 Then
        DoCmd.OpenReport "R_上海からの仕入明細書_その他", acViewPreview, , "pk_login_code = " & Me![login_code]
    End If

    '仕入れ明細の印刷
    DoCmd.OpenReport "R_上海からの仕入明細書", acViewPreview, , "pk_login_code = " & Me![login_code]
    DoCmd.OpenReport "R_上海からの仕入明細書_INVOICE別", acViewPreview, , "pk_login_code = " & Me![login_code]
    
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

    If Me![btn_shanghai_chotatsu].Caption = "上海" Then
        Me![btn_shanghai_chotatsu].Caption = "調達"
        intShanghaiChotatsu = 2
        intMenuNo = 205
    Else
        Me![btn_shanghai_chotatsu].Caption = "上海"
        intShanghaiChotatsu = 1
        intMenuNo = 107
    End If
    
    Me![INVOICE_NO_MAIN] = Null
    
    strFormName = "F_1_SAP_AMOUNT_MAIN"
    Call next_form_open

End Sub

Private Sub btn_zaiko_tenso_Click()

    '在庫転送処理へ
    strFormName = "F_1_R3在庫転送_MAIN"
    If intShanghaiChotatsu = 2 Then
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
        Me![taisyo_yyyymm] = Me![taisyo_yyyy] * 100 + Me![taisyo_mm]
    End If
    
    If intShanghaiChotatsu = 2 Then
        strSql = "SELECT pk_invoice_no_main FROM dbo.tbl_t_chotatsu_invoice_header WHERE bl_date >= '"
        strSql = strSql & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm], 1) & "' AND bl_date < '"
        strSql = strSql & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm] + 1, 1) - 1 & "' "
        strSql = strSql & "GROUP BY pk_invoice_no_main ORDER BY pk_invoice_no_main DESC "
    Else
        strSql = "SELECT pk_invoice_no_main FROM dbo.tbl_t_shanghai_invoice_header WHERE bl_date >= '"
        strSql = strSql & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm], 1) & "' AND bl_date < '"
        strSql = strSql & DateSerial(Me![taisyo_yyyy], Me![taisyo_mm] + 1, 1) - 1 & "' "
        strSql = strSql & "GROUP BY pk_invoice_no_main ORDER BY pk_invoice_no_main DESC "
    End If
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
    
    Me![taisyo_yyyy] = Year(Date)
    Me![taisyo_mm] = Month(Date)
    Me![taisyo_yyyymm] = Year(Date) * 100 + Month(Date)
    
    Call chk_display_kbn_Click
    'Me![F_1_SAP_AMOUNT_SUB].Requery
    'Me.Requery

End Sub

Private Sub invoice_no_main_AfterUpdate()
    
    Call chk_display_kbn_Click

End Sub

Private Sub taisyo_mm_LostFocus()

    Call chk_display_kbn_Click

End Sub

Private Sub taisyo_yyyy_LostFocus()

    Call chk_display_kbn_Click

End Sub

Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    strInvoiceNo = ""
    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        strInvoiceNo = Me![INVOICE_NO_MAIN]
    Else
        strInvoiceNo = "NULL"
    End If
    If strFormName = "F_1_SAP_AMOUNT_MAIN" Then
        strOpenMode = "O"
    Else
        strOpenMode = "OC"
    End If
    
    Call form_open_close("F_1_SAP_AMOUNT_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub
