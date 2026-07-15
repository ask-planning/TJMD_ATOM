Attribute VB_Name = "Form_F_4_スケジュール検索_MAIN"
Attribute VB_Base = "0{BDF7330E-A14E-4FDD-B048-F7F104C4F257}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_clear_Click()

    Me![select_pk_invoice_no_main] = Null
    Me![select_pk_invoice_no] = Null
    Me![select_syanai_sansyo_no] = Null
    
    Me![chk_pickup_kbn] = 1
    Me![chk_display_kbn] = 1
    
    Call chk_pickup_kbn_Click
    
End Sub

Private Sub btn_menu_Click()

    'メニューに戻る
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_pickup_Click()

    'On Error GoTo Err_btn_pickup

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![chk_pickup_kbn]) Or Me![chk_pickup_kbn] = 0 Then
        MsgBox ("抽出区分を選択してください。")
        Exit Sub
    End If
    
    If IsNull(Me![taisyo_yyyy]) Or Me![taisyo_yyyy] = 0 Then
        MsgBox ("対象年を入力してください。")
        Exit Sub
    End If
    
    '条件入力値格納エリア
    Dim dblTaisyoYyyyMm As Double
    Dim intTaisyoYyyy As Double
    Dim intTaisyoMm As Double
    Dim strSelectInvoiceNoMain As String
    Dim strSelectInvoiceNo As String
    Dim strSelectSyanaiSansyoNo As String
    Dim intSelectPickupKbn As Integer
    Dim intSelectDisplayKbn As Integer
    
    Dim intPickupCount As Integer
    Dim strReturnFormName As String
    
    Set cmd.ActiveConnection = conn
    
    dblTaisyoYyyyMm = 0
    intTaisyoYyyy = 0
    intTaisyoMm = 0
    strSelectInvoiceNoMain = ""
    strSelectInvoiceNo = ""
    strSelectSyanaiSansyoNo = ""
    intSelectPickupKbn = 0
    intSelectDisplayKbn = 0
    If Not IsNull(Me![taisyo_yyyymm]) Then
        dblTaisyoYyyyMm = Me![taisyo_yyyymm]
    End If
    If Not IsNull(Me![taisyo_yyyy]) Then
        intTaisyoYyyy = Me![taisyo_yyyy]
    End If
    If Not IsNull(Me![taisyo_mm]) Then
        intTaisyoMm = Me![taisyo_mm]
    End If
    
    If Not IsNull(Me![select_pk_invoice_no_main]) Then
        strSelectInvoiceNoMain = Me![select_pk_invoice_no_main]
    End If
    If Not IsNull(Me![select_pk_invoice_no]) Then
        strSelectInvoiceNo = Me![select_pk_invoice_no]
    End If
    If Not IsNull(Me![select_syanai_sansyo_no]) Then
        strSelectSyanaiSansyoNo = Me![select_syanai_sansyo_no]
    End If
    intSelectPickupKbn = Me![chk_pickup_kbn]
    If Not IsNull(Me![chk_display_kbn]) Then
        intSelectDisplayKbn = Me![chk_display_kbn]
    End If
    
    strLoginCode = Me![login_code]
    strLoginName = Me![login_name]
    strBumonCode = Me![bumon_code]
    If Not IsNull(Me![return_form_name]) Then
        strReturnFormName = Me![return_form_name]
    End If
    
    
    '上海輸入用　or　調達輸入用
    If Me![chk_pickup_kbn] = 1 Then
        strViewName = "View_Frm_schedule_kensaku_sh_invoice"
    Else
        strViewName = "tbl_t_chotatsu_invoice_header"
    End If
    
    strSql = "SELECT " & strViewName & ".pk_invoice_no_main, " & strViewName & ".bl_date, " & strViewName & ".haiso_yotei_ymd, "
    If Me![chk_display_kbn] = 2 Then
        If Me![chk_pickup_kbn] = 1 Then
            strSql = strSql & strViewName & ".pk_invoice_no_grp AS pk_invoice_no, "
        Else
            strSql = strSql & strViewName & ".pk_invoice_no_main AS pk_invoice_no, "
        End If
    End If
    If Me![chk_pickup_kbn] = 1 Then
        strSql = strSql & strViewName & ".invoice_no_print, "
    Else
        strSql = strSql & strViewName & ".pk_invoice_no_main AS invoice_no_print, "
    End If
    If Me![chk_display_kbn] = 2 Then
        strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.syanai_sansyo_no_tenso_grp AS syanai_sansyo_no_tenso, "
        strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.nounyu_plant_min AS nounyu_plant, "
        strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.pk_koubai_hacchu_no_min AS pk_koubai_hacchu_no, "
        strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.zaiko_tenso_no_min AS zaiko_tenso_no, "
        If Me![chk_pickup_kbn] = 1 Then
            strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.syukka_denpyo_no_min AS syukka_denpyo_no, "
        Else
            strSql = strSql & "tbl_t_chotatsu_syukka_denpyo.pk_syukka_denpyo_no AS syukka_denpyo_no, "
        End If
        strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.R3_soshin_ymd_min AS R3_soshin_ymd "
    Else
        strSql = strSql & "MIN(dbo.View_Frm_schedule_kensaku_hacchu.syanai_sansyo_no_tenso_grp) AS syanai_sansyo_no_tenso, "
        strSql = strSql & "MIN(dbo.View_Frm_schedule_kensaku_hacchu.nounyu_plant_min) AS nounyu_plant, "
        strSql = strSql & "MIN(dbo.View_Frm_schedule_kensaku_hacchu.pk_koubai_hacchu_no_min) AS pk_koubai_hacchu_no, "
        strSql = strSql & "MIN(dbo.View_Frm_schedule_kensaku_hacchu.zaiko_tenso_no_min) AS zaiko_tenso_no, "
        If Me![chk_pickup_kbn] = 1 Then
            strSql = strSql & "MIN(dbo.View_Frm_schedule_kensaku_hacchu.syukka_denpyo_no_min) AS syukka_denpyo_no, "
        Else
            strSql = strSql & "MIN(tbl_t_chotatsu_syukka_denpyo.pk_syukka_denpyo_no) AS syukka_denpyo_no, "
        End If
        strSql = strSql & "MIN(dbo.View_Frm_schedule_kensaku_hacchu.R3_soshin_ymd_min) AS R3_soshin_ymd "
    End If
    strSql = strSql & "FROM " & strViewName & " LEFT JOIN dbo.View_Frm_schedule_kensaku_hacchu ON "
    strSql = strSql & strViewName & ".pk_invoice_no_main = dbo.View_Frm_schedule_kensaku_hacchu.invoice_no_main "
    If Me![chk_pickup_kbn] = 1 Then
        strSql = strSql & "AND " & strViewName & ".syanai_sansyo_no = dbo.View_Frm_schedule_kensaku_hacchu.syanai_sansyo_no_tenso_grp "
    Else
        strSql = strSql & "LEFT JOIN tbl_t_chotatsu_syukka_denpyo ON "
        strSql = strSql & strViewName & ".pk_invoice_no_main = dbo.tbl_t_chotatsu_syukka_denpyo.pk_invoice_no_main "
    End If
    
    strWhere = "WHERE "
    If intTaisyoYyyy <> 0 Then
        strSql = strSql & strWhere & "Year( " & strViewName & ".bl_date) = " & intTaisyoYyyy & " "
        strWhere = "AND "
    End If
    If intTaisyoMm <> 0 Then
        strSql = strSql & strWhere & "month( " & strViewName & ".bl_date) = " & intTaisyoMm & " "
    End If
    If Len(strSelectInvoiceNoMain) > 0 Then
        strSql = strSql & strWhere & strViewName & ".pk_invoice_no_main = '" & strSelectInvoiceNoMain & "' "
        strWhere = "AND "
    End If
    If Me![chk_pickup_kbn] = 1 Then
        If Len(strSelectInvoiceNo) > 0 Then
            strSql = strSql & strWhere & strViewName & ".pk_invoice_no_grp = '" & strSelectInvoiceNo & "' "
            strWhere = "AND "
        End If
    End If
    If Len(strSelectSyanaiSansyoNo) > 0 Then
        strSql = strSql & strWhere & strViewName & ".syanai_sansyo_no_tenso_grp = '" & strSelectSyanaiSansyoNo & "' "
        strWhere = "AND "
    End If
    
    strSql = strSql & "GROUP BY " & strViewName & ".pk_invoice_no_main, " & strViewName & ".bl_date, " & strViewName & ".haiso_yotei_ymd "
    If Me![chk_display_kbn] = 2 Then
        If Me![chk_pickup_kbn] = 1 Then
            strSql = strSql & ", " & strViewName & ".pk_invoice_no_grp "
        End If
    End If
    If Me![chk_pickup_kbn] = 1 Then
        strSql = strSql & ", " & strViewName & ".invoice_no_print "
    End If
    If Me![chk_display_kbn] = 2 Then
        strSql = strSql & ", dbo.View_Frm_schedule_kensaku_hacchu.syanai_sansyo_no_tenso_grp, "
        strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.nounyu_plant_min, "
        strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.pk_koubai_hacchu_no_min, "
        strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.zaiko_tenso_no_min, "
        If Me![chk_pickup_kbn] = 1 Then
            strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.syukka_denpyo_no_min, "
        Else
            strSql = strSql & "tbl_t_chotatsu_syukka_denpyo.pk_syukka_denpyo_no, "
        End If
        strSql = strSql & "dbo.View_Frm_schedule_kensaku_hacchu.R3_soshin_ymd_min "
    End If
    strSql = strSql & "ORDER BY " & strViewName & ".bl_date DESC "
    If Me![chk_display_kbn] = 2 Then
        strSql = strSql & ", dbo.View_Frm_schedule_kensaku_hacchu.syanai_sansyo_no_tenso_grp "
    End If
    
    'SQL文実行
    Debug.Print strSql
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    
    Set rs1 = cmd.Execute(intRcnt)
    '実行結果　⇒　合ったら-1、なかったら0
    intPickupCount = intRcnt
    rs1.Close
    
    '対象レコードがなかったら、０データにしないようにデフォルトのViewを読み込んでおく
    If intPickupCount = 0 Then
        If Me![chk_pickup_kbn] = 1 Then
            strSql = "SELECT View_Frm_schedule_kensaku_default_shanghai.* FROM View_Frm_schedule_kensaku_default_shanghai "
        Else
            strSql = "SELECT View_Frm_schedule_kensaku_default_chotatsu.* FROM View_Frm_schedule_kensaku_default_chotatsu "
        End If
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
        End With
        Set rs1 = cmd.Execute(intRcnt)
        rs1.Close
    End If
    
    'レコードセットの再オープン
    rs1.CursorLocation = adUseClient
    rs1.Open strSql
    
    '画面フィールドのコントロールソースを外す
    Me.RecordSource = ""
    Me![pk_invoice_no_main].ControlSource = ""
    Me![BL_DATE].ControlSource = ""
    Me![pk_invoice_no].ControlSource = ""
    Me![haiso_yotei_ymd].ControlSource = ""
    Me![syanai_sansyo_no].ControlSource = ""
    Me![nounyu_plant].ControlSource = ""
    Me![pk_koubai_hacchu_no].ControlSource = ""
    Me![zaiko_tenso_no].ControlSource = ""
    Me![R3_soshin_ymd].ControlSource = ""
    Me![syukka_denpyo_no].ControlSource = ""
    
    'フォームオープン
    DoCmd.OpenForm "F_4_スケジュール検索_MAIN"
    
    'フォームのレコードセットを読み込んだレコードセットにする
    Set Forms![F_4_スケジュール検索_MAIN].Recordset = rs1
    
    'レコードセットのフィールドを、フォームのコントロールソースにセット、画面表示
    Me![pk_invoice_no_main].ControlSource = "pk_invoice_no_main"
    Me![BL_DATE].ControlSource = "BL_DATE"
    If intSelectDisplayKbn = 1 Then
        Me![pk_invoice_no].ControlSource = "invoice_no_print"
    Else
        Me![pk_invoice_no].ControlSource = "pk_invoice_no"
    End If
    Me![haiso_yotei_ymd].ControlSource = "haiso_yotei_ymd"
    Me![syanai_sansyo_no].ControlSource = "syanai_sansyo_no_tenso"
    Me![nounyu_plant].ControlSource = "nounyu_plant"
    Me![pk_koubai_hacchu_no].ControlSource = "pk_koubai_hacchu_no"
    Me![zaiko_tenso_no].ControlSource = "zaiko_tenso_no"
    Me![R3_soshin_ymd].ControlSource = "R3_soshin_ymd"
    Me![syukka_denpyo_no].ControlSource = "syukka_denpyo_no"

    'フォームのRecordSourceをクリアしているので、変数に取っておいた初期情報を再セットする
    Me![taisyo_yyyymm] = dblTaisyoYyyyMm
    Me![taisyo_yyyy] = intTaisyoYyyy
    Me![taisyo_mm] = intTaisyoMm
    If Len(strSelectInvoiceNoMain) > 0 Then
        Me![select_pk_invoice_no_main] = strSelectInvoiceNoMain
    End If
    If Len(strSelectInvoiceNo) > 0 Then
        Me![select_pk_invoice_no] = strSelectInvoiceNo
    End If
    If Len(strSelectSyanaiSansyoNo) > 0 Then
        Me![select_syanai_sansyo_no] = strSelectSyanaiSansyoNo
    End If
    Me![chk_pickup_kbn] = intSelectPickupKbn
    If Len(intSelectDisplayKbn) > 0 Then
        Me![chk_display_kbn] = intSelectDisplayKbn
    End If
    
    Me![login_code] = strLoginCode
    Me![login_name] = strLoginName
    Me![bumon_code] = strBumonCode
    Me![return_form_name] = strReturnFormName
    
    'メッセージ
    If intPickupCount = 0 Then
        MsgBox ("対象データはありませんでした。")
    Else
        MsgBox ("確認してください。")
    End If
    
    Set cmd = Nothing
    
    Exit Sub
    
Err_btn_pickup:
    
    '無理なSQL分や、何回も同じ項目を並べるとエラーになっちゃうので、その時はご連絡ください。
    'いろいろ初期表示、再接続、RecordSourceの変更
    If Len(intSelectPickupKbn) <= 0 Or intSelectPickupKbn = 0 Then
        intSelectPickupKbn = 1
    End If
    If Len(intTaisyoYyyy) <= 0 Or intTaisyoYyyy = 0 Then
        Me![taisyo_yyyy] = Year(Date)
    Else
        Me![taisyo_yyyy] = intTaisyoYyyy
    End If
    If Len(taisyo_mm) <= 0 Or taisyo_mm = 0 Then
        Me![taisyo_yyyy] = Month(Date)
    Else
        Me![taisyo_yyyy] = intTaisyoYyyy
    End If
    Me![taisyo_yyyymm] = intTaisyoYyyy * 100 + intTaisyoMm

    If Len(strSelectInvoiceNoMain) > 0 Then
        Me![select_pk_invoice_no_main] = strSelectInvoiceNoMain
    End If
    If Len(strSelectInvoiceNo) > 0 Then
        Me![select_pk_invoice_no] = strSelectInvoiceNo
    End If
    If Len(strSelectSyanaiSansyoNo) > 0 Then
        Me![select_syanai_sansyo_no] = strSelectSyanaiSansyoNo
    End If
    Me![chk_pickup_kbn] = intSelectPickupKbn
    If Len(intSelectDisplayKbn) > 0 Then
        Me![chk_display_kbn] = intSelectDisplayKbn
    End If
    
    Me![login_code] = strLoginCode
    Me![login_name] = strLoginName
    Me![bumon_code] = strBumonCode
    Me![return_form_name] = strReturnFormName
    Call btn_setsuzoku_Click
    
    If intSelectPickupKbn = 1 Then
        Me.RecordSource = "View_Frm_schedule_kensaku_default_shanghai"
    Else
        Me.RecordSource = "View_Frm_schedule_kensaku_default_chotatsu"
    End If
    
    MsgBox ("エラーが発生しました。" & vbLf & "抽出条件をIT室にご連絡ください。")
    
    Set cmd = Nothing
        
End Sub

Private Sub btn_sentaku_Click()
    
    '進行状況確認処理へ
    strFormName = "F_4_輸入全般_進行状況確認_MAIN"
    intMenuNo = 403
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_4_スケジュール検索_MAIN")
    intMenuNo = 402
    intMenuNoSeq = 0

End Sub

Private Sub btn_shanghai_chotatsu_Click()

    If Me![btn_shanghai_chotatsu].Caption = "上海" Then
        Me![btn_shanghai_chotatsu].Caption = "調達"
        intShanghaiChotatsu = 2
        Me![chk_pickup_kbn] = 2
    Else
        Me![btn_shanghai_chotatsu].Caption = "上海"
        intShanghaiChotatsu = 1
        Me![chk_pickup_kbn] = 1
    End If
    
    Me![select_pk_invoice_no_main] = Null
    
    strFormName = "F_4_スケジュール検索_MAIN"
    Call next_form_open
    
    Call chk_pickup_kbn_Click

End Sub

Private Sub chk_display_kbn_Click()

    If IsNull(Me![taisyo_yyyy]) Or IsNull(Me![taisyo_mm]) Or Len(Me![taisyo_yyyy]) <= 0 Or Len(Me![taisyo_mm]) <= 0 Then
        Exit Sub
    End If

    If Not IsNull(Me![taisyo_yyyy]) And Not IsNull(Me![taisyo_mm]) Then
        Me![taisyo_yyyymm] = Me![taisyo_yyyy] * 100 + Me![taisyo_mm]
    End If
    
End Sub

Private Sub chk_pickup_kbn_Click()
    
    Dim strRowSource1 As String
    Dim strRowSource2 As String
    Dim strRowSource3 As String
    
    Call chk_display_kbn_Click

    If Me![chk_pickup_kbn] = 1 Then
        Me![chk_invoice].Visible = True
        Me![select_pk_invoice_no].Visible = True
        strRowSource1 = "SELECT View__invoice_no_main_header.pk_invoice_no_main FROM View__invoice_no_main_header "
        If Not IsNull(Me![taisyo_yyyymm]) And Me![taisyo_yyyymm] <> 0 Then
            strRowSource1 = strRowSource1 & "WHERE YEAR(bl_date) * 100 + month(bl_date) = " & Me![taisyo_yyyymm] & " "
        End If
        strRowSource2 = "SELECT View__invoice_no.pk_invoice_no FROM View__invoice_no WHERE View__invoice_no.pk_invoice_no = "
        strRowSource2 = strRowSource2 & "'" & Me![pk_invoice_no_main] & "' ORDER BY View__invoice_no.pk_invoice_no DESC; "
        strRowSource3 = "SELECT tbl_t_shanghai_invoice_meisai.syanai_sansyo_no FROM tbl_t_shanghai_invoice_meisai WHERE "
        strRowSource3 = strRowSource3 & "tbl_t_shanghai_invoice_meisai.pk_invoice_no_main = '" & Me![select_pk_invoice_no_main] & "' "
        strRowSource3 = strRowSource3 & "And tbl_t_shanghai_invoice_meisai.pk_invoice_no Like '%" & Me![select_pk_invoice_no] & "%' "
        strRowSource3 = strRowSource3 & "GROUP BY tbl_t_shanghai_invoice_meisai.syanai_sansyo_no "
    
    Else
        Me![chk_invoice].Visible = False
        Me![select_pk_invoice_no].Visible = False
        strRowSource1 = "SELECT View__invoice_no_main_chotatsu.pk_invoice_no_main FROM View__invoice_no_main_chotatsu "
        If Not IsNull(Me![taisyo_yyyymm]) And Me![taisyo_yyyymm] <> 0 Then
            strRowSource1 = strRowSource1 & "WHERE YEAR(bl_date) * 100 + month(bl_date) = " & Me![taisyo_yyyymm] & " "
        End If
        strRowSource2 = "View__invoice_no_main_chotatsu"
        strRowSource3 = "SELECT tbl_t_R3_koubai_hacchu_meisai.syanai_sansyo_no FROM tbl_t_R3_koubai_hacchu_meisai WHERE "
        strRowSource3 = strRowSource3 & "tbl_t_R3_koubai_hacchu_meisai.INVOICE_NO_MAIN = '" & Me![select_pk_invoice_no_main] & "' "
        strRowSource3 = strRowSource3 & "GROUP BY tbl_t_R3_koubai_hacchu_meisai.syanai_sansyo_no "
        
    End If
    
    Me![select_pk_invoice_no_main].RowSource = strRowSource1
    Me![select_pk_invoice_no].RowSource = strRowSource2
    Me![select_syanai_sansyo_no].RowSource = strRowSource3
    Me![select_pk_invoice_no_main].Requery
    Me![select_pk_invoice_no].Requery
    Me![select_syanai_sansyo_no].Requery
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 402
    intMenuNoSeq = 0
    
    Me![taisyo_yyyy] = Year(Date)
    Me![taisyo_mm] = Month(Date)
    
    Call chk_pickup_kbn_Click

End Sub

Private Sub select_pk_invoice_no_AfterUpdate()
    
    Call chk_pickup_kbn_Click

End Sub

Private Sub select_pk_invoice_no_main_AfterUpdate()
    
    Call chk_pickup_kbn_Click

End Sub

Private Sub select_syanai_sansyo_no_AfterUpdate()
    
    Call chk_display_kbn_Click

End Sub

Private Sub taisyo_mm_AfterUpdate()

    Call chk_pickup_kbn_Click
    
End Sub

Private Sub taisyo_mm_LostFocus()

    Call chk_display_kbn_Click

End Sub

Private Sub taisyo_yyyy_AfterUpdate()

    Call chk_pickup_kbn_Click
    
End Sub

Private Sub taisyo_yyyy_LostFocus()

    Call chk_display_kbn_Click

End Sub

Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    strInvoiceNo = ""
    If strFormName = "F_4_輸入全般_進行状況確認_MAIN" Then
        strInvoiceNo = Me![pk_invoice_no_main]
    Else
        If Not IsNull(Me![select_pk_invoice_no_main]) Then
            strInvoiceNo = Me![select_pk_invoice_no_main]
        Else
            strInvoiceNo = "NULL"
        End If
    End If
    If strFormName = "F_4_スケジュール検索_MAIN" Then
        strOpenMode = "O"
    Else
        strOpenMode = "OC"
    End If
    
    Call form_open_close("F_4_スケジュール検索_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub
