Attribute VB_Name = "Form_F_1_R3在庫転送_MAIN"
Attribute VB_Base = "0{B33E0EF3-D670-4580-9F7C-0FF43CC85227}"
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
    
    intShinkiFlg = 1
    
    Call zaiko_tenso_temp_create
    
End Sub

Private Sub zaiko_tenso_temp_create()

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If

    Dim intInvoiceCount As Integer
    Dim intMeisaiCheck As Integer
    Dim intSyukkaHoho As Integer
    
    Set cmd.ActiveConnection = conn
    
    Me![BL_DATE] = Null
    Me![carton_qty] = Null
    
    If Not IsNull(Me![INVOICE_NO]) Then
        strInvoiceNo = Me![INVOICE_NO]
    Else
        strInvoiceNo = "%%"
    End If
    If Not IsNull(Me![syanai_sansyo_no]) Then
        strSyanaiSansyoNo = Me![syanai_sansyo_no]
    Else
        strSyanaiSansyoNo = "%%"
    End If
    If Not IsNull(Me![nounyu_plant]) Then
        strNounyuPlant = Me![nounyu_plant]
    Else
        strNounyuPlant = "%%"
    End If
    
    intSyukkaHoho = 0
    If Me![shanghai_chotatsu] <> 2 Then
        With cmd
            .CommandText = "usp_shanghai_invoice_header_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![shanghai_chotatsu]
            .Parameters(2) = Me![INVOICE_NO_MAIN]
            .Execute
        End With
        If Not IsNull(cmd.Parameters(4)) Then
            intSyukkaHoho = cmd.Parameters(4)
        End If
    End If
    
    'parameters(2)=1⇒shanghai、parameters(5)=1⇒新規
    intInvoiceCount = 0
    With cmd
        .CommandText = "usp_R3_zaiko_tenso_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Parameters(3) = intShinkiFlg
        .Parameters(4) = intSyukkaHoho
        .Parameters(5) = Me![INVOICE_NO_MAIN]
        .Parameters(6) = strInvoiceNo
        .Parameters(7) = strSyanaiSansyoNo
        .Parameters(8) = strNounyuPlant
        .Execute
    End With
    
    If IsNull(cmd.Parameters(9)) Or cmd.Parameters(9) = 0 Then
        If intShinkiFlg = 1 Then
            MsgBox ("未送信の在庫転送オーダーはありませんでした。")
        Else
            MsgBox ("送信済みの在庫転送オーダーはありませんでした。")
        End If
    Else
        intInvoiceCount = cmd.Parameters(9)
        If intShinkiFlg = 1 Then
            MsgBox ("内容を確認して、送信ボタンをクリックしてください。")
        Else
            MsgBox ("再作成する品目の【送信】にチェックを入れ、送信ボタンをクリックしてください。")
        End If
    End If
    
    
Exit_btn_create_Click:
    
    Me![F_1_R3在庫転送_SUB].Requery

    Set cmd = Nothing
    
    On Error Resume Next
    Call syori_log_create(intMenuNo, intMenuNoSeq, Me![INVOICE_NO], "転送作成　新規=" & intShinkiFlg & " 件数=" & intInvoiceCount)
    On Error GoTo 0
    
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

Private Sub btn_invoice_meisai_Click()

    'INVOICE明細入力へ
    strFormName = "F_1_上海輸入_INVOICE_MAIN"
    intMenuNo = 102
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_koubai_hacchu_Click()

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

Private Sub btn_shanghai_chotatsu_Click()

    'If Me![btn_shanghai_chotatsu].Caption = "上海" Then
    '    Me![btn_shanghai_chotatsu].Caption = "調達"
    '    intShanghaiChotatsu = 2
    'Else
    '    Me![btn_shanghai_chotatsu].Caption = "上海"
    '    intShanghaiChotatsu = 1
    'End If
    
    'Me![INVOICE_NO_MAIN] = Null
    
    'strFormName = "F_1_R3在庫転送_MAIN"
    'Call next_form_open

End Sub

Private Sub btn_shiyo_henko_check_Click()

    DoCmd.OpenView "View_rogistic_niji_okurihin_check", acViewNormal, acReadOnly
    
End Sub

Private Sub btn_syanai_sansyo_no_pickup_Click()
    
    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE NOを入力してください。")
        Exit Sub
    End If
    
    '社内参照番号抽出
    If Me![shanghai_chotatsu] = 2 Then
        intMenuNo = 204
    Else
        intMenuNo = 106
    End If
    intMenuNoSeq = 1
    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        strInputCode = Me![INVOICE_NO_MAIN]
    Else
        strInputCode = ""
    End If
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_recreate_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    intShinkiFlg = 2
    
    Call zaiko_tenso_temp_create
    
    
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

    Call Login_ReInput("F_1_R3在庫転送_MAIN")
    If intShanghaiChotatsu = 2 Then
        intMenuNo = 204
    Else
        intMenuNo = 106
    End If
    intMenuNoSeq = 0

End Sub

Private Sub btn_soshin_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![INVOICE_NO_MAIN]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    Dim intInvoiceCount As Integer
    
    ReturnValue = MsgBox("R3サーバーにテキストを送信します。" & vbLf & "はい／送信　いいえ／キャンセル", 4)
    If ReturnValue = vbNo Then
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_R3_zaiko_tenso_temp_count"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    If IsNull(cmd.Parameters(2)) Or cmd.Parameters(2) = 0 Then
        MsgBox ("送信対象データがありません。" & vbLf & "【送信】のチェックを確認してください。")
        GoTo Exit_btn_soshin_Click
    Else
        If IsNull(cmd.Parameters(3)) Or Me![INVOICE_NO_MAIN] <> cmd.Parameters(3) Then
            MsgBox ("選択したINVOICEと抽出済みのINVOICEが異なります。" & vbLf & "再度、作成ボタンをクリックして対象INVOICEを抽出してください。")
            GoTo Exit_btn_soshin_Click
        End If
    End If
    
    strFileName = Me![file_name]
    
    'テキストデータ作成
    With cmd
        .CommandText = "usp_R3_zaiko_tenso_text_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
 
    strSql = "SELECT tbl_t_R3_zaiko_tenso_text_temp.* FROM tbl_t_R3_zaiko_tenso_text_temp "
    strSql = strSql & "ORDER BY syanai_sansyo_no, nounyu_kijitsu_ymd, hinmoku_code "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
    
        intFieldCount = rs1.Fields.Count

        strTextName = Me![file_name]
        'strTextName = "\\172.16.34.212\share\test\atom_tenso.txt"

        'レコードの書き込み
        intFileNo = FreeFile(0)
        Open strTextName For Output As #intFileNo

        rs1.MoveFirst
        Do Until rs1.EOF
            strTextArea = ""
            For n = 0 To intFieldCount - 1
                If Not IsNull(rs1.Fields(n)) Then
                    strTextArea = strTextArea & rs1.Fields(n)
                End If
                If n <> intFieldCount - 1 Then
                    strTextArea = strTextArea & vbTab
                End If
            Next
                        
            Print #intFileNo, strTextArea
            
            rs1.MoveNext
        Loop
                
        Close #intFileNo
                
    End If
    rs1.Close
    
    With cmd
        .CommandText = "usp_R3_zaiko_tenso_soshin_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    MsgBox ("テキストデータを送信しました。" & "Ｒ３でバッチインプットを行ってください。")
    
Exit_btn_soshin_Click:

    Set rs1 = Nothing
    Set cmd = Nothing

End Sub

Private Sub F_1_R3在庫転送_SUB_Enter()

    DoCmd.GoToRecord , , acFirst
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
    'テキストの場所
    Set cmd.ActiveConnection = conn
    
    If intShanghaiChotatsu = 2 Then
        With cmd
            .CommandText = "usp_koumoku_contents_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = "FIL"
            .Parameters(2) = 6
            .Execute
        End With
    Else
        With cmd
            .CommandText = "usp_koumoku_contents_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = "FIL"
            .Parameters(2) = 3
            .Execute
        End With
    End If
    
    Me![file_name] = cmd.Parameters(3) & "\" & cmd.Parameters(4)
    
    Set cmd = Nothing
    
    If intShanghaiChotatsu = 2 Then
        'Me![INVOICE_NO_MAIN].RowSource = "View__invoice_no_main_chotatsu"
        strSql = "SELECT pk_invoice_no_main, bl_date FROM dbo.tbl_t_chotatsu_invoice_header ORDER BY bl_date DESC "
        Me![INVOICE_NO_MAIN].RowSource = strSql
        Me![btn_sap_amount].Enabled = False
        Me![btn_invoice_meisai].Enabled = False
        Me![btn_invoice_header].Caption = "配送手配"
        Me![btn_invoice_header].ForeColor = "10040115"
        intMenuNo = 204
    Else
        'Me![INVOICE_NO_MAIN].RowSource = "View__invoice_no_main"
        strSql = "SELECT tbl_t_shanghai_invoice_meisai.pk_invoice_no_main, tbl_t_shanghai_invoice_header.BL_DATE "
        strSql = strSql & "FROM tbl_t_shanghai_invoice_meisai INNER JOIN tbl_t_shanghai_invoice_header ON "
        strSql = strSql & "tbl_t_shanghai_invoice_meisai.pk_invoice_no_main = tbl_t_shanghai_invoice_header.pk_invoice_no_main "
        strSql = strSql & "WHERE (tbl_t_shanghai_invoice_meisai.boueki_flg = 0) OR (tbl_t_shanghai_invoice_meisai.boueki_flg IS NULL) "
        strSql = strSql & "GROUP BY tbl_t_shanghai_invoice_meisai.pk_invoice_no_main, tbl_t_shanghai_invoice_header.BL_DATE "
        strSql = strSql & "ORDER BY tbl_t_shanghai_invoice_header.bl_date DESC, tbl_t_shanghai_invoice_meisai.pk_invoice_no_main DESC "
        Me![INVOICE_NO_MAIN].RowSource = strSql
        Me![btn_sap_amount].Enabled = True
        Me![btn_invoice_meisai].Enabled = True
        Me![btn_invoice_header].Caption = "INVヘッダ"
        Me![btn_invoice_header].ForeColor = "8388672"
        intMenuNo = 106
    End If
    intMenuNoSeq = 0
    
End Sub

Private Sub INVOICE_NO_AfterUpdate()

    Call header_redisplay
    
End Sub

Private Sub invoice_no_main_AfterUpdate()

    Me![INVOICE_NO] = Null
    Me![syanai_sansyo_no] = Null
    Me![nounyu_plant] = Null
    
    Call header_redisplay
    
End Sub
Private Sub btn_menu_Click()
    
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub nounyu_plant_AfterUpdate()

    Call header_redisplay
    
End Sub

Private Sub shanghai_chotatsu_AfterUpdate()

    Select Case Me![shanghai_chotatsu]
        Case 1
            intShanghaiChotatsu = 1
            intMenuNo = 106
        Case 2
            intShanghaiChotatsu = 2
            intMenuNo = 204
        Case 3
            intShanghaiChotatsu = 3
            intMenuNo = 106
        Case Else
            intShanghaiChotatsu = 1
            intMenuNo = 106
    End Select
    
    Me![INVOICE_NO_MAIN] = Null
    
    strFormName = "F_1_R3在庫転送_MAIN"
    Call next_form_open

End Sub

Private Sub syanai_sansyo_no_AfterUpdate()

    Call header_redisplay
    
End Sub

Private Sub header_redisplay()

    Set cmd.ActiveConnection = conn
    
    Me![BL_DATE] = Null
    Me![carton_qty] = Null
    
    If Not IsNull(Me![INVOICE_NO]) Then
        strInvoiceNo = Me![INVOICE_NO]
    Else
        strInvoiceNo = "%%"
    End If
    If Not IsNull(Me![syanai_sansyo_no]) Then
        strSyanaiSansyoNo = Me![syanai_sansyo_no]
    Else
        strSyanaiSansyoNo = "%%"
    End If
    If Not IsNull(Me![nounyu_plant]) Then
        strNounyuPlant = Me![nounyu_plant]
    Else
        strNounyuPlant = "%%"
    End If
    
    With cmd
        .CommandText = "usp_R3_zaiko_tenso_invoice_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![shanghai_chotatsu]
        .Parameters(3) = Me![INVOICE_NO_MAIN]
        .Parameters(4) = strInvoiceNo
        .Parameters(5) = strSyanaiSansyoNo
        .Parameters(6) = strNounyuPlant
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(7)) Then
        Me![BL_DATE] = cmd.Parameters(7)
    End If
    If Not IsNull(cmd.Parameters(8)) Then
        Me![carton_qty] = cmd.Parameters(8)
    End If
    
    Set cmd = Nothing
    
    If Not IsNull(Me![INVOICE_NO_MAIN]) Then
        strSql = "SELECT pk_invoice_no FROM dbo.tbl_t_shanghai_invoice_meisai WHERE pk_invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
        strSql = strSql & "GROUP BY pk_invoice_no ORDER BY pk_invoice_no DESC"
        Me![INVOICE_NO].RowSource = strSql
        
        If Not IsNull(Me![INVOICE_NO]) Then
            strSql = "SELECT syanai_sansyo_no FROM dbo.tbl_t_shanghai_invoice_meisai WHERE pk_invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
            strSql = strSql & "AND pk_invoice_no = '" & Me![INVOICE_NO] & "' "
            strSql = strSql & "GROUP BY syanai_sansyo_no ORDER BY syanai_sansyo_no "
            Me![syanai_sansyo_no].RowSource = strSql
        Else
            strSql = "SELECT syanai_sansyo_no FROM dbo.tbl_t_shanghai_invoice_meisai WHERE pk_invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
            strSql = strSql & "GROUP BY syanai_sansyo_no ORDER BY syanai_sansyo_no "
            Me![syanai_sansyo_no].RowSource = strSql
        End If
        If Not IsNull(Me![syanai_sansyo_no]) Then
            strSql = "SELECT dbo.View__plant_hokan_basyo.hokan_basyo, dbo.View__plant_hokan_basyo.plant, dbo.View__plant_hokan_basyo.hokan_basyo_text "
            strSql = strSql & "FROM dbo.tbl_t_R3_koubai_hacchu_meisai INNER JOIN dbo.View__plant_hokan_basyo "
            strSql = strSql & "ON dbo.tbl_t_R3_koubai_hacchu_meisai.nounyu_plant = dbo.View__plant_hokan_basyo.hokan_basyo "
            strSql = strSql & "WHERE invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' AND syanai_sansyo_no = '" & Me![syanai_sansyo_no] & "' "
            strSql = strSql & "GROUP BY dbo.View__plant_hokan_basyo.hokan_basyo, dbo.View__plant_hokan_basyo.plant, dbo.View__plant_hokan_basyo.hokan_basyo_text "
            Me![nounyu_plant].RowSource = strSql
        Else
            strSql = "SELECT dbo.View__plant_hokan_basyo.hokan_basyo, dbo.View__plant_hokan_basyo.plant, dbo.View__plant_hokan_basyo.hokan_basyo_text "
            strSql = strSql & "FROM dbo.tbl_t_R3_koubai_hacchu_meisai INNER JOIN dbo.View__plant_hokan_basyo "
            strSql = strSql & "ON dbo.tbl_t_R3_koubai_hacchu_meisai.nounyu_plant = dbo.View__plant_hokan_basyo.hokan_basyo "
            strSql = strSql & "WHERE invoice_no_main = '" & Me![INVOICE_NO_MAIN] & "' "
            strSql = strSql & "GROUP BY dbo.View__plant_hokan_basyo.hokan_basyo, dbo.View__plant_hokan_basyo.plant, dbo.View__plant_hokan_basyo.hokan_basyo_text "
            Me![nounyu_plant].RowSource = strSql
        End If
        
    End If
    
    Me![BL_DATE].Requery
    Me![carton_qty].Requery
    Me![INVOICE_NO].Requery
    Me![syanai_sansyo_no].Requery
    Me![nounyu_plant].Requery
    
    
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
    If strFormName = "F_1_R3在庫転送_MAIN" Then
        strOpenMode = "O"
    Else
        strOpenMode = "OC"
    End If
    
    Call form_open_close("F_1_R3在庫転送_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub
