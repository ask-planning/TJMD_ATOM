Attribute VB_Name = "Form_F_1_上海輸入_INVOICE_MAIN"
Attribute VB_Base = "0{43DF8BB2-42CB-4F58-B061-B0AA0B0DBC38}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit
    Dim intHeaderFlg As Integer

Private Sub btn_container_settei_Click()

    'ログインのチェック
    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INOVICE_NOを入力してください。")
        Exit Sub
    End If
    
    If Me![btn_container_settei].Caption = "コンテナ設定" Then
        
        Set cmd.ActiveConnection = conn
        
        With cmd
            .CommandText = "usp_shanghai_container_no_temp_create"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = Me![pk_invoice_no_main]
            .Execute
        End With
        
        Set cmd = Nothing
        
        Me![btn_container_settei].Caption = "INVOICEに戻る"
        Me![btn_header_create].Caption = "コンテナNO更新"
        Me![F_1_上海輸入_INVOICE_SUB].SourceObject = "F_1_上海輸入_INVOICE_SUB_コンテナ設定"
        
    Else
        Me![btn_container_settei].Caption = "コンテナ設定"
        Me![btn_header_create].Caption = "HEADER作成"
        Me![F_1_上海輸入_INVOICE_SUB].SourceObject = "F_1_上海輸入_INVOICE_SUB"
    End If
    
    
    Me![F_1_上海輸入_INVOICE_SUB].Requery
    
End Sub

Private Sub btn_header_create_Click()

    'ログインのチェック
    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If

    Dim intInvoiceCount As Integer
    Dim intMeisaiCheck As Integer
    
    Dim strKeySansyoNo As String
    Dim strKeyPlant As String
    Dim strKeyHokanBasyo As String
    Dim intSeqNo As Integer
    Dim intMeisaiNo As Integer
    
    Dim intOkikae As Integer
    
    Set cmd.ActiveConnection = conn
    
     '登録済みか
    intInvoiceCount = 0
    With cmd
        .CommandText = "usp_shanghai_invoice_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(4)) And cmd.Parameters(4) > 1 Then
        MsgBox ("同じINVOICE_NOで、BL_DATE、出荷方法が異なるPACKINGがあります。" & vbLf & "PACKINGデータを修正してください。")
        GoTo Exit_btn_header_click
    End If
    
    If Not IsNull(cmd.Parameters(2)) And cmd.Parameters(2) <> 0 Then
        intInvoiceCount = cmd.Parameters(2)
    
        If Me![btn_header_create].Caption = "HEADER作成" Then
            ReturnValue = MsgBox("ここで作成されるHEADERは、B/L DATEと出荷方法のみです。" & vbLf & "既存の入力済み情報は消えてしまいます。" & vbLf & "作成しますか？", vbYesNo)
            If ReturnValue = vbNo Then
                GoTo Exit_btn_header_click
            End If
        End If
        
    End If
   
    If Me![btn_header_create].Caption = "HEADER作成" Then
        
        'PACKINGに有る情報だけで、ヘッダ作成（INOVICE_NO、出荷方法、BL DATE)
        With cmd
            .CommandText = "usp_shanghai_invoice_header_create"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = Me![pk_invoice_no_main]
            .Execute
        End With
        
        Call pk_invoice_no_main_AfterUpdate
    
    Else
    
        'INVOICE NO別コンテナNO更新
        With cmd
            .CommandText = "usp_shanghai_container_no_update"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = 1
            .Execute
        End With
        
    End If
    
    
    MsgBox ("登録完了しました。")
    
Exit_btn_header_click:

    Set cmd = Nothing
    
End Sub

Private Sub btn_hacchu_check_Click()

    '購買発注確認へ
    strFormName = "F_1_R3購買発注_MAIN"
    intMenuNo = 105
    intMenuNoSeq = 0
    Call next_form_open

End Sub

Private Sub btn_inv_torikomi_Click()
    
    strFormName = "F_1_上海輸入_INVOICE取込"
    intMenuNo = 101
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

Private Sub btn_meisai_koushin_Click()

    'ログインのチェック
    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    Dim intKoushinFlg As Integer
    
    Set cmd.ActiveConnection = conn
    
    '登録済みか
    With cmd
        .CommandText = "usp_shanghai_invoice_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If IsNull(cmd.Parameters(2)) Or cmd.Parameters(2) = 0 Then
        MsgBox ("新規登録の場合、ヘッダ入力処理でヘッダを作成後、再度実行してください｡ ")
        GoTo Exit_btn_meisei_koushin_click
    End If
    
    ReturnValue = MsgBox("置き換えますか？" & vbLf & "置き換え／はい　キャンセル／いいえ", vbYesNo)
    If ReturnValue = vbNo Then
        GoTo Exit_btn_meisei_koushin_click
    End If
    
    strSyanaiSansyoNo = "NULL"
    strInvoiceNo = "NULL"
    If Not IsNull(Me![syanai_sansyo_no]) Then
        intKoushinFlg = 3
        strSyanaiSansyoNo = Me![syanai_sansyo_no]
    Else
        If Not IsNull(Me![pk_invoice_no]) Then
            intKoushinFlg = 2
            strInvoiceNo = Me![pk_invoice_no]
        Else
            intKoushinFlg = 1
        End If
    End If
    
    With cmd
        .CommandText = "usp_shanghai_invoice_meisai_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = intKoushinFlg
        .Parameters(3) = Me![pk_invoice_no_main]
        .Parameters(4) = strInvoiceNo
        .Parameters(5) = strSyanaiSansyoNo
        .Execute
    End With
    
    MsgBox ("更新完了しました。")
    
Exit_btn_meisei_koushin_click:

    Set cmd = Nothing

End Sub

Private Sub btn_shanghai_chotatsu_Click()

    'If Me![btn_shanghai_chotatsu].Caption = "上海" Then
    '    Me![btn_shanghai_chotatsu].Caption = "貿易"
    '    intShanghaiChotatsu = 3
    'Else
    '    Me![btn_shanghai_chotatsu].Caption = "上海"
    '    intShanghaiChotatsu = 1
    'End If
    
    'intMenuNo = 102
    'Me![pk_invoice_no_main] = Null
    
    'strFormName = "F_1_上海輸入_INVOICE_MAIN"
    'Call next_form_open

End Sub

Private Sub btn_syanai_sansyo_no_pickup_Click()

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE NOを入力してください。")
        Exit Sub
    End If
    
    '社内参照番号抽出
    intMenuNo = 102
    'intMenuNoSeq = 1
    If Me![shanghai_chotatsu] = 3 Then
        intMenuNoSeq = 2
    Else
        intMenuNoSeq = 1
    End If
    If Not IsNull(Me![pk_invoice_no_main]) Then
        strInputCode = Me![pk_invoice_no_main]
    Else
        strInputCode = ""
    End If
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_sap_amount_Click()

    '金額チェックへ
    strFormName = "F_1_SAP_AMOUNT_MAIN"
    intMenuNo = 107
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_1_上海輸入_INVOICE_MAIN")
    intMenuNo = 102
    intMenuNoSeq = 0

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
    intMenuNo = 102
    intMenuNoSeq = 0
    
End Sub

Private Sub frm_hyoji_kbn_Click()

    'ログインのチェック
    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If

    Select Case frm_hyoji_kbn
        Case 1
            Me![F_1_上海輸入_INVOICE_SUB].SourceObject = "F_1_上海輸入_INVOICE_SUB"
            Me![F_1_上海輸入_INVOICE_SUB].Requery
        Case 2
            Me![F_1_上海輸入_INVOICE_SUB].SourceObject = "F_1_上海輸入_PACKING_SUB"
            Me![F_1_上海輸入_INVOICE_SUB].Requery
    End Select
    
End Sub

Private Sub pk_invoice_no_AfterUpdate()

    'ログインのチェック
    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Call sub_form_display
    
End Sub

Private Sub pk_invoice_no_main_AfterUpdate()

    'ログインのチェック
    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    Me![syukka_hoho].BackColor = "8454143"
    Me![BL_DATE].BackColor = "8454143"
    Me![syukka_hoho].Requery
    Me![BL_DATE].Requery
    
    Me![vessel] = Null
    Me![total_amount] = Null
    Me![carton_qty] = Null
    Me![syukka_hoho] = Null
    Me![BL_DATE] = Null

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    intHeaderFlg = 0
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
    
    'ヘッダ情報抽出
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
    If intHeaderFlg = 0 Then
        If Not IsNull(cmd.Parameters(2)) Then
            Me![BL_DATE] = cmd.Parameters(2)
        End If
        If Not IsNull(cmd.Parameters(3)) Then
            Me![syukka_hoho] = cmd.Parameters(3)
        End If
    End If
    
    Set cmd = Nothing
    
    If intHeaderFlg = 0 Then
        Me![syukka_hoho].BackColor = "12632256"
        Me![BL_DATE].BackColor = "12632256"
        Me![syukka_hoho].Requery
        Me![BL_DATE].Requery
    End If
    
    Call sub_form_display
    
End Sub
Private Sub btn_close_Click()

    'メインメニューへ戻る
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    
    Call next_form_open
    
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
    
    intMenuNo = 102
    Me![pk_invoice_no_main] = Null
    
    strFormName = "F_1_上海輸入_INVOICE_MAIN"
    Call next_form_open

End Sub

Private Sub syanai_sansyo_no_AfterUpdate()

    'ログインのチェック
    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Call sub_form_display
    
End Sub

Private Sub sub_form_display()

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("元INVOICE_NOを入力してください。")
        Exit Sub
    End If
    
    If Not IsNull(Me![pk_invoice_no]) Then
        strInvoiceNo = "%" & Me![pk_invoice_no] & "%"
    Else
        strInvoiceNo = "%%"
    End If
    
    If Not IsNull(Me![syanai_sansyo_no]) Then
        strSyanaiSansyoNo = "%" & Me![syanai_sansyo_no] & "%"
    Else
        strSyanaiSansyoNo = "%%"
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_shanghai_invoice_meisai_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_invoice_no_main]
        .Parameters(3) = strInvoiceNo
        .Parameters(4) = strSyanaiSansyoNo
        .Execute
    End With
    
    Set cmd = Nothing
    
    'If Not IsNull(Me![syanai_sansyo_no]) Then
    '    Me![btn_meisai_koushin].Enabled = False
    'Else
    '    Me![btn_meisai_koushin].Enabled = True
    'End If
    Me![F_1_上海輸入_INVOICE_SUB].Requery
    
    strSql = "SELECT pk_invoice_no FROM dbo.tbl_t_shanghai_invoice_meisai WHERE pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
    If Not IsNull(Me![syanai_sansyo_no]) Then
        strSql = strSql & "AND syanai_sansyo_no = '" & Me![syanai_sansyo_no] & "' "
    End If
    strSql = strSql & "GROUP BY pk_invoice_no ORDER BY pk_invoice_no DESC "
    Me![pk_invoice_no].RowSource = strSql
    
    
    strSql = "SELECT syanai_sansyo_no FROM dbo.tbl_t_shanghai_invoice_meisai WHERE pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
    If Not IsNull(Me![pk_invoice_no]) Then
        strSql = strSql & "AND pk_invoice_no = '" & Me![pk_invoice_no] & "' "
    End If
    strSql = strSql & "GROUP BY syanai_sansyo_no ORDER BY syanai_sansyo_no DESC "
    Me![syanai_sansyo_no].RowSource = strSql
    
    Me![pk_invoice_no].Requery
    Me![syanai_sansyo_no].Requery
    
    If Not IsNull(Me![pk_invoice_no]) Or Not IsNull(Me![syanai_sansyo_no]) Then
        Me![btn_meisai_koushin].Enabled = False
    Else
        Me![btn_meisai_koushin].Enabled = True
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
    If Not IsNull(Me![pk_invoice_no_main]) Then
        strInvoiceNo = Me![pk_invoice_no_main]
    Else
        strInvoiceNo = "NULL"
    End If
    strOpenMode = "OC"
    
    Call form_open_close("F_1_上海輸入_INVOICE_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

