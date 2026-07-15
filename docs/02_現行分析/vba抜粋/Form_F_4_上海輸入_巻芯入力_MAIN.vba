Attribute VB_Name = "Form_F_4_上海輸入_巻芯入力_MAIN"
Attribute VB_Base = "0{47B44ACE-4CC8-4ACA-81A2-A8B523CD70A7}"
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

Private Sub btn_excel_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE NOを入力してください。")
        Exit Sub
    End If
    
    intMenuNo = 401
    intMenuNoSeq = 1
    strInputCode = Me![pk_invoice_no_main]
    
    strFormName = "F_9_メール送信_MAIN"
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

    DoCmd.OpenReport "R_上海輸入_配送手配_巻芯", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' "
    

    '進捗状況更新
    Call schedule_update(7, Date, Me![pk_invoice_no_main], 1)
    
End Sub


Private Sub btn_nyuka_meisai_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを入力してください。")
        Exit Sub
    End If

    '進捗状況更新
    Call schedule_update(10, Date, Me![pk_invoice_no_main], 1)

    DoCmd.OpenReport "R_上海タジマ入荷明細_巻芯", acViewPreview, , "pk_login_code = '" & Me![login_code] & "' "
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_4_上海輸入_巻芯入力_MAIN")
    intMenuNo = 401
    intMenuNoSeq = 0

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
    
    Set cmd.ActiveConnection = conn
    
    Dim strPara() As String
    Dim intPara() As Integer
    
    t = 7
    
    ReDim strPara(t) As String
    ReDim intPara(t) As Integer
    
    strPara(0) = "pk_invoice_no_main":      intPara(0) = 0
    strPara(1) = "BL_DATE":                 intPara(1) = 0
    strPara(2) = "vessel":                  intPara(2) = 0
    strPara(3) = "syukka_hoho":             intPara(3) = 1
    strPara(4) = "carton_qty":              intPara(4) = 1
    strPara(5) = "otsunaka_no":             intPara(5) = 1
    strPara(6) = "update_ymd":              intPara(6) = 0
    strPara(7) = "update_login":            intPara(7) = 0
    
    If intHeaderFlg = 0 Then
    
        strSql = "INSERT INTO tbl_t_shanghai_invoice_header ("
        For n = 0 To t
            If n <> 0 Then
                strSql = strSql & ","
            End If
            strSql = strSql & strPara(n)
        Next
        strSql = strSql & ") VALUES ("
        For n = 0 To t - 3
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
        strSql = strSql & ", 5, '" & Now() & "', '" & Me![login_code] & "' )"
                
    Else
    
        strSql = "UPDATE tbl_t_shanghai_invoice_header SET "
        For n = 0 To t - 3
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
    
    Debug.Print strSql
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    'ヘッダーサブ、明細作成
    With cmd
        .CommandText = "usp_shanghai_invoice_makishin_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_invoice_no_main]
        .Parameters(3) = Me![haiso_yotei_ymd]
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
    
    Set cmd = Nothing

    '進捗状況更新
    Call schedule_update(3, Date, Me![pk_invoice_no_main], 1)
    
    MsgBox ("登録完了しました。")
    
    Call pk_invoice_no_main_AfterUpdate

End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 401
    intMenuNoSeq = 1
    
End Sub

Private Sub pk_invoice_no_main_AfterUpdate()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_invoice_no_main]) Then
        Exit Sub
    End If
    
    Dim strEditInvoiceNo As String
    Dim intInvoiceCount As Integer
    
    Set cmd.ActiveConnection = conn
    
    Me![carton_qty].BackColor = "8454143"
    Me![total_suryo].BackColor = "8454143"
    Me![carton_qty].Requery
    Me![total_suryo].Requery
    
    intHeaderFlg = 0
    strSql = "SELECT tbl_t_shanghai_invoice_header.pk_invoice_no_main, tbl_t_shanghai_invoice_header.BL_DATE, "
    strSql = strSql & "tbl_t_shanghai_invoice_header.vessel, tbl_t_shanghai_invoice_header.syukka_hoho, "
    strSql = strSql & "tbl_t_shanghai_invoice_header_sub.haiso_yotei_ymd FROM tbl_t_shanghai_invoice_header "
    strSql = strSql & "INNER JOIN tbl_t_shanghai_invoice_header_sub ON tbl_t_shanghai_invoice_header.pk_invoice_no_main = "
    strSql = strSql & "tbl_t_shanghai_invoice_header_sub.pk_invoice_no_main "
    strSql = strSql & "INNER JOIN tbl_t_shanghai_invoice_meisai_makishin ON tbl_t_shanghai_invoice_header.pk_invoice_no_main = "
    strSql = strSql & "tbl_t_shanghai_invoice_meisai_makishin.pk_invoice_no_main GROUP BY tbl_t_shanghai_invoice_header.pk_invoice_no_main, "
    strSql = strSql & "tbl_t_shanghai_invoice_header.BL_DATE, tbl_t_shanghai_invoice_header.vessel, tbl_t_shanghai_invoice_header."
    strSql = strSql & "syukka_hoho, tbl_t_shanghai_invoice_header_sub.haiso_yotei_ymd "
    strSql = strSql & "HAVING tbl_t_shanghai_invoice_header.pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
    Debug.Print strSql
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
    Else
        For Each ct In Me
            If ct.Name <> "pk_invoice_no_main" And Left(ct.Name, 5) <> "login" And ct.Name <> "bumon_code" Then
                Me(ct.Name) = Null
            End If
        Next
        Me![syukka_hoho] = 0
    End If
    On Error GoTo 0
    rs1.Close
    
    strSql = "SELECT pk_invoice_no_main, Count([carton_no]) AS [carton_qty], SUM(invoice_suryo) AS [total_suryo] "
    strSql = strSql & "FROM View_shanghai_makishin_meisai_suryo "
    strSql = strSql & "WHERE View_shanghai_makishin_meisai_suryo.pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
    strSql = strSql & "GROUP BY View_shanghai_makishin_meisai_suryo.pk_invoice_no_main "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
        rs1.MoveFirst
        Me![total_suryo] = rs1![total_suryo].Value
        Me![carton_qty] = rs1![carton_qty].Value
    Else
        Me![total_suryo] = Null
        Me![carton_qty] = Null
    End If
    rs1.Close
    
    If intHeaderFlg = 0 Then
        Me![carton_qty].BackColor = "12632256"
        Me![total_suryo].BackColor = "12632256"
        Me![carton_qty].Requery
        Me![total_suryo].Requery
    End If
    
    With cmd
        .CommandText = "usp_shanghai_invoice_makishin_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    Me![F_4_上海輸入_巻芯入力_SUB].Requery
    
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
    strOpenMode = "OC"
    
    Call form_open_close("F_4_上海輸入_巻芯入力_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub
