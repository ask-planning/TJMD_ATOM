Attribute VB_Name = "Form_F_3_立替金送金明細_MAIN"
Attribute VB_Base = "0{8462AF63-65E7-46BE-9D39-D816F72B1DA3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_clear_Click()

    Me![pk_sokin_saki] = Null
    Me![syori_ymd] = Null
    Me![pk_invoice_no_main] = Null
    
    Call sub_form_display
    
End Sub

Private Sub btn_menu_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    'スタートメニューへ戻る
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    strOpenMode = "OC"
    Call form_open_close("F_3_立替金送金明細_MAIN", strFormName, strOpenMode, "NULL")

End Sub

Private Sub btn_print_Click()

    If IsNull(Me![pk_invoice_no_main]) Then
        MsgBox ("INVOICE_NOを選択、または、入力してください。")
        Exit Sub
    End If

    'DoCmd.OpenReport "R_立替金送金明細", acViewPreview, , "pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
    DoCmd.OpenReport "R_振替伝票", acViewPreview, , "pk_invoice_no_main = '" & Me![pk_invoice_no_main] & "' "
    
End Sub

Private Sub btn_setsuzoku_Click()
    
    '再接続
    Call Login_ReInput("F_3_立替金送金明細_MAIN")
    intMenuNo = 302
    intMenuNoSeq = 0

End Sub

Private Sub btn_shanghai_chotatsu_Click()

    'If IsNull(Me![login_code]) Then
    '    Call btn_setsuzoku_Click
    'End If

    'If Me![btn_shanghai_chotatsu].Caption = "上海" Then
    '    Me![btn_shanghai_chotatsu].Caption = "調達"
    '    intShanghaiChotatsu = 2
    'Else
    '    Me![btn_shanghai_chotatsu].Caption = "上海"
    '    intShanghaiChotatsu = 1
    'End If
    
    'Me![pk_invoice_no_main] = Null
    
    'If intShanghaiChotatsu = 2 Then
    '    Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_chotatsu"
    '    strSql = "SELECT syori_ymd FROM tbl_t_tatekaekin_meisai WHERE LEFT(pk_invoice_no_main, 3) = 'SJ-' "
    '    strSql = strSql & "OR LEFT(pk_invoice_no_main, 3) = 'YJ-' GROUP BY syori_ymd ORDER BY syori_ymd DESC "
    '    Me![syori_ymd].RowSource = strSql
    'Else
    '    Me![pk_invoice_no_main].RowSource = "View__invoice_no_main"
    '    strSql = "SELECT syori_ymd FROM tbl_t_tatekaekin_meisai WHERE LEFT(pk_invoice_no_main, 3) <> 'SJ-' "
    '    strSql = strSql & "AND LEFT(pk_invoice_no_main, 3) <> 'YJ-' GROUP BY syori_ymd ORDER BY syori_ymd DESC "
    '    Me![syori_ymd].RowSource = strSql
    'End If
    'Me![pk_invoice_no_main].Requery

End Sub

Private Sub chk_shanghai_chotatsu_Click()

    If IsNull(Me![chk_shanghai_chotatsu]) Or Me![chk_shanghai_chotatsu] = 0 Then
        Exit Sub
    End If
    
    If Me![chk_shanghai_chotatsu] = 2 Then
        Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_chotatsu"
    Else
        Me![pk_invoice_no_main].RowSource = "View__invoice_no_main"
    End If
    Me![pk_invoice_no_main].Requery
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
    intMenuNo = 302
    intMenuNoSeq = 0
   
    If intShanghaiChotatsu = 2 Then
        Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_chotatsu"
    Else
        Me![pk_invoice_no_main].RowSource = "View__invoice_no_main"
    End If
    
    Me![pk_invoice_no_main].Requery
    
    Me![chk_shanghai_chotatsu] = intShanghaiChotatsu
    
End Sub

Private Sub pk_invoice_no_main_AfterUpdate()

    If IsNull(Me![pk_invoice_no_main]) Then
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_tatekaekin_meisai_pickup"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(2)) Then
        Me![pk_sokin_saki] = cmd.Parameters(2)
    End If
    If Not IsNull(cmd.Parameters(3)) Then
        Me![syori_ymd] = cmd.Parameters(3)
    End If
    
    Set cmd = Nothing
    
    Me![F_3_立替金送金明細_SUB].Requery

End Sub

Private Sub shanghai_chotatsu_AfterUpdate()

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
    
    Me![pk_invoice_no_main] = Null
    
    If Me![shanghai_chotatsu] = 2 Then
        Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_chotatsu"
        strSql = "SELECT syori_ymd FROM tbl_t_tatekaekin_meisai WHERE LEFT(pk_invoice_no_main, 3) = 'SJ-' "
        strSql = strSql & "OR LEFT(pk_invoice_no_main, 3) = 'YJ-' GROUP BY syori_ymd ORDER BY syori_ymd DESC "
        Me![syori_ymd].RowSource = strSql
    Else
        Me![pk_invoice_no_main].RowSource = "View__invoice_no_main"
        strSql = "SELECT syori_ymd FROM tbl_t_tatekaekin_meisai WHERE LEFT(pk_invoice_no_main, 3) <> 'SJ-' "
        strSql = strSql & "AND LEFT(pk_invoice_no_main, 3) <> 'YJ-' GROUP BY syori_ymd ORDER BY syori_ymd DESC "
        Me![syori_ymd].RowSource = strSql
    End If
    
    Me![pk_invoice_no_main].Requery


End Sub

Private Sub syori_ymd_AfterUpdate()

    Call sub_form_display

End Sub

Private Sub pk_sokin_saki_AfterUpdate()

    Call sub_form_display

End Sub

Private Sub sub_form_display()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    'DROP DOWNの編集
    strWhere = "WHERE "
    strSql = "SELECT syori_ymd FROM tbl_t_tatekaekin_meisai "
    If Not IsNull(Me![pk_sokin_saki]) Then
        strSql = strSql & strWhere & "pk_sokin_saki = " & Me![pk_sokin_saki] & " "
        strWhere = "AND "
    End If
    If Not IsNull(Me![syori_ymd]) Then
        strSql = strSql & strWhere & "syori_ymd = '" & Me![syori_ymd] & "' "
    End If
    strSql = strSql & "GROUP BY syori_ymd ORDER BY syori_ymd DESC "
    Me![syori_ymd].RowSource = strSql
    Me![syori_ymd].Requery
    
    strWhere = "WHERE "
    strSql = "SELECT pk_invoice_no_main FROM tbl_t_tatekaekin_meisai "
    If Not IsNull(Me![pk_sokin_saki]) Then
        strSql = strSql & strWhere & "pk_sokin_saki = " & Me![pk_sokin_saki] & " "
        strWhere = "AND "
    End If
    If Not IsNull(Me![syori_ymd]) Then
        strSql = strSql & strWhere & "syori_ymd = '" & Me![syori_ymd] & "' "
    End If
    strSql = strSql & "GROUP BY pk_invoice_no_main ORDER BY pk_invoice_no_main DESC "
    Me![pk_invoice_no_main].RowSource = strSql
    Me![pk_invoice_no_main].Requery
    
    Me![F_3_立替金送金明細_SUB].Requery


End Sub
