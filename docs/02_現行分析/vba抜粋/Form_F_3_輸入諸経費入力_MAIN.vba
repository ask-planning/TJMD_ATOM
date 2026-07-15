Attribute VB_Name = "Form_F_3_輸入諸経費入力_MAIN"
Attribute VB_Base = "0{7409898B-B986-47F0-9644-B08A812AB7D3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_close_Click()
    
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    strOpenMode = "OC"
    
    Call next_form_open

End Sub

Private Sub btn_code_master_Click()

    strFormName = "F_M_経費コード登録_MAIN"
    intMenuNo = 907
    intMenuNoSeq = 0
    strOpenMode = "OC"
    
    Call next_form_open
    
End Sub

Private Sub btn_gokei_kakunin_Click()

    Set cmd.ActiveConnection = conn
    
    '仕入先コード抽出
    With cmd
        .CommandText = "usp_import_keihi_header_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(2)) Then
        Me![shiiresaki_code] = cmd.Parameters(2)
        Me![shiiresaki_code].Requery
    End If
    If Not IsNull(cmd.Parameters(3)) Then
        Me![insurance_seikyu_ymd] = cmd.Parameters(3)
        Me![taisyo_yyyy] = Year(cmd.Parameters(3))
        Me![taisyo_mm] = Month(cmd.Parameters(3))
    End If
    
    '経費合計
    With cmd
        .CommandText = "usp_import_keihi_gokei"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
    End With
    
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
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
    
    Set cmd = Nothing

    If Not IsNull(Me![chk_shanghai_chotatsu]) And Me![chk_shanghai_chotatsu] = 3 Then
        Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_boueki"
        Me![shiiresaki_code].RowSource = "View__shiiresaki_code_shanghai"
    Else
        If Not IsNull(Me![chk_shanghai_chotatsu]) And Me![chk_shanghai_chotatsu] = 2 Then
            Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_chotatsu"
            Me![shiiresaki_code].RowSource = "View__shiiresaki_code_chotatsu"
        Else
            Me![pk_invoice_no_main].RowSource = "View__invoice_no_main"
            Me![shiiresaki_code].RowSource = "View__shiiresaki_code_shanghai"
        End If
    End If
    Me![pk_invoice_no_main].Requery
    Me![shiiresaki_code].Requery

    '進捗状況更新
    Call schedule_update(17, Date, Me![pk_invoice_no_main], Me![shanghai_chotatsu])

End Sub

Private Sub btn_hoken_henko_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![insurance_seikyu_ymd]) Then
        MsgBox ("保険請求日を入力してください。")
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_import_keihi_seikyubi_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_invoice_no_main]
        .Parameters(2) = Me![insurance_seikyu_ymd]
        .Execute
    End With
    
    Set cmd = Nothing
    
    MsgBox ("請求日を変更しました。")
    

End Sub

Private Sub btn_matome_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![taisyo_yyyy]) Or IsNull(Me![taisyo_mm]) Then
        MsgBox ("対象年月を入力してください。")
        Exit Sub
    Else
        If IsDate(Me![taisyo_yyyy] & "/" & Me![taisyo_mm] & "/" & 1) = False Then
            MsgBox ("対象年月が不正です。")
            Exit Sub
        End If
    End If
    If IsNull(Me![chk_shanghai_chotatsu]) Or Me![chk_shanghai_chotatsu] = 0 Then
        MsgBox ("処理選択を選択してください。")
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_mail_soshin_invoice_no_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![chk_shanghai_chotatsu]
        .Parameters(3) = Me![taisyo_yyyy] * 100 + Me![taisyo_mm]
        .Execute
    End With
    
    Set cmd = Nothing

    intMenuNo = 301
    intMenuNoSeq = Me![chk_shanghai_chotatsu]
    strInputCode = Me![taisyo_yyyy] * 100 + Me![taisyo_mm]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_modoru_Click()

    If IsNull(Me![return_form_name]) Then
        strFormName = "F_0_START_MENU"
    Else
        strFormName = Me![return_form_name]
    End If
    intMenuNo = 0
    intMenuNoSeq = 0
    strOpenMode = "OC"
    
    Call next_form_open
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_3_輸入諸経費入力_MAIN")
    intMenuNo = 301
    intMenuNoSeq = 0
    If intShanghaiChotatsu = 0 Then
        intShanghaiChotatsu = 1
    End If

End Sub

Private Sub btn_shanghai_chotatsu_Click()

    'If Me![btn_shanghai_chotatsu].Caption = "上海" Then
    '    Me![btn_shanghai_chotatsu].Caption = "調達"
    '    intShanghaiChotatsu = 2
    '    Me![chk_shanghai_chotatsu] = 2
    'Else
    '    If Me![btn_shanghai_chotatsu].Caption = "調達" Then
    '        Me![btn_shanghai_chotatsu].Caption = "貿易"
    '        intShanghaiChotatsu = 3
    '      Me![chk_shanghai_chotatsu] = 1
    '    Else
    '        Me![btn_shanghai_chotatsu].Caption = "上海"
    '        intShanghaiChotatsu = 1
    '        Me![chk_shanghai_chotatsu] = 1
    '    End If
    'End If
    
    'Me![pk_invoice_no_main] = Null
    
    'strFormName = "F_3_輸入諸経費入力_MAIN"
    'Call next_form_open
    
    'Call chk_shanghai_chotatsu_Click

End Sub

Private Sub chk_shanghai_chotatsu_Click()

    If IsNull(Me![chk_shanghai_chotatsu]) Or Me![chk_shanghai_chotatsu] = 0 Then
        Exit Sub
    End If
    
    Select Case Me![chk_shanghai_chotatsu]
        Case 3
            Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_boueki"
            Me![shiiresaki_code].RowSource = "View__shiiresaki_code_shanghai"
        Case 2
            Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_chotatsu"
            Me![shiiresaki_code].RowSource = "View__shiiresaki_code_chotatsu"
        Case Else
            Me![pk_invoice_no_main].RowSource = "View__invoice_no_main"
            Me![shiiresaki_code].RowSource = "View__shiiresaki_code_shanghai"
    End Select
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
    intMenuNo = 301
    intMenuNoSeq = 0
   
    Select Case Me![chk_shanghai_chotatsu]
        Case 3
            Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_boueki"
            Me![shiiresaki_code].RowSource = "View__shiiresaki_code_shanghai"
        Case 2
            Me![pk_invoice_no_main].RowSource = "View__invoice_no_main_chotatsu"
            Me![shiiresaki_code].RowSource = "View__shiiresaki_code_chotatsu"
        Case Else
            Me![pk_invoice_no_main].RowSource = "View__invoice_no_main"
            Me![shiiresaki_code].RowSource = "View__shiiresaki_code_shanghai"
    End Select
    
    Me![pk_invoice_no_main].Requery
    Me![shiiresaki_code].Requery
    
    Me![chk_shanghai_chotatsu] = intShanghaiChotatsu
    
End Sub

Private Sub pk_invoice_no_main_AfterUpdate()
    
    If IsNull(Me![pk_invoice_no_main]) Then
        Exit Sub
    End If
    
    Call btn_gokei_kakunin_Click
    
    DoCmd.GoToRecord , , acFirst
    
    Me![F_3_輸入諸経費入力_SUB].Requery
    
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
    If strFormName = "F_3_輸入諸経費入力_MAIN" Then
        strOpenMode = "O"
    Else
        strOpenMode = "OC"
    End If
    
    Call form_open_close("F_3_輸入諸経費入力_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

Private Sub shanghai_chotatsu_AfterUpdate()

    Select Case Me![shanghai_chotatsu]
        Case 1
            intShanghaiChotatsu = 1
            Me![chk_shanghai_chotatsu] = 1
        Case 2
            intShanghaiChotatsu = 2
            Me![chk_shanghai_chotatsu] = 2
        Case 3
            intShanghaiChotatsu = 3
            Me![chk_shanghai_chotatsu] = 1
        Case Else
            intShanghaiChotatsu = 1
            Me![chk_shanghai_chotatsu] = 1
    End Select
    
    Me![pk_invoice_no_main] = Null
    
    strFormName = "F_3_輸入諸経費入力_MAIN"
    Call next_form_open
    
    Call chk_shanghai_chotatsu_Click

End Sub
