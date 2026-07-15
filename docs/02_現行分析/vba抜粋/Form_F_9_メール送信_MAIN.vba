Attribute VB_Name = "Form_F_9_メール送信_MAIN"
Attribute VB_Base = "0{D50F433E-F9E0-4EC0-BE42-ACD4C0A40890}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_menu_Click()

    strFormName = "F_0_START_MENU"
    Call next_form_open

End Sub

Private Sub btn_setsuzoku_Click()
    
    '再接続
    Call Login_ReInput("F_9_メール送信_MAIN")

End Sub

Private Sub btn_soshin_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![excel_name]) Then
        MsgBox ("EXCEL名称を入力してください。")
        Exit Sub
    End If
    
    If IsNull(Me![mail_subject]) Then
        MsgBox ("件名を入力してください。")
        Exit Sub
    End If
    
    'エラーでコネクションが切れると、グローバル変数はクリアされてしまうので、入力内容からMENU_NOをリセット
    If intMenuNo = 0 Then
        If Not IsNull(Me![menu_no]) Then
            intMenuNo = Me![menu_no]
            If Not IsNull(Me![menu_no_seq]) Then
                intMenuNoSeq = Me![menu_no_seq]
            End If
        Else
            Call mail_soshin_menu_no_reset
            If intMenuNo = 0 Then
                MsgBox ("接続が一度切れたため、処理情報が消されてしまっています。" & vbLf & "【戻るボタン】で一度該当処理フォームに戻って、再度実行してください。")
                Exit Sub
            End If
        End If
    End If
    
    Dim intPickupCount As Integer
    Dim strStoredName As String
    Dim strStoredPara As String
    Dim intStoredPara As Integer
    
    Set cmd.ActiveConnection = conn
    
    'ORDER_NOまたは、INVOCIE_NOの入力が必要かどうか
    With cmd
        .CommandText = "usp_koumoku_contents_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "MAL"
        .Parameters(2) = intMenuNo
        .Execute
    End With
    'あったら、未入力チェック（エラーはスルーにしておこう）
    On Error Resume Next
    If Not IsNull(cmd.Parameters(4)) Then
        If IsNull(Me![input_code]) Then
            MsgBox (cmd.Parameters(4) & "を入力してください。")
            GoTo Exit_btn_soshin_Click
        End If
    End If
    On Error GoTo 0
    
    '送信テーブル名
    With cmd
        .CommandText = "usp_koumoku_contents_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "SND"
        .Parameters(2) = intMenuNo * 10 + intMenuNoSeq
        .Execute
    End With
    If Not IsNull(cmd.Parameters(3)) Then
        strSql = "DELETE " & cmd.Parameters(3) & " FROM " & cmd.Parameters(3) & " "
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
    End If
    
    'ストアド、ストアドのパラメータ、確認フォーム名称の取得
    With cmd
        .CommandText = "usp_koumoku_contents_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = "MCR"
        .Parameters(2) = intMenuNo * 10 + intMenuNoSeq
        .Execute
    End With
    
    If IsNull(cmd.Parameters(3)) Then
        MsgBox ("送信設定がされていません。" & vbLf & "IT室に問い合わせてください。")
        GoTo Exit_btn_soshin_Click
    End If
    
    strStoredName = cmd.Parameters(3)
    If Len(cmd.Parameters(4)) > 0 Then
        strFormName = cmd.Parameters(4)
    Else
        strFormName = ""
    End If
    If Not IsNull(cmd.Parameters(5)) Then
        strStoredPara = cmd.Parameters(5)
    Else
        strStoredPara = "LOGIN_CODE"
    End If
    If Not IsNull(cmd.Parameters(7)) Then
        intStoredPara = cmd.Parameters(7)
    Else
        intStoredPara = 0
    End If
    intPickupCount = 0
    
    '対象ORDER、またはINVOICEをEXPORT用のテンポラリに抽出
    If intStoredPara <> 0 Then
        With cmd
            .CommandText = strStoredName
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me(strStoredPara)
            .Parameters(2) = intStoredPara
            .Execute
        End With
        If Not IsNull(cmd.Parameters(3)) Then
            intPickupCount = cmd.Parameters(3)
        End If
    Else
        With cmd
            .CommandText = strStoredName
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me(strStoredPara)
            .Execute
        End With
        If Not IsNull(cmd.Parameters(2)) Then
            intPickupCount = cmd.Parameters(2)
        End If
    End If
        
    
    'intMenuNo = Me![chk_menu_no]
    If Not IsNull(Me![chk_menu_no_seq]) Then
        If Me![chk_menu_no_seq] = 0 Then
            intMenuNoSeq = 1
            Me![menu_no_seq] = 1
        Else
            intMenuNoSeq = Me![chk_menu_no_seq]
            Me![menu_no_seq] = Me![chk_menu_no_seq]
        End If
    End If
    
    If intPickupCount = 0 Then
        MsgBox ("対照データはありませんでした。")
        GoTo Exit_btn_soshin_Click
    End If
    
    
    If Len(strFormName) > 0 Then
        ReturnValue = MsgBox("内容を確認しますか？" & "確認（はい）／このまま送信（いいえ）", vbYesNo)
        If ReturnValue = vbYes Then
        
            On Error Resume Next
            strSql = "insert into tbl_t_mail_soshin_rireki (pk_menu_no, pk_menu_no_seq, soshin_login, soshin_ymd, soshin_status, bikou) "
            strSql = strSql & "VALUES (" & intMenuNo & "," & intMenuNoSeq & ", '" & Me![login_code] & "', '" & Now() & "', 'DIS', '"
            strSql = strSql & "form_type=" & Me![kakunin_form] & "' ) "
            With cmd
                .CommandText = strSql
                .CommandType = adCmdText
                .Execute
            End With
            On Error GoTo 0
            
            If Me![kakunin_form] = 2 Then
                DoCmd.OpenForm strFormName, acFormDS
            Else
                DoCmd.OpenForm strFormName
                Forms(strFormName)![login_code] = Me![login_code]
                Forms(strFormName)![login_name] = Me![login_name]
            End If
            
            GoTo Exit_btn_soshin_Click
            
        End If
    End If
    
    Set cmd = Nothing
    
    ReturnValue = mail_send(intMenuNo, intMenuNoSeq)
    
    If ReturnValue = "OK" Then
    
        MsgBox ("メール送信完了しました。")
        
    End If
       
    Exit Sub
    
    
Exit_btn_soshin_Click:
    
    Set cmd = Nothing
    
End Sub

Private Sub btn_modoru_Click()

    If IsNull(Me![return_form_name]) Then
        strFormName = "F_0_START_MENU"
    Else
        strFormName = Me![return_form_name]
    End If
    Call next_form_open
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    strCurrentForm = "F_9_メール送信_MAIN"
    
    Me![chk_menu_no_seq].Visible = False
    
End Sub

Private Sub chk_menu_no_Click()
        
    Call sub_form_display

End Sub

Private Sub chk_menu_no_seq_Click()
        
    Call sub_form_display

End Sub

Private Sub sub_form_display()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If IsNull(Me![chk_menu_no]) Then
        Exit Sub
    End If
    
    'intMenuNo = Me![chk_menu_no]
    If Not IsNull(Me![chk_menu_no_seq]) Then
        If Me![chk_menu_no_seq] = 0 Then
            intMenuNoSeq = 1
            Me![menu_no_seq] = 1
        Else
            intMenuNoSeq = Me![chk_menu_no_seq]
            Me![menu_no_seq] = Me![chk_menu_no_seq]
        End If
    End If
    
    Call mail_soshin_form_open(intMenuNo, intMenuNoSeq)
    
End Sub

Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    strInvoiceNo = "NULL"
    If Me![lbl_input_code].Caption = "INVOICE_NO" Then
        If Not IsNull(Me![input_code]) Then
            strInvoiceNo = Me![input_code]
        End If
    End If
    strOpenMode = "OC"
    
    Call form_open_close("F_9_メール送信_MAIN", strFormName, strOpenMode, strInvoiceNo)
    

End Sub
