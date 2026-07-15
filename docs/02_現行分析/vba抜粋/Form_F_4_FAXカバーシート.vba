Attribute VB_Name = "Form_F_4_FAXカバーシート"
Attribute VB_Base = "0{472AB219-823D-4BA9-9756-F9457120D20B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit


Private Sub bnt_menu_Click()

    'メニューに戻る
    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    strOpenMode = "OC"
    Call form_open_close("F_4_FAXカバーシート", strFormName, strOpenMode, "NULL")
    
End Sub


Private Sub btn_clear_Click()

    Me![midashi] = "FAX送信のご案内"
    Me![pk_soshin_saki] = Null
    Me![soshin_ymd] = Date
    Me![soshin_saki_tanto] = Null
    Me![sofu_maisu] = 1
    Me![youken] = Null
    Me![syurui] = 3
    Me![renraku_jiko1] = Null
    Me![renraku_jiko2] = Null

End Sub

Private Sub btn_print_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_soshin_saki]) Or IsNull(Me![midashi]) Or IsNull(Me![tajima_tanto]) Or IsNull(Me![soshin_ymd]) _
        Or IsNull(Me![sofu_maisu]) Or IsNull(Me![syurui]) Then
            MsgBox ("黄色いフィールドは必須項目です。")
            Exit Sub
    End If

    Set cmd.ActiveConnection = conn
    
    strSql = "DELETE tbl_t_fax_soshin FROM tbl_t_fax_soshin WHERE pk_soshin_saki = " & Me![pk_soshin_saki] & " "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    strSql = "INSERT INTO tbl_t_fax_soshin (pk_soshin_saki, midashi, tajima_tanto, soshin_ymd, sofu_maisu, syurui, soshin_saki_tanto, "
    strSql = strSql & "youken, renraku_jiko1, renraku_jiko2, input_login ) VALUES ( " & Me![pk_soshin_saki] & ",'" & Me![midashi] & "','" & Me![tajima_tanto] & "','"
    strSql = strSql & Me![soshin_ymd] & "'," & Me![sofu_maisu] & "," & Me![syurui] & ","
    If IsNull(Me![soshin_saki_tanto]) Then
        strSql = strSql & "null,"
    Else
        strSql = strSql & "'" & Me![soshin_saki_tanto] & "',"
    End If
    If IsNull(Me![youken]) Then
        strSql = strSql & "null,"
    Else
        strSql = strSql & "'" & Me![youken] & "',"
    End If
    If IsNull(Me![renraku_jiko1]) Then
        strSql = strSql & "null , "
    Else
        strSql = strSql & "'" & Me![renraku_jiko1] & "' , "
    End If
    If IsNull(Me![renraku_jiko2]) Then
        strSql = strSql & "null, "
    Else
        strSql = strSql & "'" & Me![renraku_jiko2] & "', "
    End If
    strSql = strSql & "'" & Me![login_code] & "' ) "
    'Debug.Print strSql
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With

    DoCmd.OpenReport "R_ファックスカバーシート", acViewPreview, , "pk_soshin_saki = " & Me![pk_soshin_saki]
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_4_FAXカバーシート")
    intMenuNo = 404
    intMenuNoSeq = 0

End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 404
    intMenuNoSeq = 0

    Me![midashi] = "FAX送信のご案内"
    
    Set cmd.ActiveConnection = conn
    
    strSql = "SELECT tbl_m_koumoku.contents1 FROM tbl_m_koumoku WHERE pk_koumoku_code = 'FSH' "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    Set rs1 = cmd.Execute(intRcnt)
    
    If intRcnt <> 0 Then
    
        On Error Resume Next
        rs1.MoveFirst
        Do Until rs1.EOF
            Me("lbl_syurui" & rs1![pk_seq_no].Value).Caption = rs1![contents1].Value
            rs1.MoveNext
        Loop
        On Error GoTo 0
        
    End If
    rs1.Close
    
    Set rs1 = Nothing
    Set cmd = Nothing
    
End Sub

Private Sub pk_soshin_saki_AfterUpdate()

    Set cmd.ActiveConnection = conn
    
    strSql = "SELECT tbl_t_fax_soshin.* FROM tbl_t_fax_soshin WHERE pk_soshin_saki = " & Me![pk_soshin_saki] & " "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    Set rs1 = cmd.Execute(intRcnt)
    
    If intRcnt <> 0 Then
        rs1.MoveFirst
        
        On Error Resume Next
        For Each ct In Me
            For n = 0 To rs1.Fields.Count - 1
                If ct.Name = rs1.Fields(n).Name Then
                    Me(ct.Name) = rs1.Fields(n)
                End If
            Next
        Next
        On Error GoTo 0
    
    Else
        rs1.Close
        
        With cmd
            .CommandText = "usp_koumoku_contents_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = "FAX"
            .Parameters(2) = Me![pk_soshin_saki]
            .Execute
        End With
        
        If Not Not IsNull(cmd.Parameters(4)) Then
            Me![soshin_saki_tanto] = cmd.Parameters(4)
        Else
            Me![soshin_saki_tanto] = Null
        End If
        
        Me![soshin_ymd] = Date
        Me![sofu_maisu] = 1
        Me![youken] = Null
        Me![syurui] = 3
        Me![renraku_jiko1] = Null
        Me![renraku_jiko2] = Null
    End If
    rs1.Close
    
    Set rs1 = Nothing
    Set cmd = Nothing
    
    If IsNull(Me![soshin_ymd]) Then
        Me![soshin_ymd] = Date
    End If

End Sub
