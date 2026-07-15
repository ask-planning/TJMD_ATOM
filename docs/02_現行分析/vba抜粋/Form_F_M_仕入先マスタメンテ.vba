Attribute VB_Name = "Form_F_M_仕入先マスタメンテ"
Attribute VB_Base = "0{701EBFC1-1509-4BA7-B07D-54DA341D8D29}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_haiso_tehai_Click()

    '配送手配書の作成処理へ
    strFormName = "F_2_調達輸入_INVOICE_HEADER"
    intMenuNo = 203
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_menu_Click()

    'メニューに戻る
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_sakujyo_Click()

    If IsNull(Me![pk_shiiresaki_code]) Then
        Exit Sub
    End If
    
    ReturnValue = MsgBox("削除しますか？", vbYesNo)
    If ReturnValue = vbNo Then
        Exit Sub
    End If
    
    With cmd
        .CommandText = "usp_invoice_default_master_delete"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_shiiresaki_code]
        .Execute
    End With
    
    MsgBox ("削除完了しました。")
    
    Call pk_shiiresaki_code_AfterUpdate
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_M_仕入先マスタメンテ")
    intMenuNo = 905
    intMenuNoSeq = 0

End Sub

Private Sub btn_torikomi_Click()

    '購買発注取込処理へ
    strFormName = "F_1_R3購買発注取込"
    intMenuNo = 201
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_touroku_Click()

    DoCmd.GoToControl "pk_shiiresaki_code"
    
    If IsNull(Me![pk_shiiresaki_code]) Then
        MsgBox ("仕入先コードを入力してください。")
        Exit Sub
    End If
    
    Dim intFieldCount As Integer
    Dim strTblFieldName() As String
    Dim intTblFieldType() As Integer
    
    
    Set cmd.ActiveConnection = conn
    
    'ヘッダ部作成
    intFieldCount = 18
    
    ReDim strTblFieldName(intFieldCount) As String
    ReDim intTblFieldType(intFieldCount) As Integer
    
    strTblFieldName(0) = "pk_shiiresaki_code":          intTblFieldType(0) = 0
    strTblFieldName(1) = "shiiresaki_name_s":           intTblFieldType(1) = 0
    strTblFieldName(2) = "shiiresaki_name":             intTblFieldType(2) = 0
    strTblFieldName(3) = "shiiresaki_name_JPN":         intTblFieldType(3) = 0
    strTblFieldName(4) = "syukka_hoho":                 intTblFieldType(4) = 1
    strTblFieldName(5) = "ETD":                         intTblFieldType(5) = 0
    strTblFieldName(6) = "ETA":                         intTblFieldType(6) = 0
    strTblFieldName(7) = "haiso_type":                  intTblFieldType(7) = 1
    strTblFieldName(8) = "otsunaka_no":                 intTblFieldType(8) = 1
    strTblFieldName(9) = "shinchoku_jyokyo_syurui":     intTblFieldType(9) = 1
    strTblFieldName(10) = "invoice_maisu":              intTblFieldType(10) = 1
    strTblFieldName(11) = "packing_maisu":              intTblFieldType(11) = 1
    strTblFieldName(12) = "bl_maisu":                   intTblFieldType(12) = 1
    strTblFieldName(13) = "arrival_notice_maisu":       intTblFieldType(13) = 1
    strTblFieldName(14) = "insurance_maisu":            intTblFieldType(14) = 1
    strTblFieldName(15) = "bl_option":                  intTblFieldType(15) = 1
    strTblFieldName(16) = "update_ymd":                 intTblFieldType(16) = 3
    strTblFieldName(17) = "update_login":               intTblFieldType(17) = 3
    
    If Me![syori_kbn_text] = "新規" Then
    
        strSql = "INSERT INTO tbl_m_invoice_default ("
        For n = 0 To intFieldCount - 1
            If n <> 0 Then
                strSql = strSql & ","
            End If
            strSql = strSql & strTblFieldName(n)
        Next
        strSql = strSql & ") VALUES ("
        For n = 0 To intFieldCount - 1
            If n <> 0 Then
                strSql = strSql & ","
            End If
            Select Case intTblFieldType(n)
                Case 0
                    If IsNull(Me(strTblFieldName(n))) Or Len(Me(strTblFieldName(n))) <= 0 Then
                        strSql = strSql & "null"
                    Else
                        strSql = strSql & "'" & Me(strTblFieldName(n)) & "'"
                    End If
                Case 1
                    If IsNull(Me(strTblFieldName(n))) Or Len(Me(strTblFieldName(n))) <= 0 Then
                        strSql = strSql & "null"
                    Else
                        strSql = strSql & Me(strTblFieldName(n))
                    End If
                Case 3
                    If strTblFieldName(n) = "update_ymd" Then
                        strSql = strSql & "'" & Now() & "'"
                    End If
                    If strTblFieldName(n) = "update_login" Then
                        strSql = strSql & "'" & Me![login_code] & "'"
                    End If
            End Select
        Next
        strSql = strSql & " )"
                
    Else
    
        strSql = "UPDATE tbl_m_invoice_default SET "
        For n = 0 To intFieldCount - 1
            If n <> 0 Then
                strSql = strSql & ","
            End If
            Select Case intTblFieldType(n)
                Case 0
                    If IsNull(Me(strTblFieldName(n))) Or Len(Me(strTblFieldName(n))) <= 0 Then
                        strSql = strSql & strTblFieldName(n) & " = null"
                    Else
                        strSql = strSql & strTblFieldName(n) & " = '" & Me(strTblFieldName(n)) & "'"
                    End If
                Case 1
                    If IsNull(Me(strTblFieldName(n))) Or Len(Me(strTblFieldName(n))) <= 0 Then
                        strSql = strSql & strTblFieldName(n) & " = 0"
                    Else
                        strSql = strSql & strTblFieldName(n) & " = " & Me(strTblFieldName(n))
                    End If
                Case 3
                    If strTblFieldName(n) = "update_ymd" Then
                        strSql = strSql & strTblFieldName(n) & " = '" & Now() & "'"
                    End If
                    If strTblFieldName(n) = "update_login" Then
                        strSql = strSql & strTblFieldName(n) & " = '" & Me![login_code] & "'"
                    End If
            End Select
        Next
        strSql = strSql & " WHERE pk_shiiresaki_code = '" & Me![pk_shiiresaki_code] & "' "
        
    End If
    
    Debug.Print strSql
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    MsgBox ("登録完了しました。")
    
    Set cmd = Nothing
    
    Call pk_shiiresaki_code_AfterUpdate

End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 905
    intMenuNoSeq = 0
    
End Sub

Private Sub pk_shiiresaki_code_AfterUpdate()

    If IsNull(Me![pk_shiiresaki_code]) Then
        Exit Sub
    End If
    
    If Len(Me![pk_shiiresaki_code]) < 6 Or Len(Me![pk_shiiresaki_code]) > 10 Then
        MsgBox ("仕入先コードの桁数が違います。")
        Exit Sub
    End If
    
    If Len(Me![pk_shiiresaki_code]) < 10 Then
        Me![pk_shiiresaki_code] = Right("0000000000" & Me![pk_shiiresaki_code], 10)
        Me![pk_shiiresaki_code].Requery
    End If
    
    Dim ct As Control
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_shiiresaki_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_shiiresaki_code]
        .Execute
    End With
    If IsNull(cmd.Parameters(2)) Then
        MsgBox ("R3に存在しない仕入先コードです。")
        GoTo Exit_pk_shiiresaki_code_afterupdate
    End If
    
    On Error Resume Next
    For Each ct In Me
        If ct.Name <> "pk_shiiresaki_code" And Left(ct.Name, 5) <> "login" And ct.Name <> "bumon_code" And ct.Name <> "return_form_name" Then
            Me(ct.Name) = Null
        End If
    Next
    On Error GoTo 0
    
    strSql = "SELECT tbl_m_invoice_default.* FROM tbl_m_invoice_default WHERE pk_shiiresaki_code = '" & Me![pk_shiiresaki_code] & "' "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    Set rs1 = cmd.Execute(intRcnt)
    
    If intRcnt <> 0 Then
    
        Me![syori_kbn_text] = "変更"
        rs1.MoveFirst
        
        For n = 0 To rs1.Fields.Count - 1
            For Each ct In Me
                If ct.Name = rs1.Fields(n).Name Then
                    Me(ct.Name) = rs1.Fields(n).Value
                End If
            Next
        Next
            
    Else
        Me![syori_kbn_text] = "新規"
    End If
    rs1.Close
    
    Set rs1 = Nothing
    
Exit_pk_shiiresaki_code_afterupdate:
    
    Set cmd = Nothing
    
    Me![pk_shiiresaki_code].Requery
    
End Sub

Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    strOpenMode = "OC"
    
    Call form_open_close("F_M_仕入先マスタメンテ", strFormName, strOpenMode, "NULL")
    
End Sub

