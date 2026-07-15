Attribute VB_Name = "Form_F_M_社内参照番号_発番_MAIN"
Attribute VB_Base = "0{0C7A68BC-2460-4ECD-99E3-9B79475A40C9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit
Private Sub bnt_hatsuban_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If Me![kigo] <> 1 And IsNull(Me![shiiresaki_code]) Then
        MsgBox ("発注先を選択してください。")
        Exit Sub
    End If

    Dim intSansyoBango As Integer
    Dim intLen As Integer
    
    Me![renban] = Null
    Me![remarks] = Null
    Me![hacchu_tanto] = Null
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_syanai_sansyo_no_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![kigo]
        .Execute
    End With
    
    If Not IsNull(cmd.Parameters(2)) Then
        intSansyoBango = cmd.Parameters(2) + 1
    Else
        intSansyoBango = 1
    End If
    
    Set cmd = Nothing
    
    If intSansyoBango > 999 Then
        intLen = 4
    Else
        intLen = 3
    End If
    
    Select Case Me![kigo]
        Case 1
            Me![pk_syanai_sansyo_no] = "KH" & Right("0000" & intSansyoBango, intLen)
        Case 2
            Me![pk_syanai_sansyo_no] = "HSZ" & Right("0000" & intSansyoBango, intLen)
        Case 3
            Me![pk_syanai_sansyo_no] = "TNB" & Right("0000" & intSansyoBango, intLen)
    End Select
    Me![renban] = intSansyoBango
    Me![hacchu_tanto] = Me![login_name]
    
End Sub

Private Sub btn_hosozai_input_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    '支給包装材発注処理へ
    strFormName = "F_1_支給包装材発注_MAIN"
    intMenuNo = 108
    intMenuNoSeq = 0
    strOpenMode = "OC"
    Call form_open_close("F_M_社内参照番号_発番_MAIN", strFormName, strOpenMode, "NULL")

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
    Call form_open_close("F_M_社内参照番号_発番_MAIN", strFormName, strOpenMode, "NULL")

End Sub

Private Sub btn_setsuzoku_Click()
    
    '再接続
    Call Login_ReInput("F_M_社内参照番号_発番_MAIN")
    intMenuNo = 902
    intMenuNoSeq = 0

End Sub

Private Sub btn_touroku_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![kigo]) Or IsNull(Me![pk_syanai_sansyo_no]) Then
        MsgBox ("参照番号が選択/入力されていません。")
        Exit Sub
    End If
    

    If Me![kigo] <> 1 And IsNull(Me![shiiresaki_code]) Then
        MsgBox ("発注先を選択してください。")
        Exit Sub
    End If
    
    
    '入力した文字列から、記号を判断
    'If IsNull(me![kigo]) Then
        If Left(Me![pk_syanai_sansyo_no], 2) = "KH" Then
            Me![kigo] = 1
            If IsNumeric(Mid(Me![pk_syanai_sansyo_no], 3, 4)) Then
                Me![renban] = Mid(Me![pk_syanai_sansyo_no], 3, 4)
            Else
                If Not IsNumeric(Mid(Me![pk_syanai_sansyo_no], 3, 3)) Then
                    MsgBox ("社内参照番号の数字部分は３桁です。")
                    Exit Sub
                Else
                    Me![renban] = Mid(Me![pk_syanai_sansyo_no], 3, 3)
                End If
            End If
        Else
            If Left(Me![pk_syanai_sansyo_no], 3) = "HSZ" Then
                Me![kigo] = 2
                If Not IsNumeric(Mid(Me![pk_syanai_sansyo_no], 4, 3)) Then
                    MsgBox ("社内参照番号の数字部分は３桁です。")
                    Exit Sub
                Else
                    Me![renban] = Mid(Me![pk_syanai_sansyo_no], 4, 3)
                End If
            Else
                If Left(Me![pk_syanai_sansyo_no], 3) = "TNB" Then
                    Me![kigo] = 3
                    If Not IsNumeric(Mid(Me![pk_syanai_sansyo_no], 4, 3)) Then
                        MsgBox ("社内参照番号の数字部分は３桁です。")
                        Exit Sub
                    Else
                        Me![renban] = Mid(Me![pk_syanai_sansyo_no], 4, 3)
                    End If
                Else
                    MsgBox ("社内参照番号は、『KH』、『HSZ』または、『TNB』で始まる番号を入力してください。")
                    Exit Sub
                End If
            End If
        End If
    'End If
    
    Set cmd.ActiveConnection = conn
    
    '登録済みの番号のチェック
    With cmd
        .CommandText = "usp_syanai_sansyo_no_sonzai_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_syanai_sansyo_no]
        .Execute
    End With
    If IsNull(cmd.Parameters(2)) Then
        intShinkiFlg = 1
    Else
        intShinkiFlg = 0
    End If
    
    If intShinkiFlg = 0 Then
        ReturnValue = MsgBox("登録済みの社内参照番号です、上書きしますか？", 4)
        If ReturnValue = vbNo Then
            GoTo Exit_btn_touroku_click
        End If
    End If
    
    '社内参照番号管理テーブル更新
    Dim strPara() As String
    Dim intPara() As Integer
    
    intFieldCount = 10
    ReDim strPara(intFieldCount) As String
    ReDim intPara(intFieldCount) As Integer
    
    strPara(0) = "pk_syanai_sansyo_no":         intPara(0) = 0
    strPara(1) = "kigo":                        intPara(1) = 1
    strPara(2) = "renban":                      intPara(2) = 1
    strPara(3) = "shiiresaki_code":             intPara(3) = 0
    strPara(4) = "hacchu_tanto":                intPara(4) = 0
    strPara(5) = "remarks":                     intPara(5) = 0
    strPara(6) = "oem_flg":                     intPara(6) = 1
    strPara(7) = "syanai_sansyo_no_before":     intPara(7) = 0
    strPara(8) = "update_ymd":                  intPara(8) = 0
    strPara(9) = "update_login":                intPara(9) = 0
    
    'sql文作成
    If intShinkiFlg = 1 Then
        '新規
        strSql = "INSERT INTO tbl_m_syanai_sansyo_no ("
        For n = 0 To intFieldCount - 1
            If n <> 0 Then
                strSql = strSql & ", "
            End If
            strSql = strSql & strPara(n)
        Next
        strSql = strSql & " ) VALUES ( "
        For n = 0 To intFieldCount - 4
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
        strSql = strSql & ", '" & Me![pk_syanai_sansyo_no] & "', '" & Now() & "', '" & Me![login_code] & "' ) "
                
    Else
    
        '更新
        strSql = "UPDATE tbl_m_syanai_sansyo_no SET "
        For n = 0 To intFieldCount - 3
            If n <> 0 Then
                strSql = strSql & ", "
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
        strSql = strSql & "update_ymd = '" & Now() & "' , syanai_sansyo_no_before = '" & Me![pk_syanai_sansyo_no] & "', "
        strSql = strSql & "update_login = '" & Me![login_code] & "' "
        strSql = strSql & "WHERE [pk_syanai_sansyo_no] = '" & Me![pk_syanai_sansyo_no] & "' "
        
    End If
    
    'Debug.Print strSql
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    Me![F_M_社内参照番号_発番_SUB].Requery
    
    MsgBox ("登録完了しました。")
    
Exit_btn_touroku_click:
    
    Set cmd = Nothing
    
    
End Sub

Private Sub kigo_Click()

    If IsNull(Me![kigo]) Or Me![kigo] = 0 Then
        Exit Sub
    End If
    
    If Not IsNull(Me![pk_syanai_sansyo_no]) Then
        If (Me![kigo] = 1 And (Left(Me![pk_syanai_sansyo_no], 3) = "HSZ" Or Left(Me![pk_syanai_sansyo_no], 3) = "TNB")) _
            Or (Me![kigo] = 2 And (Left(Me![pk_syanai_sansyo_no], 2) = "KH" Or Left(Me![pk_syanai_sansyo_no], 3) = "TNB")) _
            Or (Me![kigo] = 3 And (Left(Me![pk_syanai_sansyo_no], 2) = "KH" Or Left(Me![pk_syanai_sansyo_no], 3) = "HSZ")) Then
                Me![pk_syanai_sansyo_no] = Null
                Me![renban] = Null
                Me![hacchu_tanto] = Null
                Me![remarks] = Null
                Me![shiiresaki_code] = Null
                Me![OEM_FLG] = 0
        End If
    End If
    
    
    strSql = "SELECT pk_syanai_sansyo_no, shiiresaki_code, remarks, renban FROM dbo.tbl_m_syanai_sansyo_no "
    strSql = strSql & "WHERE kigo = " & Me![kigo] & " ORDER BY renban DESC "
    Me![pk_syanai_sansyo_no].RowSource = strSql
    
    Me![pk_syanai_sansyo_no].Requery
    
    Me![F_M_社内参照番号_発番_SUB].Requery
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    intMenuNo = 902
    intMenuNoSeq = 0
    
End Sub

Private Sub pk_syanai_sansyo_no_AfterUpdate()
    
    If IsNull(Me![pk_syanai_sansyo_no]) Or Len(Me![pk_syanai_sansyo_no]) <= 0 Then
        Exit Sub
    End If
    
    If Left(Me![pk_syanai_sansyo_no], 2) <> "KH" And Left(Me![pk_syanai_sansyo_no], 3) <> "HSZ" Then
        MsgBox ("社内参照番号は、『KH』、『HSZ』または、『TNB』で始まる番号を入力してください。")
    End If
    
    Set cmd.ActiveConnection = conn
    
    strSql = "SELECT tbl_m_syanai_sansyo_no.* FROM tbl_m_syanai_sansyo_no WHERE pk_syanai_sansyo_no = '" & Me![pk_syanai_sansyo_no] & "' "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    Set rs1 = cmd.Execute(intRcnt)
    On Error Resume Next
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
    Else
        If Left(Me![pk_syanai_sansyo_no], 2) = "KH" Then
            Me![kigo] = 1
            If IsNumeric(Mid(Me![pk_syanai_sansyo_no], 3, 4)) Then
                Me![renban] = Mid(Me![pk_syanai_sansyo_no], 3, 4)
            Else
                Me![renban] = Mid(Me![pk_syanai_sansyo_no], 3, 3)
            End If
        Else
            If Left(Me![pk_syanai_sansyo_no], 3) = "HSZ" Then
                Me![kigo] = 2
                Me![renban] = Mid(Me![pk_syanai_sansyo_no], 4, 3)
            Else
                If Left(Me![pk_syanai_sansyo_no], 3) = "TNB" Then
                    Me![kigo] = 3
                    Me![renban] = Mid(Me![pk_syanai_sansyo_no], 4, 3)
                End If
            End If
        End If
    End If
    rs1.Close
    
    Set rs1 = Nothing
    Set cmd = Nothing
        

End Sub
