Attribute VB_Name = "Form_F_1_支給包装材発注_MAIN"
Attribute VB_Base = "0{6D358C05-33CF-46A2-9EFA-25F5E8AFB3D6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_clear_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    Me![pk_syanai_sansyo_no] = Null
    Me![plant] = Null
    Me![hokan_basyo] = Null
    Me![touroku_tanto] = Null
    Me![touroku_ymd] = Null
    Me![koubai_hacchu_no] = Null
    Me![BL_DATE] = Null
    Me![touroku_tanto] = Me![login_code]
    'Me![pk_shiiresaki_code] = Null
    'Me![pk_shiiresaki_code].Requery
    
    Set cmd.ActiveConnection = conn
    
    strSql = "DELETE tbl_t_shikyu_hosozai_hacchu_temp FROM tbl_t_shikyu_hosozai_hacchu_temp WHERE pk_login_code = '" & Me![login_code] & "' "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    Set cmd = Nothing
    
    Me![F_1_支給包装材発注_SUB1].Requery
    Me![F_1_支給包装材発注_SUB2].Requery
    
End Sub

Private Sub btn_hacchu_soshin_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_syanai_sansyo_no]) Then
        MsgBox ("作成するORDERを抽出して、送信ボタンをクリックして下さい。")
        Exit Sub
    End If

    Set cmd.ActiveConnection = conn
    
    ReturnValue = MsgBox("R3サーバーにテキストを送信します。" & vbLf & "はい／送信　いいえ／キャンセル", 4)
    If ReturnValue = vbNo Then
        Exit Sub
    End If
    
    Dim intHacchuCount As Integer
    Dim intOrderSeq As Integer
    Dim strKeyPlant As String
    Dim strKeyHokanBasyo As String
    
    '入力データの有無確認
    With cmd
        .CommandText = "usp_shikyu_hosozai_temp_count"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    If IsNull(cmd.Parameters(2)) Or cmd.Parameters(2) = 0 Then
        MsgBox ("対象ORDERがありません｡ " & vbLf & "ORDERを抽出して再度実行してください。")
        GoTo Exit_btn_hacchu_soshin_click
    End If
    
    '発注済かどうかの確認
    If IsNull(cmd.Parameters(3)) Or cmd.Parameters(3) = 0 Then
        ReturnValue = MsgBox("R3送信済みのORDERです。 " & vbLf & "再送しますか？", vbYesNo + vbDefaultButton2)
        If ReturnValue = vbNo Then
            GoTo Exit_btn_hacchu_soshin_click
        End If
    End If
    
    'ーHから、個別包装材への分解データ作成、R3バッチ登録用データ作成
    strFileName = Me![file_name]
    
    With cmd
        .CommandText = "usp_shikyu_hosozai_bunkai_meisai_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_syanai_sansyo_no]
        .Execute
    End With
    
    With cmd
        .CommandText = "usp_shikyu_hosozai_text_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    strSql = "SELECT dbo.tbl_t_R3_hacchu_text_temp.* FROM dbo.tbl_t_R3_hacchu_text_temp "
    strSql = strSql & "ORDER BY syanai_sansyo_no, nounyu_kijitsu_ymd, hinmoku_code "
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
    
        intFieldCount = rs1.Fields.Count

        strTextName = Me![file_name]

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
        .CommandText = "usp_shikyu_hosozai_hacchu_soshin_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    
    MsgBox ("テキストデータを送信しました。" & "Ｒ３でバッチインプットを行ってください。")
    
    Me![F_1_支給包装材発注_SUB1].Requery
    Me![F_1_支給包装材発注_SUB2].Requery
    
Exit_btn_hacchu_soshin_click:

    Set cmd = Nothing

End Sub

Private Sub btn_hatsuban_Click()

    'R3購買発注明細へ
    strFormName = "F_M_社内参照番号_発番_MAIN"
    intMenuNo = 902
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_hosozai_pickup_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    '品目振替処理用：　対象品目抽出（分解前本体)

    If IsNull(Me![pk_syanai_sansyo_no]) Then
        MsgBox ("社内参照番号を入力してください。")
        Exit Sub
    End If
    
    If Not IsNull(Me![hokan_basyo]) Or Not IsNull(Me![koubai_hacchu_no]) Then
        MsgBox ("抽出は社内参照番号単位です。" & vbLf & "保管場所、購買発注NOは無視されます。")
    End If

    intMenuNo = 108
    intMenuNoSeq = 3
    strInputCode = Me![pk_syanai_sansyo_no]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_ko_bunkai_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    '品目振替処理用：　対象品目抽出（分解後明細)

    If IsNull(Me![pk_syanai_sansyo_no]) Then
        MsgBox ("社内参照番号を入力してください。")
        Exit Sub
    End If
    
    If Not IsNull(Me![hokan_basyo]) Or Not IsNull(Me![koubai_hacchu_no]) Then
        MsgBox ("抽出は社内参照番号単位です。" & vbLf & "保管場所、購買発注NOは無視されます。")
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_mail_soshin_temp_select_code_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_syanai_sansyo_no]
        .Execute
    End With
    
    Set cmd = Nothing

    intMenuNo = 108
    intMenuNoSeq = 2
    strInputCode = Me![pk_syanai_sansyo_no]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_koubai_hacchu_Click()

    'R3購買発注明細へ
    strFormName = "F_1_R3購買発注_MAIN"
    intMenuNo = 105
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub btn_pickup_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    If Not IsNull(Me![pk_syanai_sansyo_no]) Then
        strSyanaiSansyoNo = "%" & Me![pk_syanai_sansyo_no] & "%"
    Else
        strSyanaiSansyoNo = "%%"
    End If
    If Not IsNull(Me![hokan_basyo]) Then
        strNounyuPlant = "%" & Me![hokan_basyo] & "%"
    Else
        strNounyuPlant = "%%"
    End If
    If Not IsNull(Me![koubai_hacchu_no]) Then
        strKoubaiHacchuNo = "%" & Me![koubai_hacchu_no] & "%"
    Else
        strKoubaiHacchuNo = "%%"
    End If
    If Not IsNull(Me![BL_DATE]) Then
        datBLDate = Me![BL_DATE]
    Else
        datBLDate = #1/1/1900#
    End If
    
    Me![chk_shinki_kizon] = 2
    Call chk_shinki_kizon_Click

    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_shikyu_hosozai_hacchu_temp_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = strSyanaiSansyoNo
        .Parameters(3) = strNounyuPlant
        .Parameters(4) = strKoubaiHacchuNo
        .Parameters(5) = datBLDate
        .Execute
    End With
    If IsNull(cmd.Parameters(6)) Or cmd.Parameters(6) = 0 Then
        Me![plant] = Null
        Me![hokan_basyo] = Null
        Me![touroku_tanto] = Null
        Me![touroku_ymd] = Null
        Me![koubai_hacchu_no] = Null
        Me![BL_DATE] = Null
        MsgBox ("対象ORDERはありませんでした。")
    Else
    
        '１ORDERのみだったら、ヘッダ情報再表示
        If cmd.Parameters(6) = 1 Then
            With cmd
                .CommandText = "usp_shikyu_hosozai_hacchu_header_pickup"
                .CommandType = adCmdStoredProc
                .Parameters.Refresh
                .Parameters(1) = Me![login_code]
            End With
            Set rs1 = cmd.Execute(intRcnt)
            If intRcnt <> 0 Then
                rs1.MoveFirst
                On Error Resume Next
                For n = 0 To rs1.Fields.Count - 1
                    For Each ct In Me
                        If ct.Name = rs1.Fields(n).Name Then
                            Me(ct.Name) = rs1.Fields(n).Value
                        End If
                    Next
                Next
                On Error GoTo 0
            End If
            rs1.Close
            
            Me![pk_shiiresaki_code].Requery
            Me![koubai_hacchu_no].Requery
            Me![hokan_basyo].Requery
            Me![touroku_tanto].Requery
        End If
        
    End If
    
    Set rs1 = Nothing
    Set cmd = Nothing
    
    Me![F_1_支給包装材発注_SUB1].Requery
    

End Sub

Private Sub btn_recreate_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    If IsNull(Me![pk_syanai_sansyo_no]) Then
        MsgBox ("社内参照番号を入力して、作成するORDERを抽出して下さい。")
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    '対象データがあるか？
    With cmd
        .CommandText = "usp_shikyu_hosozai_temp_count"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    If IsNull(cmd.Parameters(2)) Or cmd.Parameters(2) = 0 Then
        MsgBox ("対象ORDERがありません｡ " & vbLf & "ORDERを抽出して再度実行してください。")
        GoTo Exit_btn_recreate_click
    End If
    
    '子品目分解データの再作成
    With cmd
        .CommandText = "usp_shikyu_hosozai_bunkai_meisai_create"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = Me![pk_syanai_sansyo_no]
        .Execute
    End With
    
    MsgBox ("分解後明細を再作成しました。")
    
Exit_btn_recreate_click:

    Set cmd = Nothing

End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_1_支給包装材発注_MAIN")
    intMenuNo = 108
    intMenuNoSeq = 0

End Sub

Private Sub btn_shanghai_renraku_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If

    '上海連絡用：　対象品目抽出

    If IsNull(Me![pk_syanai_sansyo_no]) Then
        MsgBox ("社内参照番号を入力してください。")
        Exit Sub
    End If
    
    If Not IsNull(Me![hokan_basyo]) Or Not IsNull(Me![koubai_hacchu_no]) Then
        MsgBox ("抽出は社内参照番号単位です。" & vbLf & "保管場所、購買発注NOは無視されます。")
    End If

    intMenuNo = 108
    intMenuNoSeq = 1
    strInputCode = Me![pk_syanai_sansyo_no]
    
    strFormName = "F_9_メール送信_MAIN"
    Call next_form_open

End Sub

Private Sub btn_touroku_Click()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    Dim intSyanaiSansyoNo As Integer
    Dim intMiHacchuCount As Integer

    'If IsNull(me![pk_syanai_sansyo_no]) Then
    '    MsgBox ("社内参照番号を入力して、作成するORDERを抽出して下さい。")
    '    Exit Sub
    'End If

    If Me![chk_shinki_kizon] = 1 Then
        If IsNull(Me![pk_shiiresaki_code]) Or Len(Me![pk_shiiresaki_code]) <= 0 Then
            MsgBox ("発注先を選択してください。")
            GoTo Exit_btn_touroku_click
        End If
        If IsNull(Me![BL_DATE]) Or Len(Me![BL_DATE]) <= 0 Then
            MsgBox ("船積日を入力してください。")
            GoTo Exit_btn_touroku_click
        End If
        'If IsNull(Me![plant]) Or Len(Me![plant]) <= 0 Then
        '    MsgBox ("プラントを入力してください。")
        '    GoTo Exit_btn_touroku_click
        'End If
        'If IsNull(Me![hokan_basyo]) Or Len(Me![hokan_basyo]) <= 0 Then
        '    MsgBox ("保管場所を入力してください。")
        '    GoTo Exit_btn_touroku_click
        'End If
        If IsNull(Me![touroku_tanto]) Or Len(Me![touroku_tanto]) <= 0 Then
            Me![touroku_tanto] = Me![login_code]
        End If
    End If
    
    'BL DATEのチェック（新規の場合⇒処理日とチェック、登録済みの場合⇒登録日とチェック）
    If Not IsNull(Me![BL_DATE]) Then
        If (IsNull(Me![touroku_ymd]) And Me![BL_DATE] < Date) _
            Or (Not IsNull(Me![touroku_ymd]) And Me![BL_DATE] < Me![touroku_ymd]) Then
            ReturnValue = MsgBox("BL DATEが登録日より前の日付です" & vbLf & vbLf & "このまま登録しますか？", vbYesNo)
            If ReturnValue = vbNo Then
                Exit Sub
            End If
        End If
    End If

    Set cmd.ActiveConnection = conn
    
    '対象データがあるか？
    intMiHacchuCount = 0
    With cmd
        .CommandText = "usp_shikyu_hosozai_temp_count"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    If IsNull(cmd.Parameters(2)) Or cmd.Parameters(2) = 0 Then
        MsgBox ("入力データがありません。")
        GoTo Exit_btn_touroku_click
    End If
    
    '発注済かどうかの確認
    If IsNull(cmd.Parameters(3)) Or cmd.Parameters(3) = 0 Then
        ReturnValue = MsgBox("全て発注済のORDERです。" & vbLf & "更新しますか？。" & vbLf & "更新（はい）／キャンセル（いいえ）", vbYesNo)
        If ReturnValue = vbNo Then
            GoTo Exit_btn_touroku_click
        End If
    Else
        intMiHacchuCount = cmd.Parameters(3)
    End If
    
    '新規の場合、社内参照番号発番テーブルのチェック、発番
    intShinkiFlg = 0
    If Not IsNull(Me![pk_syanai_sansyo_no]) And Len(Me![pk_syanai_sansyo_no]) > 0 Then
        
        '登録済み支給包装材発注データの存在チェック
        With cmd
            .CommandText = "usp_shikyu_hosozai_sonzai_check"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![pk_syanai_sansyo_no]
            .Execute
        End With
        
        If IsNull(cmd.Parameters(2)) Or cmd.Parameters(2) = 0 Then
            If Me![chk_shinki_kizon] = 2 Then
                MsgBox ("登録済みに存在しない社内参照番号です。")
                GoTo Exit_btn_touroku_click
            End If
            intShinkiFlg = 1
        Else
            If Me![chk_shinki_kizon] = 1 Then
                MsgBox ("登録済みの社内参照番号です。")
                GoTo Exit_btn_touroku_click
            End If
            If intMiHacchuCount <> 0 Then
                '未発注の品目があった
                intShinkiFlg = 2
            End If
        End If
        
    Else
    
        intShinkiFlg = 1
        
    End If
    
    '入力内容のチェック
    intErrNo = 0
    With cmd
        .CommandText = "usp_shikyu_hosozai_err_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Execute
    End With
    If Not IsNull(cmd.Parameters(3)) And cmd.Parameters(3) <> 0 Then
        intErrNo = 2
    End If
    If Not IsNull(cmd.Parameters(2)) And cmd.Parameters(2) <> 0 Then
        intErrNo = 1
    End If
    Me![F_1_支給包装材発注_SUB1].Requery
        
    'メッセージの表示
    Select Case intErrNo
        Case 1
            MsgBox ("エラーデータがあります。")
            GoTo Exit_btn_touroku_click
        Case 2
            ReturnValue = MsgBox("WARNINGデータがあります、このま登録しますか？" & vbLf & "登録（はい）／確認（いいえ）", vbYesNo)
            If ReturnValue = vbNo Then
                GoTo Exit_btn_touroku_click
            End If
    End Select
    
    '新規の場合、社内参照番号発番
    If intShinkiFlg = 1 Then
    
        '連番の最大値取得
        With cmd
            .CommandText = "usp_syanai_sansyo_no_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = 2
            .Execute
        End With
        If Not IsNull(cmd.Parameters(2)) Then
            intSyanaiSansyoNo = cmd.Parameters(2) + 1
        End If
        
        Me![pk_syanai_sansyo_no] = "HSZ" & Format(intSyanaiSansyoNo, "000")
        Me![pk_syanai_sansyo_no].Requery
        
        '社内参照番号管理テーブル更新
        strSql = "INSERT INTO tbl_m_syanai_sansyo_no (shiiresaki_code, pk_syanai_sansyo_no, kigo, renban, hacchu_tanto, update_ymd, "
        strSql = strSql & "[remarks], syanai_sansyo_no_before, update_login ) VALUES ( '" & Me![pk_shiiresaki_code] & "', 'HSZ"
        strSql = strSql & Format(intSyanaiSansyoNo, "000") & "', " & "2, " & intSyanaiSansyoNo & ", '" & Me![touroku_tanto] & "', '"
        strSql = strSql & Now() & "', '" & "自動発番で使用" & "', " & "'HSZ" & Format(intSyanaiSansyoNo, "000") & "', '" & Me![login_code] & "' ) "
        'Debug.Print strSql
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
        
        strSql = "UPDATE tbl_t_shikyu_hosozai_hacchu_temp SET pk_syanai_sansyo_no = '" & Me![pk_syanai_sansyo_no] & "' "
        strSql = strSql & "FROM tbl_t_shikyu_hosozai_hacchu_temp "
        strSql = strSql & "WHERE tbl_t_shikyu_hosozai_hacchu_temp.pk_login_code = '" & Me![login_code] & "' "
        With cmd
            .CommandText = strSql
            .CommandType = adCmdText
            .Execute
        End With
        
    End If
    
    '支給包装材発注テーブル更新
    With cmd
        .CommandText = "usp_shikyu_hosozai_hacchu_update"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![login_code]
        .Parameters(2) = intShinkiFlg
        .Parameters(3) = Me![pk_syanai_sansyo_no]
        .Execute
    End With
    
    
    If intShinkiFlg = 0 Then
        
        '未発注が無かったら、分解後明細の再作成
        With cmd
            .CommandText = "usp_shikyu_hosozai_bunkai_meisai_create"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = Me![login_code]
            .Parameters(2) = Me![pk_syanai_sansyo_no]
            .Execute
        End With
        
    End If
    
    MsgBox ("登録完了しました。")
    
    Me![pk_syanai_sansyo_no].Requery
    Me![F_1_支給包装材発注_SUB1].Requery
    Me![F_1_支給包装材発注_SUB2].Requery
        
Exit_btn_touroku_click:
     
     Set rs1 = Nothing
     Set cmd = Nothing

End Sub

Private Sub btn_unmatch_Click()

    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_shikyu_hosozai_unmatch_temp_create"
        .CommandType = adCmdStoredProc
        .Execute
    End With
    
    Set cmd = Nothing
    
    DoCmd.OpenView "View_shikyu_hosozai_unmatch_check", acViewNormal
    
End Sub

Private Sub chk_shinki_kizon_Click()

    If Me![chk_shinki_kizon] = 1 Then
        Me![btn_shinki].ForeColor = "255"
        Me![btn_kizon].ForeColor = "13056"
        
        Call btn_clear_Click
        
        Me![pk_syanai_sansyo_no].Enabled = False
        Me![pk_syanai_sansyo_no].Locked = True
        Me![plant].Visible = False
        Me![hokan_basyo].Visible = False
        
    Else
        Me![btn_shinki].ForeColor = "13056"
        Me![btn_kizon].ForeColor = "255"
        
        Me![pk_syanai_sansyo_no].Enabled = True
        Me![pk_syanai_sansyo_no].Locked = False
        Me![plant].Visible = True
        Me![hokan_basyo].Visible = True
    End If
    
End Sub

Private Sub F_1_支給包装材発注_SUB1_Enter()

    If Forms![F_1_支給包装材発注_MAIN]![chk_shinki_kizon] = 1 Then
        If IsNull(Forms![F_1_支給包装材発注_MAIN]![pk_shiiresaki_code]) Then
            MsgBox ("仕入先コードを入力してください。")
        Else
            If IsNull(Forms![F_1_支給包装材発注_MAIN]![BL_DATE]) Then
                MsgBox ("BL DATEを入力してください。")
            Else
                'If IsNull(Forms![F_1_支給包装材発注_MAIN]![plant]) Then
                '    MsgBox ("plantを入力してください。")
                'Else
                '    If IsNull(Forms![F_1_支給包装材発注_MAIN]![hokan_basyo]) Then
                '        MsgBox ("保管場所を入力してください。")
                '    Else
                        If IsNull(Forms![F_1_支給包装材発注_MAIN]![touroku_tanto]) Then
                            MsgBox ("登録担当者を入力してください。")
                        End If
                '    End If
                'End If
            End If
        End If
    End If
        
    DoCmd.GoToRecord , , acFirst
    
End Sub

Private Sub btn_torikomi_Click()

    '購買発注取込処理へ
    strFormName = "F_1_R3購買発注取込"
    intMenuNo = 104
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    DoCmd.Maximize
    
    intMenuNo = 205
    intMenuNoSeq = 0
    
    Me![file_name] = "\\R3aws-data\tempdata\butsuryu\R3支給包装材発注.txt"
    
    Me![chk_shinki_kizon] = 1
    Me![btn_shinki].ForeColor = "255"
    Me![btn_kizon].ForeColor = "13056"
        
    Me![pk_syanai_sansyo_no].Enabled = False
    Me![pk_syanai_sansyo_no].Locked = True
    Me![plant].Visible = False
    Me![hokan_basyo].Visible = False
    
End Sub
Private Sub btn_close_Click()
    
    'スタートメニューへ
    strFormName = "F_0_START_MENU"
    intMenuNo = 0
    intMenuNoSeq = 0
    Call next_form_open
    
End Sub

Private Sub next_form_open()

    If IsNull(Me![login_code]) Then
        Call btn_setsuzoku_Click
    End If
    
    strInvoiceNo = "NULL"
    strOpenMode = "OC"
    
    Call form_open_close("F_1_支給包装材発注_MAIN", strFormName, strOpenMode, strInvoiceNo)
    
End Sub

Private Sub pk_syanai_sansyo_no_AfterUpdate()

    Me![hokan_basyo] = Null
    Me![koubai_hacchu_no] = Null
    Me![BL_DATE] = Null
    
End Sub
