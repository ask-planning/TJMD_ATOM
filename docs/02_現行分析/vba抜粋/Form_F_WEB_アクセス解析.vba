Attribute VB_Name = "Form_F_WEB_アクセス解析"
Attribute VB_Base = "0{384361AA-62B6-4570-85D1-DFCD65D5BCE5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit


Private Sub btn_clear_Click()

        Me![select_menu_no] = Null
        Me![select_ip] = Null
        Me![start_ymd] = Null
        Me![end_ymd] = Null
        
End Sub

Private Sub btn_menu_Click()

    DoCmd.Close
    
End Sub

Private Sub btn_pickup_Click()

    If Me![end_ymd] < Me![start_ymd] Then
        MsgBox ("日付入力エラー")
        Exit Sub
    End If

    Dim strAnd As String
    Dim strFromKu As String
    Dim strWhereKu As String
    Dim strGroupByKu As String
    Dim strOrderByKu As String
    
    Dim intMenuNo As Integer
    Dim strJigyosyoIp As String
    Dim intHyojiKbn As Integer
    Dim datStartYMD As Date
    Dim datEndYMD As Date
    Dim intPickupCount As Integer
    
    
    Dim strsql2 As String
    
    Set cmd.ActiveConnection = conn
    
    intMenuNo = 0
    intHyojiKbn = 0
    datStartYMD = 1900 / 1 / 1
    datEndYMD = 1900 / 1 / 1
    strJigyosyoIp = ""
    If Not IsNull(Me![select_menu_no]) Then
        intMenuNo = Me![select_menu_no]
    End If
    If Not IsNull(Me![select_ip]) Then
        strJigyosyoIp = Me![select_ip]
    End If
    If Not IsNull(Me![start_ymd]) Then
        datStartYMD = Me![start_ymd]
    End If
    If Not IsNull(Me![end_ymd]) Then
        datEndYMD = Me![end_ymd]
    End If
    If IsNull(Me![frm_hyoji_kbn]) Then
        intHyojiKbn = 1
    Else
        intHyojiKbn = Me![frm_hyoji_kbn]
    End If
    
    intPickupCount = 0
    
    strSql = ""
    strsql2 = ""
    
    
    strsql2 = "SELECT Count([View_web_access].menu_no) AS pickup_count "
    Select Case intHyojiKbn
        Case 1
            strSql = "SELECT [View_web_access].menu_no, [View_web_access].pk_menu, [View_web_access].menu_text as koumoku1, null as koumoku2, "
            strSql = strSql & "CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN CASE WHEN tbl_m_jigyosyo_ip2.kyoten_ip is null THEN '不明' "
            strSql = strSql & "ELSE tbl_m_jigyosyo_ip2.kyoten_name END ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END as koumoku5, "
            strSql = strSql & "Count([View_web_access].eigyo_ip) AS koumoku3, null as koumoku4, null as koumoku6, null as koumoku7 "
        Case 2
            strSql = "SELECT [View_web_access].menu_no, [View_web_access].pk_menu, [View_web_access].menu_text as koumoku1, "
            strSql = strSql & "CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN CASE WHEN tbl_m_jigyosyo_ip2.kyoten_ip is null THEN '不明' "
            strSql = strSql & "ELSE tbl_m_jigyosyo_ip2.kyoten_name END ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END as koumoku5, "
            'strsql = strsql & "CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN '不明' ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END as koumoku5, "
            strSql = strSql & "[View_web_access].eigyo_ip , null as koumoku3, CASE WHEN [View_web_access].login_code is null OR "
            strSql = strSql & "[View_web_access].login_code = '999999' THEN hinmoku_v3_sql.dbo.tbl_t_remote_ip.IP_user_name ELSE TJM_master.dbo.tbl_m_login_sys.shain_name END "
            strSql = strSql & "as koumoku6, [View_web_access].remarks as koumoku7, "
            strSql = strSql & "[View_web_access].accessymd as koumoku4 , [View_web_access].remoteaddr as koumoku2 "
        Case 3
            strSql = "SELECT [View_web_access].menu_no, [View_web_access].pk_menu, [View_web_access].menu_text as koumoku1, null as koumoku2, "
            strSql = strSql & "null as koumoku5, Count([View_web_access].menu_no) AS koumoku3, null as koumoku4, null as koumoku6, null as koumoku7 "
        Case 4
            strSql = "SELECT [View_web_access].menu_no, [View_web_access].pk_menu, [View_web_access].menu_text as koumoku1, "
            'strsql = strsql & "CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN '不明' ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END as koumoku5, "
            strSql = strSql & "CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN CASE WHEN tbl_m_jigyosyo_ip2.kyoten_ip is null THEN '不明' "
            strSql = strSql & "ELSE tbl_m_jigyosyo_ip2.kyoten_name END ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END as koumoku5, "
            strSql = strSql & "[View_web_access].eigyo_ip , Count([View_web_access].remoteaddr) as koumoku3, "
            strSql = strSql & "null as koumoku4 , [View_web_access].remoteaddr as koumoku2, hinmoku_v3_sql.dbo.tbl_t_remote_ip.IP_user_name as koumoku6, hinmoku_v3_sql.dbo.tbl_t_remote_ip.IP_busyo as koumoku7 "
        Case 5
            strSql = "SELECT [View_web_access].menu_no, [View_web_access].pk_menu, [View_web_access].menu_text as koumoku1, null as koumoku2, "
            'strsql = strsql & "CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN '不明' ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END as koumoku5, "
            strSql = strSql & "CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN CASE WHEN tbl_m_jigyosyo_ip2.kyoten_ip is null THEN '不明' "
            strSql = strSql & "ELSE tbl_m_jigyosyo_ip2.kyoten_name END ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END as koumoku5, "
            strSql = strSql & "Count([View_web_access].eigyo_ip) AS koumoku3, null as koumoku4, null as koumoku6, hinmoku_v3_sql.dbo.tbl_t_remote_ip.IP_busyo as koumoku7 "
        Case Else
            strSql = "SELECT [View_web_access].menu_no, [View_web_access].pk_menu, [View_web_access].menu_text as koumoku1, null as koumoku2, "
            'strsql = strsql & "CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN '不明' ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END as koumoku5, "
            strSql = strSql & "CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN CASE WHEN tbl_m_jigyosyo_ip2.kyoten_ip is null THEN '不明' "
            strSql = strSql & "ELSE tbl_m_jigyosyo_ip2.kyoten_name END ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END as koumoku5, "
            strSql = strSql & "Count([View_web_access].eigyo_ip) AS koumoku3, null as koumoku4, null as koumoku6, null as koumoku7 "
    End Select
    
    strFromKu = "FROM [View_web_access] LEFT JOIN TJM_master.dbo.tbl_m_jigyosyo_ip ON [View_web_access].eigyo_ip = TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip "
    strFromKu = strFromKu & "LEFT JOIN TJM_master.dbo.tbl_m_jigyosyo_ip as tbl_m_jigyosyo_ip2 ON [View_web_access].eigyo_ip2 = tbl_m_jigyosyo_ip2.kyoten_ip "
    strFromKu = strFromKu & "LEFT JOIN hinmoku_v3_sql.dbo.tbl_t_remote_ip ON [View_web_access].[remoteaddr] = hinmoku_v3_sql.dbo.tbl_t_remote_ip.IP_address "
    strFromKu = strFromKu & "LEFT JOIN TJM_master.dbo.tbl_m_login_sys ON [View_web_access].[login_code] = TJM_master.dbo.tbl_m_login_sys.pk_shain_code "

    strAnd = " WHERE "
    strWhereKu = ""
    If Not IsNull(Me![select_menu_no]) Then
        strWhereKu = strAnd & "[View_web_access].menu_no = " & Me![select_menu_no] & " "
        strAnd = " AND "
    End If
    If Not IsNull(Me![select_ip]) Then
        strWhereKu = strWhereKu & strAnd & "TJM_master.dbo.tbl_m_jigyosyo_ip.[kyoten_ip] = '" & Me![select_ip] & "' "
        strAnd = " AND "
    End If
    If Not IsNull(Me![start_ymd]) Then
        strWhereKu = strWhereKu & strAnd & "[View_web_access].accessymd between '" & datStartYMD & "' AND '" & datEndYMD + 1 & "' "
        strAnd = " AND "
    End If
    
    strGroupByKu = ""
    If intHyojiKbn = 1 Or intHyojiKbn = 3 Or intHyojiKbn = 4 Or intHyojiKbn = 5 Then
        strGroupByKu = "GROUP BY [View_web_access].pk_menu,[View_web_access].menu_no,[View_web_access].menu_text,[View_web_access].menu_no "
        If intHyojiKbn = 1 Or intHyojiKbn = 4 Or intHyojiKbn = 5 Then
            'strGroupByKu = strGroupByKu & ", CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN '不明' ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END "
            strGroupByKu = strGroupByKu & ", CASE WHEN TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_ip IS NULL THEN CASE WHEN tbl_m_jigyosyo_ip2.kyoten_ip is null THEN '不明' "
            strGroupByKu = strGroupByKu & "ELSE tbl_m_jigyosyo_ip2.kyoten_name END ELSE TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name END "
        End If
        If intHyojiKbn = 4 Then
            strGroupByKu = strGroupByKu & ",TJM_master.dbo.tbl_m_jigyosyo_ip.kyoten_name, [View_web_access].eigyo_ip, [View_web_access].remoteaddr "
            strGroupByKu = strGroupByKu & ", hinmoku_v3_sql.dbo.tbl_t_remote_ip.IP_user_name, hinmoku_v3_sql.dbo.tbl_t_remote_ip.IP_busyo "
        End If
        If intHyojiKbn = 5 Then
            strGroupByKu = strGroupByKu & ", hinmoku_v3_sql.dbo.tbl_t_remote_ip.IP_busyo "
        End If
    End If
    
    strOrderByKu = "ORDER BY [View_web_access].menu_no"
    If intHyojiKbn = 2 Then
        strOrderByKu = strOrderByKu & ",[View_web_access].accessymd desc "
    End If
    
    strSql = strSql & strFromKu & strWhereKu & strGroupByKu & strOrderByKu
    strsql2 = strsql2 & strFromKu & strWhereKu & strGroupByKu
    
    Debug.Print strSql
    
    With cmd
        .CommandText = strsql2
        .CommandType = adCmdText
    End With
    
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt <> 0 Then
        rs1.MoveFirst
        If intHyojiKbn <> 2 Then
            Do Until rs1.EOF
                intPickupCount = intPickupCount + 1
                rs1.MoveNext
            Loop
        Else
            intPickupCount = rs1![pickup_count].Value
        End If
    End If
    rs1.Close
    
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
    End With
    
    '実行結果が０の時、「ありませんでした。」
    Set rs1 = cmd.Execute(intRcnt)
    If intRcnt = 0 Then
        MsgBox ("該当データはありませんでした。")
        '空データの表示
        strSql = "SELECT 0 as menu_no, '該当データ無し' as koumoku1,null as koumoku2,remoteuser as koumoku3, null as koumoku5, "
        strSql = strSql & "null as koumoku6 , null as koumoku7 FROM tbl_t_web_access Where pk_menu = '該当データ無し'"
        'GoTo Exit_kokyaku_phone_no_AfterUpdate
    End If
       
    rs1.Close
    
    'レコードセットの再オープン
    rs1.CursorLocation = adUseClient
    rs1.Open strSql
    
    '画面フィールドのコントロールソースを外す
    Me.RecordSource = ""
    With Me
        ![menu_no].ControlSource = ""
        ![koumoku1].ControlSource = ""
        ![koumoku2].ControlSource = ""
        ![koumoku3].ControlSource = ""
        ![koumoku4].ControlSource = ""
        ![koumoku5].ControlSource = ""
        ![koumoku6].ControlSource = ""
        ![koumoku7].ControlSource = ""
    End With
    
    'フォームオープン
    DoCmd.OpenForm "F_WEB_アクセス解析"
    
    'フォームのレコードセットを読み込んだレコードセットにする
    Set Forms![F_WEB_アクセス解析].Recordset = rs1
    
    'レコードセットのフィールドを、フォームのコントロールソースにセット、画面表示
    Me![menu_no].ControlSource = "menu_no"
    Me![koumoku1].ControlSource = "koumoku1"
    Me![koumoku2].ControlSource = "koumoku2"
    Me![koumoku3].ControlSource = "koumoku3"
    Me![koumoku4].ControlSource = "koumoku4"
    Me![koumoku5].ControlSource = "koumoku5"
    Me![koumoku6].ControlSource = "koumoku6"
    Me![koumoku7].ControlSource = "koumoku7"
    If intHyojiKbn = 1 Or intHyojiKbn = 3 Or intHyojiKbn = 4 Or intHyojiKbn = 5 Then
        Me![lbl_koumoku3].Caption = "アクセス数"
        Me![koumoku3].Visible = True
        Me![koumoku4].Visible = False
    Else
        Me![lbl_koumoku3].Caption = "アクセス日付"
        Me![koumoku3].Visible = False
        Me![koumoku4].Visible = True
    End If
    If intHyojiKbn = 2 Then
        Me![lbl_koumoku7].Caption = "REMARKS"
    Else
        Me![lbl_koumoku7].Caption = "部署"
    End If
    
    If intMenuNo <> 0 Then
        Me![select_menu_no] = intMenuNo
    End If
    If Len(strJigyosyoIp) > 0 Then
        Me![select_ip] = strJigyosyoIp
    End If
    If datStartYMD <> 1900 / 1 / 1 Then
        Me![start_ymd] = datStartYMD
    End If
    If datEndYMD <> 1900 / 1 / 1 Then
        Me![end_ymd] = datEndYMD
    End If
    Me![frm_hyoji_kbn] = intHyojiKbn
    Me![pickup_count] = intPickupCount
    
    Me![start_ymd].Requery
    Me![end_ymd].Requery
    Me![select_menu_no].Requery
    Me![select_ip].Requery
    
    Set rs1 = Nothing
    Set cmd = Nothing
    
End Sub

Private Sub btn_setsuzoku_Click()

    '再接続

    ReturnValue = make_connect()
    

End Sub

Private Sub Form_Open(Cancel As Integer)

    '再接続

    ReturnValue = make_connect()
    
    DoCmd.Maximize
    Me![start_ymd] = Date
    Me![end_ymd] = Date

End Sub

Private Sub start_ymd_AfterUpdate()

    If IsNull(Me![start_ymd]) Then
        Me![end_ymd] = Null
    End If

    If IsNull(Me![end_ymd]) Then
        Me![end_ymd] = Me![start_ymd]
    End If
End Sub
