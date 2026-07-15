Attribute VB_Name = "Form_F_M_CALENDER_詳細登録"
Attribute VB_Base = "0{0E3D0A88-A160-48C8-B637-C8FEE35F117E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit
    
    Dim intSyoriNo As Integer
    Dim strChkInvoiceNo As String

Private Sub btn_menu_Click()

    Dim intYY As Integer
    Dim intMM As Integer
    
    If Not IsNull(Me![yymmdd]) Then
        intYY = Year(Me![yymmdd])
        intMM = Month(Me![yymmdd])
    Else
        intYY = Year(Date)
        intMM = Month(Date)
    End If
    intCalenderPlantKbn = Me![frm_plant_kbn]
    
    DoCmd.Close acForm, "F_M_CALENDER_詳細登録"

    Call calender_display(intYY, intMM, intCalenderPlantKbn)
    
End Sub

Private Sub btn_setsuzoku_Click()

    Call Login_ReInput("F_M_CALENDER_詳細登録")

End Sub

Private Sub btn_touroku_Click()
    
    If (IsNull(Me![yymmdd]) Or Len(Me![yymmdd]) <= 0) _
        Or (IsNull(Me![yymmdd]) Or Len(Me![yymmdd]) <= 0) Then
        MsgBox ("日付、午前午後が未入力です。")
        Exit Sub
    End If
    If (IsNull(Me![frm_plant_kbn]) Or Me![frm_plant_kbn] = 0) Then
        MsgBox ("プラント（倉庫）を選択して下さい。")
        Exit Sub
    End If
    
    If (Len(Me![carton_qty1]) > 0 And Not IsNumeric(Me![carton_qty1])) _
        Or Len(Me![carton_qty2]) > 0 And Not IsNumeric(Me![carton_qty2]) Then
        MsgBox ("カートン数数字項目です。")
        Exit Sub
    End If
    
    If Me![frm_plant_kbn] = 2 Then
        strPlantCode = "_0103"
    ElseIf Me![frm_plant_kbn] = 3 Then
        strPlantCode = "_0410"
    Else
        strPlantCode = ""
    End If
    
    Set cmd.ActiveConnection = conn
    
    strSql = "UPDATE TJM_master.dbo.tbl_m_calender_jigyosyo SET container_" & Me![am_pm] & strPlantCode & " = "
    If IsNull(Me![container_name1]) Or Len(Me![container_name1]) <= 0 Then
        strSql = strSql & "null "
    Else
        strSql = strSql & "'" & Me![container_name1] & "' "
    End If
    
    strSql = strSql & ",container_size_" & Me![am_pm] & strPlantCode & " = "
    If IsNull(Me![container_size1]) Or Len(Me![container_size1]) <= 0 Then
        strSql = strSql & "null "
    Else
        strSql = strSql & "'" & Me![container_size1] & "' "
    End If
    
    strSql = strSql & ",invoice_no_" & Me![am_pm] & strPlantCode & " = "
    If IsNull(Me![invoice_no_main1]) Or Len(Me![invoice_no_main1]) <= 0 Then
        strSql = strSql & "null "
    Else
        strSql = strSql & "'" & Me![invoice_no_main1] & "' "
    End If
    
    strSql = strSql & ",carton_qty_" & Me![am_pm] & strPlantCode & " = "
    If IsNull(Me![carton_qty1]) Or Len(Me![carton_qty1]) <= 0 Then
        strSql = strSql & "null "
    Else
        strSql = strSql & Me![carton_qty1] & " "
    End If
    
    
    strSql = strSql & ",container_" & Me![am_pm] & "2" & strPlantCode & " = "
    If IsNull(Me![container_name2]) Or Len(Me![container_name2]) <= 0 Then
        strSql = strSql & "null "
    Else
        strSql = strSql & "'" & Me![container_name2] & "' "
    End If
    
    strSql = strSql & ",container_size_" & Me![am_pm] & "2" & strPlantCode & " = "
    If IsNull(Me![container_size2]) Or Len(Me![container_size2]) <= 0 Then
        strSql = strSql & "null "
    Else
        strSql = strSql & "'" & Me![container_size2] & "' "
    End If
    
    strSql = strSql & ",invoice_no_" & Me![am_pm] & "2" & strPlantCode & " = "
    If IsNull(Me![invoice_no_main2]) Or Len(Me![invoice_no_main2]) <= 0 Then
        strSql = strSql & "null "
    Else
        strSql = strSql & "'" & Me![invoice_no_main2] & "' "
    End If
    
    strSql = strSql & ",carton_qty_" & Me![am_pm] & "2" & strPlantCode & " = "
    If IsNull(Me![carton_qty2]) Or Len(Me![carton_qty2]) <= 0 Then
        strSql = strSql & "null "
    Else
        strSql = strSql & Me![carton_qty2] & " "
    End If
    
    strSql = strSql & ", update_ymd = '" & Now() & "', update_login = '" & Me![login_code] & "' "
    strSql = strSql & "WHERE pk_yymmdd = '" & Me![yymmdd] & "' "
    
    Debug.Print strSql
    With cmd
        .CommandText = strSql
        .CommandType = adCmdText
        .Execute
    End With
    
    Set cmd = Nothing
    
    MsgBox ("登録完了しました。")

End Sub

Private Sub invoice_no_main1_AfterUpdate()

    If IsNull(Me![invoice_no_main1]) Or Len(Me![invoice_no_main1]) <= 0 Then
        Exit Sub
    End If


    strChkInvoiceNo = Me![invoice_no_main1]
    intSyoriNo = 1
    
    Call invoice_header_get
    
End Sub

Private Sub invoice_no_main2_AfterUpdate()

    If IsNull(Me![invoice_no_main2]) Or Len(Me![invoice_no_main2]) <= 0 Then
        Exit Sub
    End If
    
    strChkInvoiceNo = Me![invoice_no_main2]
    intSyoriNo = 2
    
    Call invoice_header_get

End Sub

Private Sub invoice_header_get()


    Set cmd.ActiveConnection = conn
    
    If Left(strChkInvoiceNo, 3) = "SJ-" Or Left(strChkInvoiceNo, 3) = "YJ-" Then
    
        '上海のINVOICEからヘッダ情報を抽出
        Me("container_name" & intSyoriNo) = Null
        Me("container_size" & intSyoriNo) = Null
        Me("carton_qty" & intSyoriNo) = Null
    
        With cmd
            .CommandText = "usp_shanghai_invoice_gokei_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = strChkInvoiceNo
            .Execute
        End With
        
        If Not IsNull(cmd.Parameters(4)) Then
            If Left(strChkInvoiceNo, 3) = "SJ-" Then
                Me("container_name" & intSyoriNo) = "上海田島"
            Else
                Me("container_name" & intSyoriNo) = "上海庸助貿易"
            End If
            Me("container_size" & intSyoriNo) = cmd.Parameters(7)
            Me("carton_qty" & intSyoriNo) = cmd.Parameters(4)
        End If
        
    Else
    
        '調達のINVOICEからヘッダ情報を抽出
        With cmd
            .CommandText = "usp_chotatsu_invoice_header_get"
            .CommandType = adCmdStoredProc
            .Parameters.Refresh
            .Parameters(1) = strChkInvoiceNo
            .Execute
        End With
        If Not IsNull(cmd.Parameters(3)) Then
            Me("carton_qty" & intSyoriNo) = cmd.Parameters(3)
        End If
        If Not IsNull(cmd.Parameters(4)) Then
            Me("container_name" & intSyoriNo) = cmd.Parameters(4)
        End If
        
    End If
    
    Set cmd = Nothing
    
    Me.Requery

End Sub
