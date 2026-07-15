Attribute VB_Name = "Form_F_1_SAP_AMOUNT_SUB1"
Attribute VB_Base = "0{007AF8C2-D695-4C91-90E5-93343A2C2499}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_select_Click()

    If IsNull(Me![koubai_hacchu_no]) Then
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_kaigai_shiire_jisseki_err_check"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Forms![F_1_SAP_AMOUNT_MAIN]![login_code]
        .Parameters(2) = Me![koubai_hacchu_no]
        .Execute
    End With
    
    If IsNull(cmd.Parameters(3)) Or cmd.Parameters(3) = 0 Then
        MsgBox ("対象データはありませんでした。")
    Else

        '確認フォームへ
        strFormName = "F_1_SAP_AMOUNT_CHECK_MAIN"
        intMenuNo = 107
        intMenuNoSeq = 0
    
        strInvoiceNo = ""
        strOpenMode = "O"
        
        Call form_open_close("F_1_SAP_AMOUNT_MAIN", strFormName, strOpenMode, "NULL")

    End If
    
    Set cmd = Nothing
    
End Sub

Private Sub pk_syanai_sansyo_no_AfterUpdate()
    
    Set cmd.ActiveConnection = conn
    
    strSql = "SELECT tbl_t_shanghai_invoice_meisai.syanai_sansyo_no, tbl_t_shanghai_invoice_header.BL_DATE, "
    strSql = strSql & "tbl_t_shanghai_invoice_meisai.pk_invoice_no, Sum(tbl_t_shanghai_invoice_meisai.AMOUNT) AS INVOICE_AMOUNT "
    strSql = strSql & "FROM tbl_t_shanghai_invoice_header INNER JOIN tbl_t_shanghai_invoice_meisai ON "
    strSql = strSql & "tbl_t_shanghai_invoice_header.pk_invoice_no_main = tbl_t_shanghai_invoice_meisai.pk_invoice_no_main "
    strSql = strSql & "GROUP BY tbl_t_shanghai_invoice_meisai.syanai_sansyo_no, tbl_t_shanghai_invoice_meisai.pk_invoice_no, "
    strSql = strSql & "tbl_t_shanghai_invoice_header.BL_DATE HAVING tbl_t_shanghai_invoice_meisai.syanai_sansyo_no = '" & Me![pk_syanai_sansyo_no] & "' "

    Set rs1 = cmd.Execute(intRcnt)
    
    If intRcnt <> 0 Then
        Me![INVOICE_AMOUNT] = rs1![INVOICE_AMOUNT].Value
        Me![BL_DATE] = rs1![BL_DATE].Value
        Me![pk_yyyymm] = Year(rs1![BL_DATE].Value) * 100 + Month(rs1![BL_DATE].Value)
        Me![pk_invoice_no] = rs1![pk_invoice_no].Value
    Else
        Me![INVOICE_AMOUNT] = 0
    End If
    
    rs1.Close
    Set cmd = Nothing


End Sub

Private Sub koubai_hacchu_no_AfterUpdate()

    If Len(Me![koubai_hacchu_no]) <> 10 Or (IsNumeric(Me![koubai_hacchu_no]) And Left(Me![koubai_hacchu_no], 2) <> "45") Then
        MsgBox ("購買発注NOが不正です。")
        Exit Sub
    End If

End Sub
