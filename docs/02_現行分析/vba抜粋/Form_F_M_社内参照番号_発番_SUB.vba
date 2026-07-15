Attribute VB_Name = "Form_F_M_社内参照番号_発番_SUB"
Attribute VB_Base = "0{3376E958-AB0B-4EB1-87AE-A16AF97FD865}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub pk_syanai_sansyo_no_AfterUpdate()

    If Not IsNull(Me![pk_syanai_sansyo_no]) Then
        If Left(Me![pk_syanai_sansyo_no], 2) = "KH" Then
            Me![kigo] = 1
            If IsNumeric(Mid(Me![pk_syanai_sansyo_no], 3, 4)) Then
                Me![renban] = Mid(Me![pk_syanai_sansyo_no], 3, 4)
                Me![syanai_sansyo_no_before] = Me![pk_syanai_sansyo_no]
            Else
                If Not IsNumeric(Mid(Me![pk_syanai_sansyo_no], 3, 3)) Then
                    MsgBox ("社内参照番号の数字部分は３桁です。")
                    '入力前の番号に戻す
                    Me![pk_syanai_sansyo_no] = Me![syanai_sansyo_no_before]
                Else
                    Me![renban] = Mid(Me![pk_syanai_sansyo_no], 3, 3)
                    Me![syanai_sansyo_no_before] = Me![pk_syanai_sansyo_no]
                End If
            End If
        Else
            If Left(Me![pk_syanai_sansyo_no], 3) = "HSZ" Then
                Me![kigo] = 2
                If Not IsNumeric(Mid(Me![pk_syanai_sansyo_no], 4, 3)) Then
                    MsgBox ("社内参照番号の数字部分は３桁です。")
                    '入力前の番号に戻す
                    Me![pk_syanai_sansyo_no] = Me![syanai_sansyo_no_before]
                Else
                    Me![renban] = Mid(Me![pk_syanai_sansyo_no], 4, 3)
                    Me![syanai_sansyo_no_before] = Me![pk_syanai_sansyo_no]
                End If
            Else
                MsgBox ("社内参照番号は、『KH』、または『HSZ』で始まる番号を入力してください。")
                '入力前の番号に戻す
                Me![pk_syanai_sansyo_no] = Me![syanai_sansyo_no_before]
            End If
        End If
    End If

    Me![update_ymd] = Now()
    
End Sub

Private Sub hacchu_tanto_AfterUpdate()

    Me![update_ymd] = Now()
    
End Sub

Private Sub remarks_AfterUpdate()

    Me![update_ymd] = Now()
    
End Sub
