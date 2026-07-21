Namespace Forms.Expense
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class ExpenseCodeMaintenanceForm
        Inherits Atom.App.Common.BaseForm

        <System.Diagnostics.DebuggerNonUserCode()> _
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            Try
                If disposing AndAlso components IsNot Nothing Then
                    components.Dispose()
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        Private components As System.ComponentModel.IContainer

        <System.Diagnostics.DebuggerStepThrough()> _
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Me.barTools = New System.Windows.Forms.FlowLayoutPanel()
            Me.btnAdd = New System.Windows.Forms.Button()
            Me.btnDelete = New System.Windows.Forms.Button()
            Me.btnSave = New System.Windows.Forms.Button()
            Me.dgvList = New System.Windows.Forms.DataGridView()
            Me.bsExpense = New System.Windows.Forms.BindingSource(Me.components)
            Me.colCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colText = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colKbn = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.colRef = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colYmd = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.colUsr = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.barTools.SuspendLayout()
            CType(Me.dgvList, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.bsExpense, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'barTools
            '
            Me.barTools.Controls.Add(Me.btnAdd)
            Me.barTools.Controls.Add(Me.btnDelete)
            Me.barTools.Controls.Add(Me.btnSave)
            Me.barTools.Dock = System.Windows.Forms.DockStyle.Top
            Me.barTools.Name = "barTools"
            Me.barTools.Padding = New System.Windows.Forms.Padding(6, 6, 6, 4)
            Me.barTools.Size = New System.Drawing.Size(984, 40)
            Me.barTools.TabIndex = 0
            '
            'btnAdd
            '
            Me.btnAdd.Margin = New System.Windows.Forms.Padding(0, 0, 6, 0)
            Me.btnAdd.Name = "btnAdd"
            Me.btnAdd.Size = New System.Drawing.Size(100, 26)
            Me.btnAdd.TabIndex = 0
            Me.btnAdd.Text = "新規行追加"
            Me.btnAdd.UseVisualStyleBackColor = True
            '
            'btnDelete
            '
            Me.btnDelete.Margin = New System.Windows.Forms.Padding(0, 0, 6, 0)
            Me.btnDelete.Name = "btnDelete"
            Me.btnDelete.Size = New System.Drawing.Size(90, 26)
            Me.btnDelete.TabIndex = 1
            Me.btnDelete.Text = "行削除"
            Me.btnDelete.UseVisualStyleBackColor = True
            '
            'btnSave
            '
            Me.btnSave.Name = "btnSave"
            Me.btnSave.Size = New System.Drawing.Size(90, 26)
            Me.btnSave.TabIndex = 2
            Me.btnSave.Text = "保存"
            Me.btnSave.UseVisualStyleBackColor = True
            '
            'dgvList
            '
            Me.dgvList.AllowUserToAddRows = True
            Me.dgvList.AllowUserToDeleteRows = True
            Me.dgvList.AutoGenerateColumns = False
            Me.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.dgvList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colCode, Me.colText, Me.colKbn, Me.colRef, Me.colYmd, Me.colUsr})
            Me.dgvList.DataSource = Me.bsExpense
            Me.dgvList.Dock = System.Windows.Forms.DockStyle.Fill
            Me.dgvList.Name = "dgvList"
            Me.dgvList.RowHeadersWidth = 24
            Me.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            Me.dgvList.Size = New System.Drawing.Size(984, 491)
            Me.dgvList.TabIndex = 1
            '
            'colCode
            '
            Me.colCode.DataPropertyName = "pk_keihi_code"
            Me.colCode.HeaderText = "経費コード"
            Me.colCode.Name = "colCode"
            Me.colCode.Width = 120
            '
            'colText
            '
            Me.colText.DataPropertyName = "keihi_code_text"
            Me.colText.HeaderText = "経費コード名称"
            Me.colText.Name = "colText"
            Me.colText.Width = 220
            '
            'colKbn
            '
            Me.colKbn.DataPropertyName = "keihi_kbn"
            Me.colKbn.HeaderText = "経費区分"
            Me.colKbn.Name = "colKbn"
            Me.colKbn.Width = 110
            '
            'colRef
            '
            Me.colRef.DataPropertyName = "keihi_code_ref"
            Me.colRef.HeaderText = "参照コード"
            Me.colRef.Name = "colRef"
            Me.colRef.Width = 100
            '
            'colYmd
            '
            Me.colYmd.DataPropertyName = "update_ymd"
            Me.colYmd.HeaderText = "更新日時"
            Me.colYmd.Name = "colYmd"
            Me.colYmd.[ReadOnly] = True
            Me.colYmd.Width = 150
            '
            'colUsr
            '
            Me.colUsr.DataPropertyName = "update_login"
            Me.colUsr.HeaderText = "更新者"
            Me.colUsr.Name = "colUsr"
            Me.colUsr.[ReadOnly] = True
            Me.colUsr.Width = 90
            '
            'ExpenseCodeMaintenanceForm
            '
            Me.ContentPanel.Controls.Add(Me.dgvList)
            Me.ContentPanel.Controls.Add(Me.barTools)
            Me.Name = "ExpenseCodeMaintenanceForm"
            Me.barTools.ResumeLayout(False)
            CType(Me.dgvList, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.bsExpense, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
        End Sub

        Friend WithEvents barTools As System.Windows.Forms.FlowLayoutPanel
        Friend WithEvents btnAdd As System.Windows.Forms.Button
        Friend WithEvents btnDelete As System.Windows.Forms.Button
        Friend WithEvents btnSave As System.Windows.Forms.Button
        Friend WithEvents dgvList As System.Windows.Forms.DataGridView
        Friend WithEvents bsExpense As System.Windows.Forms.BindingSource
        Friend WithEvents colCode As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents colText As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents colKbn As System.Windows.Forms.DataGridViewComboBoxColumn
        Friend WithEvents colRef As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents colYmd As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents colUsr As System.Windows.Forms.DataGridViewTextBoxColumn
    End Class
End Namespace
