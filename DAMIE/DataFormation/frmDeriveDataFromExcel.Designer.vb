﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDeriveDataFromExcel
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDeriveDataFromExcel))
        Me.btnExport = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.BtnChoosePath = New System.Windows.Forms.Button()
        Me.ListBoxWorksheets = New System.Windows.Forms.ListBox()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtbxSavePath = New System.Windows.Forms.TextBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.lbSheetName = New System.Windows.Forms.Label()
        Me.AddFileOrDirectoryFiles1 = New DAMIE.AME_UserControl.AddFileOrDirectoryFiles()
        Me.MyDataGridView1 = New DAMIE.AME_UserControl.myDataGridView()
        Me.SheetName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RangeName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Txtbox_DateFormat = New System.Windows.Forms.TextBox()
        Me.btn_DateFormat = New System.Windows.Forms.Button()
        Me.ChkboxParseDate = New System.Windows.Forms.CheckBox()
        Me.ChkBxOpenExcelWhileFinished = New System.Windows.Forms.CheckBox()
        CType(Me.MyDataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(452, 351)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(75, 23)
        Me.btnExport.TabIndex = 2
        Me.btnExport.Text = "输出"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'BtnChoosePath
        '
        Me.BtnChoosePath.BackColor = System.Drawing.SystemColors.Control
        Me.BtnChoosePath.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnChoosePath.ForeColor = System.Drawing.SystemColors.InfoText
        Me.BtnChoosePath.Location = New System.Drawing.Point(453, 319)
        Me.BtnChoosePath.Name = "BtnChoosePath"
        Me.BtnChoosePath.Size = New System.Drawing.Size(74, 23)
        Me.BtnChoosePath.TabIndex = 3
        Me.BtnChoosePath.Text = "选择..."
        Me.BtnChoosePath.UseVisualStyleBackColor = False
        '
        'ListBoxWorksheets
        '
        Me.ListBoxWorksheets.AllowDrop = True
        Me.ListBoxWorksheets.FormattingEnabled = True
        Me.ListBoxWorksheets.HorizontalScrollbar = True
        Me.ListBoxWorksheets.ItemHeight = 12
        Me.ListBoxWorksheets.Location = New System.Drawing.Point(13, 38)
        Me.ListBoxWorksheets.Name = "ListBoxWorksheets"
        Me.ListBoxWorksheets.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxWorksheets.Size = New System.Drawing.Size(407, 136)
        Me.ListBoxWorksheets.TabIndex = 6
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(426, 114)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(100, 24)
        Me.btnRemove.TabIndex = 6
        Me.btnRemove.Text = "移除"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(155, 12)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "进行数据提取的Excel工作簿"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 303)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 12)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "保存至"
        '
        'txtbxSavePath
        '
        Me.txtbxSavePath.BackColor = System.Drawing.Color.White
        Me.txtbxSavePath.Location = New System.Drawing.Point(11, 321)
        Me.txtbxSavePath.Margin = New System.Windows.Forms.Padding(0)
        Me.txtbxSavePath.Name = "txtbxSavePath"
        Me.txtbxSavePath.Size = New System.Drawing.Size(427, 21)
        Me.txtbxSavePath.TabIndex = 1
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.BackColor = System.Drawing.SystemColors.Control
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 384)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(539, 8)
        Me.ProgressBar1.TabIndex = 9
        '
        'BackgroundWorker1
        '
        '
        'lbSheetName
        '
        Me.lbSheetName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbSheetName.Location = New System.Drawing.Point(11, 356)
        Me.lbSheetName.Name = "lbSheetName"
        Me.lbSheetName.Size = New System.Drawing.Size(427, 25)
        Me.lbSheetName.TabIndex = 10
        Me.lbSheetName.Text = "."
        '
        'AddFileOrDirectoryFiles1
        '
        Me.AddFileOrDirectoryFiles1.BackColor = System.Drawing.Color.Transparent
        Me.AddFileOrDirectoryFiles1.Location = New System.Drawing.Point(427, 38)
        Me.AddFileOrDirectoryFiles1.Margin = New System.Windows.Forms.Padding(0)
        Me.AddFileOrDirectoryFiles1.Name = "AddFileOrDirectoryFiles1"
        Me.AddFileOrDirectoryFiles1.Size = New System.Drawing.Size(100, 68)
        Me.AddFileOrDirectoryFiles1.TabIndex = 17
        '
        'MyDataGridView1
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("SimSun", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.MyDataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.MyDataGridView1.ColumnHeadersHeight = 25
        Me.MyDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.MyDataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SheetName, Me.RangeName})
        Me.MyDataGridView1.Location = New System.Drawing.Point(11, 181)
        Me.MyDataGridView1.Name = "MyDataGridView1"
        Me.MyDataGridView1.RowTemplate.Height = 23
        Me.MyDataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.MyDataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.MyDataGridView1.Size = New System.Drawing.Size(346, 110)
        Me.MyDataGridView1.TabIndex = 14
        '
        'SheetName
        '
        Me.SheetName.HeaderText = "工作表名称"
        Me.SheetName.Name = "SheetName"
        Me.SheetName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.SheetName.ToolTipText = "提取的工作表名称包含于要进行检索的工作表名称，比如输入""CX""，则会提取工作簿中第一个名称中含有""CX""的工作表。"
        Me.SheetName.Width = 183
        '
        'RangeName
        '
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Blue
        Me.RangeName.DefaultCellStyle = DataGridViewCellStyle2
        Me.RangeName.HeaderText = "区域范围"
        Me.RangeName.Name = "RangeName"
        Me.RangeName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.RangeName.ToolTipText = "示例： A1:C3"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Txtbox_DateFormat)
        Me.Panel1.Controls.Add(Me.btn_DateFormat)
        Me.Panel1.Controls.Add(Me.ChkboxParseDate)
        Me.Panel1.Location = New System.Drawing.Point(366, 181)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(161, 88)
        Me.Panel1.TabIndex = 18
        '
        'Txtbox_DateFormat
        '
        Me.Txtbox_DateFormat.Enabled = False
        Me.Txtbox_DateFormat.Location = New System.Drawing.Point(13, 33)
        Me.Txtbox_DateFormat.Name = "Txtbox_DateFormat"
        Me.Txtbox_DateFormat.Size = New System.Drawing.Size(143, 21)
        Me.Txtbox_DateFormat.TabIndex = 3
        '
        'btn_DateFormat
        '
        Me.btn_DateFormat.Enabled = False
        Me.btn_DateFormat.Location = New System.Drawing.Point(13, 60)
        Me.btn_DateFormat.Name = "btn_DateFormat"
        Me.btn_DateFormat.Size = New System.Drawing.Size(75, 23)
        Me.btn_DateFormat.TabIndex = 2
        Me.btn_DateFormat.Text = "日期格式"
        Me.btn_DateFormat.UseVisualStyleBackColor = True
        '
        'ChkboxParseDate
        '
        Me.ChkboxParseDate.AutoSize = True
        Me.ChkboxParseDate.Location = New System.Drawing.Point(3, 3)
        Me.ChkboxParseDate.Name = "ChkboxParseDate"
        Me.ChkboxParseDate.Size = New System.Drawing.Size(132, 16)
        Me.ChkboxParseDate.TabIndex = 0
        Me.ChkboxParseDate.Text = "提取文件名中的日期"
        Me.ChkboxParseDate.UseVisualStyleBackColor = True
        '
        'ChkBxOpenExcelWhileFinished
        '
        Me.ChkBxOpenExcelWhileFinished.AutoSize = True
        Me.ChkBxOpenExcelWhileFinished.Location = New System.Drawing.Point(366, 275)
        Me.ChkBxOpenExcelWhileFinished.Name = "ChkBxOpenExcelWhileFinished"
        Me.ChkBxOpenExcelWhileFinished.Size = New System.Drawing.Size(138, 16)
        Me.ChkBxOpenExcelWhileFinished.TabIndex = 22
        Me.ChkBxOpenExcelWhileFinished.Text = "操作完成后打开Excel"
        Me.ChkBxOpenExcelWhileFinished.UseVisualStyleBackColor = True
        '
        'frmDeriveDataFromExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(539, 396)
        Me.Controls.Add(Me.ChkBxOpenExcelWhileFinished)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.AddFileOrDirectoryFiles1)
        Me.Controls.Add(Me.MyDataGridView1)
        Me.Controls.Add(Me.lbSheetName)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnRemove)
        Me.Controls.Add(Me.ListBoxWorksheets)
        Me.Controls.Add(Me.BtnChoosePath)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.txtbxSavePath)
        Me.Controls.Add(Me.Label4)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmDeriveDataFromExcel"
        Me.Text = "从Excel提取数据"
        CType(Me.MyDataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BtnChoosePath As System.Windows.Forms.Button
    Friend WithEvents ListBoxWorksheets As System.Windows.Forms.ListBox
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtbxSavePath As System.Windows.Forms.TextBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents lbSheetName As System.Windows.Forms.Label
    Friend WithEvents MyDataGridView1 As AME_UserControl.myDataGridView
    Friend WithEvents AddFileOrDirectoryFiles1 As AME_UserControl.AddFileOrDirectoryFiles
    Friend WithEvents SheetName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RangeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Txtbox_DateFormat As System.Windows.Forms.TextBox
    Friend WithEvents btn_DateFormat As System.Windows.Forms.Button
    Friend WithEvents ChkboxParseDate As System.Windows.Forms.CheckBox
    Friend WithEvents ChkBxOpenExcelWhileFinished As System.Windows.Forms.CheckBox

End Class
