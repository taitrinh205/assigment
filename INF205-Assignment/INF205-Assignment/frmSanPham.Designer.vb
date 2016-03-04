<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSanPham
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
        Me.dgvSP = New System.Windows.Forms.DataGridView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnTimSP = New System.Windows.Forms.Button()
        Me.btnXoaSP = New System.Windows.Forms.Button()
        Me.btnSuaSP = New System.Windows.Forms.Button()
        Me.btnThemSP = New System.Windows.Forms.Button()
        Me.txtChitiet = New System.Windows.Forms.TextBox()
        Me.lblMaLoai = New System.Windows.Forms.Label()
        Me.txtSoluong = New System.Windows.Forms.TextBox()
        Me.lblChiTiet = New System.Windows.Forms.Label()
        Me.txtDongia = New System.Windows.Forms.TextBox()
        Me.lblHangSx = New System.Windows.Forms.Label()
        Me.txtTenSP = New System.Windows.Forms.TextBox()
        Me.lblTenSP = New System.Windows.Forms.Label()
        Me.txtMaSP = New System.Windows.Forms.TextBox()
        Me.lblMaSP = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtLoaisp = New System.Windows.Forms.TextBox()
        CType(Me.dgvSP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvSP
        '
        Me.dgvSP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvSP.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dgvSP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSP.Location = New System.Drawing.Point(14, 261)
        Me.dgvSP.Name = "dgvSP"
        Me.dgvSP.Size = New System.Drawing.Size(658, 161)
        Me.dgvSP.TabIndex = 10
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.btnTimSP)
        Me.GroupBox1.Controls.Add(Me.btnXoaSP)
        Me.GroupBox1.Controls.Add(Me.btnSuaSP)
        Me.GroupBox1.Controls.Add(Me.btnThemSP)
        Me.GroupBox1.Controls.Add(Me.txtLoaisp)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtChitiet)
        Me.GroupBox1.Controls.Add(Me.lblMaLoai)
        Me.GroupBox1.Controls.Add(Me.txtSoluong)
        Me.GroupBox1.Controls.Add(Me.lblChiTiet)
        Me.GroupBox1.Controls.Add(Me.txtDongia)
        Me.GroupBox1.Controls.Add(Me.lblHangSx)
        Me.GroupBox1.Controls.Add(Me.txtTenSP)
        Me.GroupBox1.Controls.Add(Me.lblTenSP)
        Me.GroupBox1.Controls.Add(Me.txtMaSP)
        Me.GroupBox1.Controls.Add(Me.lblMaSP)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Aqua
        Me.GroupBox1.Location = New System.Drawing.Point(224, 15)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(448, 228)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Nhập Thông Tin Sản Phẩm"
        '
        'btnTimSP
        '
        Me.btnTimSP.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.btnTimSP.ForeColor = System.Drawing.Color.SteelBlue
        Me.btnTimSP.Image = Global.INF205_Assignment.My.Resources.Resources.Search
        Me.btnTimSP.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnTimSP.Location = New System.Drawing.Point(356, 121)
        Me.btnTimSP.Name = "btnTimSP"
        Me.btnTimSP.Size = New System.Drawing.Size(86, 66)
        Me.btnTimSP.TabIndex = 9
        Me.btnTimSP.Text = "Tìm"
        Me.btnTimSP.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnTimSP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnTimSP.UseVisualStyleBackColor = True
        '
        'btnXoaSP
        '
        Me.btnXoaSP.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.btnXoaSP.ForeColor = System.Drawing.Color.DarkRed
        Me.btnXoaSP.Image = Global.INF205_Assignment.My.Resources.Resources.Delete
        Me.btnXoaSP.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnXoaSP.Location = New System.Drawing.Point(262, 121)
        Me.btnXoaSP.Name = "btnXoaSP"
        Me.btnXoaSP.Size = New System.Drawing.Size(88, 66)
        Me.btnXoaSP.TabIndex = 8
        Me.btnXoaSP.Text = "Xóa"
        Me.btnXoaSP.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnXoaSP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnXoaSP.UseVisualStyleBackColor = True
        '
        'btnSuaSP
        '
        Me.btnSuaSP.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.btnSuaSP.ForeColor = System.Drawing.Color.DarkOrange
        Me.btnSuaSP.Image = Global.INF205_Assignment.My.Resources.Resources.Edit
        Me.btnSuaSP.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnSuaSP.Location = New System.Drawing.Point(356, 32)
        Me.btnSuaSP.Name = "btnSuaSP"
        Me.btnSuaSP.Size = New System.Drawing.Size(88, 70)
        Me.btnSuaSP.TabIndex = 7
        Me.btnSuaSP.Text = "Sửa"
        Me.btnSuaSP.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnSuaSP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnSuaSP.UseVisualStyleBackColor = True
        '
        'btnThemSP
        '
        Me.btnThemSP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnThemSP.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.btnThemSP.ForeColor = System.Drawing.Color.ForestGreen
        Me.btnThemSP.Image = Global.INF205_Assignment.My.Resources.Resources.ADD1
        Me.btnThemSP.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnThemSP.Location = New System.Drawing.Point(262, 32)
        Me.btnThemSP.Name = "btnThemSP"
        Me.btnThemSP.Size = New System.Drawing.Size(88, 70)
        Me.btnThemSP.TabIndex = 6
        Me.btnThemSP.Text = "Thêm"
        Me.btnThemSP.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnThemSP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnThemSP.UseVisualStyleBackColor = True
        '
        'txtChitiet
        '
        Me.txtChitiet.Location = New System.Drawing.Point(101, 164)
        Me.txtChitiet.Name = "txtChitiet"
        Me.txtChitiet.Size = New System.Drawing.Size(155, 23)
        Me.txtChitiet.TabIndex = 5
        '
        'lblMaLoai
        '
        Me.lblMaLoai.Location = New System.Drawing.Point(12, 167)
        Me.lblMaLoai.Name = "lblMaLoai"
        Me.lblMaLoai.Size = New System.Drawing.Size(55, 22)
        Me.lblMaLoai.TabIndex = 0
        Me.lblMaLoai.Text = "Chi Tiết :"
        '
        'txtSoluong
        '
        Me.txtSoluong.Location = New System.Drawing.Point(101, 131)
        Me.txtSoluong.Name = "txtSoluong"
        Me.txtSoluong.Size = New System.Drawing.Size(155, 23)
        Me.txtSoluong.TabIndex = 4
        '
        'lblChiTiet
        '
        Me.lblChiTiet.Location = New System.Drawing.Point(12, 140)
        Me.lblChiTiet.Name = "lblChiTiet"
        Me.lblChiTiet.Size = New System.Drawing.Size(83, 22)
        Me.lblChiTiet.TabIndex = 0
        Me.lblChiTiet.Text = "Số Lượng :"
        '
        'txtDongia
        '
        Me.txtDongia.Location = New System.Drawing.Point(101, 98)
        Me.txtDongia.Name = "txtDongia"
        Me.txtDongia.Size = New System.Drawing.Size(155, 23)
        Me.txtDongia.TabIndex = 3
        '
        'lblHangSx
        '
        Me.lblHangSx.Location = New System.Drawing.Point(9, 105)
        Me.lblHangSx.Name = "lblHangSx"
        Me.lblHangSx.Size = New System.Drawing.Size(86, 22)
        Me.lblHangSx.TabIndex = 0
        Me.lblHangSx.Text = "Đơn Giá :"
        '
        'txtTenSP
        '
        Me.txtTenSP.Location = New System.Drawing.Point(101, 65)
        Me.txtTenSP.Name = "txtTenSP"
        Me.txtTenSP.Size = New System.Drawing.Size(155, 23)
        Me.txtTenSP.TabIndex = 2
        '
        'lblTenSP
        '
        Me.lblTenSP.Location = New System.Drawing.Point(9, 70)
        Me.lblTenSP.Name = "lblTenSP"
        Me.lblTenSP.Size = New System.Drawing.Size(86, 22)
        Me.lblTenSP.TabIndex = 0
        Me.lblTenSP.Text = "Tên sản phẩm:"
        '
        'txtMaSP
        '
        Me.txtMaSP.Location = New System.Drawing.Point(101, 32)
        Me.txtMaSP.Name = "txtMaSP"
        Me.txtMaSP.Size = New System.Drawing.Size(155, 23)
        Me.txtMaSP.TabIndex = 1
        '
        'lblMaSP
        '
        Me.lblMaSP.Location = New System.Drawing.Point(11, 35)
        Me.lblMaSP.Name = "lblMaSP"
        Me.lblMaSP.Size = New System.Drawing.Size(86, 22)
        Me.lblMaSP.TabIndex = 0
        Me.lblMaSP.Text = "Mã sản phẩm:"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.Image = Global.INF205_Assignment.My.Resources.Resources.product
        Me.PictureBox1.Location = New System.Drawing.Point(14, 13)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(204, 230)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 196)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 25)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Loại SP"
        '
        'txtLoaisp
        '
        Me.txtLoaisp.Location = New System.Drawing.Point(101, 193)
        Me.txtLoaisp.Name = "txtLoaisp"
        Me.txtLoaisp.Size = New System.Drawing.Size(155, 23)
        Me.txtLoaisp.TabIndex = 5
        '
        'frmSanPham
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.INF205_Assignment.My.Resources.Resources.wallpaper_approved3
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.ClientSize = New System.Drawing.Size(686, 434)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.dgvSP)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.Name = "frmSanPham"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quản Lý Sản Phẩm"
        CType(Me.dgvSP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvSP As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSoluong As System.Windows.Forms.TextBox
    Friend WithEvents lblChiTiet As System.Windows.Forms.Label
    Friend WithEvents txtDongia As System.Windows.Forms.TextBox
    Friend WithEvents lblHangSx As System.Windows.Forms.Label
    Friend WithEvents txtTenSP As System.Windows.Forms.TextBox
    Friend WithEvents lblTenSP As System.Windows.Forms.Label
    Friend WithEvents txtMaSP As System.Windows.Forms.TextBox
    Friend WithEvents lblMaSP As System.Windows.Forms.Label
    Friend WithEvents btnTimSP As System.Windows.Forms.Button
    Friend WithEvents btnXoaSP As System.Windows.Forms.Button
    Friend WithEvents btnSuaSP As System.Windows.Forms.Button
    Friend WithEvents btnThemSP As System.Windows.Forms.Button
    Friend WithEvents txtChitiet As System.Windows.Forms.TextBox
    Friend WithEvents lblMaLoai As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents txtLoaisp As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
