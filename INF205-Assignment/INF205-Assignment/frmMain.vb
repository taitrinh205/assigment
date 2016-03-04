Public Class frmMain
    'Sự kiện form Main load'
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmLogin.ShowDialog()
    End Sub
    'Sự kiện Click mở form Sản Phẩm từ Menu'
    Private Sub SảnPhẩmToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SảnPhẩmToolStripMenuItem.Click
        frmSanPham.ShowDialog()
    End Sub
    'Sự kiện Click mở form Khách Hàng từ Menu'
    Private Sub KháchHàngToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KháchHàngToolStripMenuItem.Click
        frmKhachHang.ShowDialog()
    End Sub
    
    'Sự kiện Click Đăng xuất từ Menu'
    Private Sub ThoátToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThoátToolStripMenuItem.Click
        Me.Hide()
        frmLogin.ShowDialog()
    End Sub
    'Sự kiện Thoát ứng dụng từ Menu'
    Private Sub ThoátToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles Thoát2ToolStripMenuItem1.Click
        Application.Exit()
    End Sub
End Class