Imports System.Data.SqlClient
'Sự kiện Load form Login'
Public Class frmLogin
    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtTenDangNhap.Clear()
        txtMatKhau.Clear()
    End Sub

    'Sự kiện Click Login'
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Dim chuoiketnoi As String = "workstation id=taitntps02714.mssql.somee.com;packet size=4096;user id=taitntps02714_SQLLogin_1;pwd=bul9ipjdpk;data source=taitntps02714.mssql.somee.com;persist security info=False;initial catalog=taitntps02714"
        Dim ketnoi As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim sqlAdapter As New SqlDataAdapter("Select * from NhanVien where MaNhanVien='" & txtTenDangNhap.Text & "' And password='" & txtMatKhau.Text & "' ", ketnoi)
        Dim tb As New DataTable

        Try
            ketnoi.Open()
            sqlAdapter.Fill(tb)
            If tb.Rows.Count > 0 Then
                MessageBox.Show("Đăng nhập thành công")
                frmMain.Show()
            Else
                MessageBox.Show("Sai tên đăng nhập hoặc mật khẩu")
                txtTenDangNhap.Clear()
                txtMatKhau.Clear()
                txtTenDangNhap.Focus()
            End If
        Catch ex As Exception
        End Try
    End Sub

    'Sự kiện Click Exit'
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
