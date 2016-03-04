Imports System.Data.SqlClient
Imports System.Data.DataTable
Public Class frmKhachHang
    Dim database As New DataTable ' Tạo đối tượng database để lưu trữ dữ liệu từ Database Online
    'Tạo chuỗi kết nối để kết nối tới Database Online
    Dim chuoiconnect As String = "workstation id=taitntps02714.mssql.somee.com;packet size=4096;user id=taitntps02714_SQLLogin_1;pwd=bul9ipjdpk;data source=taitntps02714.mssql.somee.com;persist security info=False;initial catalog=taitntps02714"
    Dim connect As SqlConnection = New SqlConnection(chuoiconnect)

    'Sự kiện Load form KhachHang'
    Private Sub frmKhachHang_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect) ' Tạo đối tượng kết nối tới DB  Online
        ' Câu truy vấn để get dữ liệu
        Dim Query2 As SqlDataAdapter = New SqlDataAdapter("select * from KhachHang", connect)
        'Kết nối mở ra
        connect.Open()
        'Đổ dữ liệu vào đối tượng database
        Query2.Fill(database)
        'Hiển thị dữ liệu ra Datagridview
        dgvKH.DataSource = database.DefaultView
    End Sub

    'Sự kiện Click ô dữ liệu trong DataGridView'
    Private Sub dgvKH_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvKH.CellContentClick
        Dim index As Integer = dgvKH.CurrentCell.RowIndex
        txtMaKH.Text = dgvKH.Item(0, index).Value
        txtTenKH.Text = dgvKH.Item(1, index).Value
        txtSDT.Text = dgvKH.Item(2, index).Value
        txtDiaChi.Text = dgvKH.Item(3, index).Value
    End Sub

    'Sự kiện Click Thêm khách hàng'
    Private Sub btnThemKH_Click(sender As Object, e As EventArgs) Handles btnThemKH.Click
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        'Tạo query câu truy vấn Insert into
        Dim TaoKH As String = "insert into KhachHang values (@MaKhachHang,@TenKhachHang,@SoDT,@DiaChi)"
        'Tạo đối tượng để thực thi câu truy vấn với DB ONline
        Dim AddKH As New SqlCommand(TaoKH, connect)
        'Kết nối mở ra
        connect.Open()

        Try
            'Truyền giá trị trong các ô textbox cho các biến tương ứng
            AddKH.Parameters.AddWithValue("@MaKhachHang", txtMaKH.Text)
            AddKH.Parameters.AddWithValue("@TenKhachHang", txtTenKH.Text)
            AddKH.Parameters.AddWithValue("@SoDT", txtSDT.Text)
            AddKH.Parameters.AddWithValue("@DiaChi", txtDiaChi.Text)
            'Exucute là ghi dữ liệu vào Database
            MessageBox.Show("Thêm thành công khách hàng")
            AddKH.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Không thành công!")
        End Try
        txtTenKH.Clear()
        txtMaKH.Clear()
        txtSDT.Clear()
        txtDiaChi.Clear()
        txtMaKH.Focus()

        'Cập nhật dữ liệu'
        Loaddata()
    End Sub

    'Định nghĩa hàm Loaddata để cập nhật lại dữ liệu'
    Private Sub Loaddata()
        database.Clear()
        dgvKH.DataSource = database
        dgvKH.DataSource = Nothing
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        Dim Query2 As SqlDataAdapter = New SqlDataAdapter("select * from KhachHang", connect)

        connect.Open()
        Query2.Fill(database)
        dgvKH.DataSource = database.DefaultView
    End Sub

    'Sự kiện Click Sửa khách hàng'
    Private Sub btnSuaKH_Click(sender As Object, e As EventArgs) Handles btnSuaKH.Click
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        connect.Open()
        Dim SuaKH As String = "Update KhachHang set TenKhachHang= @TenKhachHang, SoDT= @SoDT, DiaChi= @DiaChi where MaKhachHang= @MaKhachHang"
        Dim EditKH As New SqlCommand(SuaKH, connect)

        Try
            EditKH.Parameters.AddWithValue("@MaKhachHang", txtMaKH.Text)
            EditKH.Parameters.AddWithValue("@TenKhachHang", txtTenKH.Text)
            EditKH.Parameters.AddWithValue("@SoDT", txtSDT.Text)
            EditKH.Parameters.AddWithValue("@DiaChi", txtDiaChi.Text)
            EditKH.ExecuteNonQuery()
            connect.Close()
            MessageBox.Show("Sửa Thành Công")
        Catch ex As Exception
            MessageBox.Show("Sửa Không Thành Công")
        End Try
        'Cập nhật dữ liệu'
        Loaddata()
    End Sub

    'Sự kiện Click Xóa khách hàng'
    Private Sub btnXoaKH_Click(sender As Object, e As EventArgs) Handles btnXoaKH.Click
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        connect.Open()
        Dim XoaKH As String = "Delete from KhachHang where MaKhachHang=@MaKhachHang"
        Dim DelKH As New SqlCommand(XoaKH, connect)
        Try
            DelKH.Parameters.AddWithValue("@MaKhachHang", txtMaKH.Text)
            DelKH.ExecuteNonQuery()
            MessageBox.Show("Xóa Thành Công")
        Catch ex As Exception
            MessageBox.Show("Xóa Không Thành Công")
        End Try
        txtTenKH.Clear()
        txtMaKH.Clear()
        txtSDT.Clear()
        txtDiaChi.Clear()
        txtMaKH.Focus()
        Loaddata()
    End Sub

    'Sự kiện Click Tìm khách hàng'
    Private Sub btnTimKH_Click(sender As Object, e As EventArgs) Handles btnTimKH.Click
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        'kết nối mở'
        connect.Open()
        Dim database As New DataTable
        database.Clear()
        dgvKH.DataSource = database
        dgvKH.DataSource = Nothing
        Dim connectnone As SqlConnection = New SqlConnection(chuoiconnect)
        Dim Query As SqlDataAdapter = New SqlDataAdapter("select * from KhachHang where MaKH like '" & txtMaKH.Text & "'", connectnone)
        connectnone.Open()
        Query.Fill(database)
        dgvKH.DataSource = database.DefaultView
    End Sub
End Class