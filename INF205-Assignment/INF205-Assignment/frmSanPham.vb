Imports System.Data.SqlClient
Imports System.Data.DataTable
Public Class frmSanPham
    Dim database As New DataTable ' Tạo đối tượng database để lưu trữ dữ liệu từ Database Online'
    'Tạo chuỗi kết nối để kết nối tới Database Online'
    Dim chuoiconnect As String = "workstation id=taitntps02714.mssql.somee.com;packet size=4096;user id=taitntps02714_SQLLogin_1;pwd=bul9ipjdpk;data source=taitntps02714.mssql.somee.com;persist security info=False;initial catalog=taitntps02714"
    Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
    'Sự kiện Load form SanPham'
    Private Sub frmSanPham_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect) ' Tạo đối tượng kết nối tới DB  Online'
        ' Câu truy vấn để get dữ liệu'
        Dim Query1 As SqlDataAdapter = New SqlDataAdapter("select * from SanPham", connect)
        'Kết nối mở ra
        connect.Open()
        'Đổ dữ liệu vào đối tượng database'
        Query1.Fill(database)
        'Hiển thị dữ liệu ra Datagridview'
        dgvSP.DataSource = database.DefaultView
    End Sub
    'Sự kiện Click ô dữ liệu trong Datagridview'
    Private Sub dgvSP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSP.CellContentClick
        Dim index As Integer = dgvSP.CurrentCell.RowIndex
        txtMaSP.Text = dgvSP.Item(0, index).Value
        txtTenSP.Text = dgvSP.Item(1, index).Value
        txtDongia.Text = dgvSP.Item(2, index).Value
        txtSoluong.Text = dgvSP.Item(3, index).Value
        txtChitiet.Text = dgvSP.Item(4, index).Value
        txtLoaisp.Text = dgvSP.Item(5, index).Value
    End Sub
    'Sự kiện Click thêm sản phẩm'
    Private Sub btnThem_Click(sender As Object, e As EventArgs) Handles btnThemSP.Click
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        'Tạo query câu truy vấn Insert into'
        Dim TaoSP As String = "insert into SanPham values (@MaSP,@TenSP,@DonGia,@SoLuong,@ChiTietSP,@LoaiSanPham)"
        'Tạo đối tượng để thực thi câu truy vấn với DB online'
        Dim AddSP As New SqlCommand(TaoSP, connect)
        'Kết nối mở ra
        connect.Open()

        Try
            'Truyền giá trị trong các ô textbox cho các biến tương ứng'
            AddSP.Parameters.AddWithValue("@MaSP", txtMaSP.Text)
            AddSP.Parameters.AddWithValue("@TenSP", txtTenSP.Text)
            AddSP.Parameters.AddWithValue("@DonGia", txtDongia.Text)
            AddSP.Parameters.AddWithValue("@SoLuong", txtSoluong.Text)
            AddSP.Parameters.AddWithValue("@ChiTietSP", txtChitiet.Text)
            AddSP.Parameters.AddWithValue("@LoaiSanPham", txtLoaisp.Text)
            MessageBox.Show("Thêm thành công sản phẩm")
            'Exucute là ghi dữ liệu vào Database'
            AddSP.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Không thành công!")
        End Try
        txtMaSP.Clear()
        txtTenSP.Clear()
        txtDongia.Clear()
        txtSoluong.Clear()
        txtChitiet.Clear()
        txtLoaisp.Clear()
        txtMaSP.Focus()

        'Cập nhật dữ liệu'
        Loaddata()
    End Sub
    'Định nghĩa hàm Loaddata để cập nhật lại dữ liệu'
    Private Sub Loaddata()
        database.Clear()
        dgvSP.DataSource = database
        dgvSP.DataSource = Nothing
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        Dim Query1 As SqlDataAdapter = New SqlDataAdapter("select * from SanPham", connect)

        connect.Open()
        Query1.Fill(database)
        dgvSP.DataSource = database.DefaultView
    End Sub
    'Sự kiện Click Sửa sản phẩm'
    Private Sub btnSua_Click(sender As Object, e As EventArgs) Handles btnSuaSP.Click
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        'kết nối mở'
        connect.Open()
        'Tạo query câu truy vấn Update'
        Dim SuaSP As String = "Update SanPham set TenSP= @TenSP, DonGia= @DonGia, SoLuong= @SoLuong, ChiTietSP=@ChiTietSP, LoaiSanPham_MaLoai=@LoaiSanPham_MaLoai where MaSP= @MaSP"
        'Tạo đối tượng để thực thi câu truy vấn với DB online'
        Dim EditSP As New SqlCommand(SuaSP, connect)

        Try
            'Truyền giá trị trong các ô textbox cho các biến tương ứng'
            EditSP.Parameters.AddWithValue("@MaSP", txtMaSP.Text)
            EditSP.Parameters.AddWithValue("@TenSP", txtTenSP.Text)
            EditSP.Parameters.AddWithValue("@DonGia", txtDongia.Text)
            EditSP.Parameters.AddWithValue("@SoLuong", txtSoluong.Text)
            EditSP.Parameters.AddWithValue("@ChiTietSP", txtChitiet.Text)
            EditSP.Parameters.AddWithValue("@LoaiSanPham_MaLoai", txtLoaisp.Text)
            'Exucute là ghi dữ liệu vào Database'
            EditSP.ExecuteNonQuery()
            connect.Close()
            MessageBox.Show("Sửa Thành Công")
        Catch ex As Exception
            MessageBox.Show("Sửa Không Thành Công")
        End Try
        txtMaSP.Clear()
        txtTenSP.Clear()
        txtDongia.Clear()
        txtSoluong.Clear()
        txtChitiet.Clear()
        txtLoaisp.Clear()
        txtMaSP.Focus()
        'Cập nhật dữ liệu'
        Loaddata()
    End Sub
    'Sự kiện Click Xóa sản phẩm'
    Private Sub btnXoa_Click(sender As Object, e As EventArgs) Handles btnXoaSP.Click
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        'kết nối mở'
        connect.Open()
        'Tạo query câu lệnh truy vấn Delete'
        Dim XoaSP As String = "Delete from SanPham where MaSP=@MaSP"
        'Tạo đối tượng để thực thi câu truy vấn với DB online'
        Dim DelSP As New SqlCommand(XoaSP, connect)
        Try
            'Truyền giá trị trong các ô textbox cho các biến tương ứng'
            DelSP.Parameters.AddWithValue("@MaSP", txtMaSP.Text)
            'Exucute là ghi dữ liệu vào Database'
            DelSP.ExecuteNonQuery()
            MessageBox.Show("Xóa Thành Công")
        Catch ex As Exception
            MessageBox.Show("Xóa Không Thành Công")
        End Try
        txtMaSP.Clear()
        txtTenSP.Clear()
        txtDongia.Clear()
        txtSoluong.Clear()
        txtChitiet.Clear()
        txtLoaisp.Clear()
        txtMaSP.Focus()
        'Cập nhật dữ liệu'
        Loaddata()
    End Sub
    'Sự kiện Click Tìm sản phẩm'
    Private Sub btnTimSP_Click(sender As Object, e As EventArgs) Handles btnTimSP.Click
        Dim connect As SqlConnection = New SqlConnection(chuoiconnect)
        'kết nối mở'
        connect.Open()
        Dim database As New DataTable
        database.Clear()
        dgvSP.DataSource = database
        dgvSP.DataSource = Nothing
        Dim connectnone As SqlConnection = New SqlConnection(chuoiconnect)
        Dim Query As SqlDataAdapter = New SqlDataAdapter("select * from SanPham where MaSP like '" & txtMaSP.Text & "'", connectnone)
        connectnone.Open()
        Query.Fill(database)
        dgvSP.DataSource = database.DefaultView     
    End Sub
End Class