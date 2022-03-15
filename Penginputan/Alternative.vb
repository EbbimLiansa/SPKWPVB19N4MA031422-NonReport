Imports System.Data.OleDb
Public Class Alternative
    'Cancel Mousemove
    Private Sub Panel17_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel17.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.White
    End Sub

    Private Sub Panel21_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel21.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.White
    End Sub

    Private Sub Panel4_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel4.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.White
    End Sub

    Private Sub Panel5_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel5.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.White
    End Sub

    Private Sub Panel9_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel9.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.White
    End Sub

    'Mousemove
    Private Sub Panel7_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel7.MouseMove
        Panel12.BackColor = Color.FromArgb(48, 94, 75)
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.FromArgb(227, 230, 227)
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.White
    End Sub

    Private Sub Panel6_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel6.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.FromArgb(48, 94, 75)
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.FromArgb(227, 230, 227)
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.White
    End Sub

    Private Sub Panel8_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel8.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.FromArgb(48, 94, 75)
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.FromArgb(227, 230, 227)
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.White
    End Sub

    Private Sub Panel10_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel10.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.FromArgb(48, 94, 75)
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.FromArgb(227, 230, 227)
        Panel11.BackColor = Color.White
    End Sub

    Private Sub Panel11_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel11.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.FromArgb(48, 94, 75)

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.FromArgb(227, 230, 227)
    End Sub

    'Entry Move
    Private Sub Panel22_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel22.MouseMove
        Panel12.BackColor = Color.White
        Panel13.BackColor = Color.White
        Panel14.BackColor = Color.White
        Panel15.BackColor = Color.White
        Panel16.BackColor = Color.White

        Panel7.BackColor = Color.White
        Panel6.BackColor = Color.White
        Panel8.BackColor = Color.White
        Panel10.BackColor = Color.White
        Panel11.BackColor = Color.White

        If TxtUsername.Text = "" Then
            TxtUsername.Text = "Kode Alternative"
        End If

        If TxtPassword.Text = "" Then
            TxtPassword.Text = "Nama Alternative"
        End If
    End Sub

    Private Sub TxtUsername_MouseMove(sender As Object, e As MouseEventArgs) Handles TxtUsername.MouseMove
        If TxtUsername.Text = "Kode Alternative" Then
            TxtUsername.Text = ""
        End If
    End Sub

    Private Sub TxtPassword_MouseMove(sender As Object, e As MouseEventArgs) Handles TxtPassword.MouseMove
        If TxtPassword.Text = "Nama Alternative" Then
            TxtPassword.Text = ""
        End If
    End Sub

    'Menu
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Landing.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Landing.Show()
        Me.Close()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        WeightedProduct.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        WeightedProduct.Show()
        Me.Close()
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Laporan.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Laporan.Show()
        Me.Close()
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        About.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        About.Show()
        Me.Close()
    End Sub

    'Exit
    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        Select Case MsgBox("Apakah Anda Ingin Menutup Aplikasi Ini." & vbCrLf & "Pilih Ya Jika Anda Yakin ?", vbYesNo Or vbQuestion Or vbDefaultButton1, "Notice : Menutup Aplikasi")
            Case vbNo
                Exit Sub
            Case vbYes
                End
        End Select
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Select Case MsgBox("Apakah Anda Ingin Menutup Aplikasi Ini." & vbCrLf & "Pilih Ya Jika Anda Yakin ?", vbYesNo Or vbQuestion Or vbDefaultButton1, "Notice : Menutup Aplikasi")
            Case vbNo
                Exit Sub
            Case vbYes
                End
        End Select
    End Sub

    'Clear TXT UN PW
    Private Sub Button2_MouseMove(sender As Object, e As MouseEventArgs) Handles Button2.MouseMove
        If TxtUsername.Text = "Kode Alternative" Then
            TxtUsername.Text = ""
        End If

        If TxtPassword.Text = "Nama Alternative" Then
            TxtPassword.Text = ""
        End If
    End Sub

    'Cek Kriteria
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TxtUsername.Text <> "" And TxtPassword.Text <> "" Then
            SqlQuery = "select * from kriteria order by `Kode Kriteria` Asc"
            PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
            DataReader = PerintahDatabase.ExecuteReader
            DataGridView2.Columns.Clear()
            DataGridView2.Columns.Add("", "Kode Kriteria")
            DataGridView2.Columns.Add("", "Nama Kriteria")
            DataGridView2.Columns.Add("", "Nilai Huruf")
            DataGridView2.Columns.Add("", "Nilai Angka")
            Dim i As Integer = 0
            While DataReader.Read
                DataGridView2.Rows.Add()
                DataGridView2.Item(0, i).Value = DataReader(0)
                DataGridView2.Item(1, i).Value = DataReader(1)
                SqlQuery = "select * from data_kriteria where `Kode Alternatif`='" & TxtUsername.Text & "' and `Kode Kriteria`='" & DataReader(0) & "'"
                PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                Dim DataReader2 As OleDbDataReader = PerintahDatabase.ExecuteReader
                While DataReader2.Read
                    DataGridView2.Item(2, i).Value = DataReader2(2)
                    DataGridView2.Item(3, i).Value = DataReader2(3)
                End While
                DataReader2.Close()
                i += 1
            End While
            DataReader.Close()
            TxtUsername.Enabled = False
        Else
            MsgBox("Harap Isi Kode Dan Nama Alternatif!")
        End If
    End Sub

    'Baru
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TxtUsername.Enabled = True
        TxtUsername.Clear()
        TxtUsername.Focus()
        TxtPassword.Clear()
        DataGridView2.Columns.Clear()
    End Sub

    'Variabel 1
    Function cekDataKriteria() As Boolean
        Dim hasil As Boolean = True
        If DataGridView2.RowCount > 0 Then
            For i As Integer = 0 To DataGridView2.RowCount - 2
                If DataGridView2.Item(2, i).Value = "" Then
                    MsgBox("Nilai Huruf, Kriteria " & DataGridView2.Item(1, i).Value & " belum diisi")
                    hasil = False
                    GoTo selesai
                ElseIf Val(DataGridView2.Item(3, i).Value) < 1 Then
                    MsgBox("Nilai Angka, Kriteria " & DataGridView2.Item(1, i).Value & " harus lebih besar dari pada nol")
                    hasil = False
                    GoTo selesai
                End If
            Next
        Else
            MsgBox("Klik Cek Kriteria!")
            hasil = False
        End If
selesai:
        Return hasil
    End Function

    'Show Data
    Sub tampilData()
        DataGridView1.Columns.Clear()
        SqlQuery = "select * from Alternatif order by `Kode Alternatif`"
        PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
        DataReader = PerintahDatabase.ExecuteReader
        Dim i As Integer = 0
        While DataReader.Read
            If i = 0 Then
                DataGridView1.Columns.Add("", "Kode Alternatif")
                DataGridView1.Columns.Add("", "Nama Alternatif")
            End If
            DataGridView1.Rows.Add()
            DataGridView1.Item(0, i).Value = DataReader(0)
            DataGridView1.Item(1, i).Value = DataReader(1)
            SqlQuery = "select * from kriteria order by `Kode Kriteria`"
            PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
            Dim DataReader2 As OleDbDataReader = PerintahDatabase.ExecuteReader
            Dim j As Integer = 0
            While DataReader2.Read
                If i = 0 Then
                    DataGridView1.Columns.Add("", DataReader2(1))
                End If
                SqlQuery = "select * from data_kriteria where `Kode Alternatif`='" & DataReader(0) & "' and `Kode Kriteria`='" & DataReader2(0) & "'"
                PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                Dim DataReader3 As OleDbDataReader = PerintahDatabase.ExecuteReader
                While DataReader3.Read
                    DataGridView1.Item(j + 2, i).Value = DataReader3(2) & "(" & DataReader3(3) & ")"
                End While
                DataReader3.Close()
                j += 1
            End While
            DataReader2.Close()
            i += 1
        End While
        DataReader.Close()
    End Sub

    'Action
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TxtUsername.Text <> "" And TxtPassword.Text <> "" Then
            If cekDataKriteria() = True Then
                SqlQuery = "insert into Alternatif(`Kode Alternatif`,`Nama Alternatif`) values('" & TxtUsername.Text & "','" & TxtPassword.Text & "')"
                PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                Try
                    PerintahDatabase.ExecuteNonQuery()
                    For i As Integer = 0 To DataGridView2.RowCount - 2
                        SqlQuery = "insert into data_kriteria(`Kode Alternatif`,`Kode Kriteria`,`Nilai Huruf`,`Nilai Angka`) values ('" &
                        TxtUsername.Text & "','" & DataGridView2.Item(0, i).Value & "','" &
                        DataGridView2.Item(2, i).Value & "','" & DataGridView2.Item(3, i).Value & "')"
                        PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                        Try
                            PerintahDatabase.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message)
                            SqlQuery = "delete from Alternatif where `Kode Alternatif`='" & TxtUsername.Text & "'"
                            PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                            PerintahDatabase.ExecuteNonQuery()
                        End Try
                    Next
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Else
            MsgBox("Isilah Kode Alternatif Dan Nama Alternatif")
        End If
        tampilData()
    End Sub

    'Variabel 2
    Sub HapusDb()
        SqlQuery = "delete from Alternatif where `Kode Alternatif`='" & TxtUsername.Text & "'"
        PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
        PerintahDatabase.ExecuteNonQuery()

        SqlQuery = "delete from data_kriteria where `Kode Alternatif`='" & TxtPassword.Text & "'"
        PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
        PerintahDatabase.ExecuteNonQuery()
    End Sub

    'Action
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TxtUsername.Text <> "" And TxtPassword.Text <> "" Then
            If cekDataKriteria() = True Then
                Try
                    HapusDb()
                    Call Button3_Click(sender, e)
                    tampilData()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Else
            MsgBox("Isilah Kode Alternatif Dan Nama Alternatif")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TxtUsername.Text <> "" And TxtPassword.Text <> "" Then
            Try
                HapusDb()
                tampilData()
                Call Button1_Click(sender, e)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            MsgBox("Isilah Kode Alternatif dan Nama Alternatif")
        End If
    End Sub

    'Load
    Private Sub Showhow(sender As Object, e As EventArgs) Handles MyBase.Load
        konek()
        tampilData()
    End Sub

    'Dgv 1
    Private Sub MengisiTextBoxKetikaDataGridDiKlik(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex < DataGridView1.RowCount - 1 Then
            TxtUsername.Text = DataGridView1.Item(0, e.RowIndex).Value
            TxtPassword.Text = DataGridView1.Item(1, e.RowIndex).Value
            TxtUsername.Enabled = False
            Call Button2_Click(sender, e)
        End If
    End Sub

    'Txt UN
    Private Sub CariAlternatifBerdasarkanKodeAlternatif(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtUsername.KeyPress
        If Asc(e.KeyChar) = 13 Then
            SqlQuery = "select * from Alternatif where `Kode Alternatif`='" & TxtUsername.Text & "'"
            PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
            DataReader = PerintahDatabase.ExecuteReader
            DataReader.Read()
            If DataReader.HasRows Then
                TxtUsername.Text = DataReader(0)
                TxtPassword.Text = DataReader(1)
                TxtUsername.Enabled = False
                TxtPassword.Focus()
                Call Button2_Click(sender, e)
            Else
                TxtUsername.Enabled = False
                TxtPassword.Focus()
                Call Button2_Click(sender, e)
            End If
        End If
    End Sub
End Class