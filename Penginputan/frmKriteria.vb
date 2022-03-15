Imports System.Data.OleDb
Public Class frmKriteria
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
            TxtUsername.Text = "Kode Kriteria"
        End If

        If TxtPassword.Text = "" Then
            TxtPassword.Text = "Nama Kriteria"
        End If

        If Bobot.Text = "" Then
            Bobot.Text = "Bobot Kriteria"
        End If

        If ComboBox1.Text = "" Then
            ComboBox1.Text = "-- Pilih Atribut --"
        End If
    End Sub

    Private Sub TxtUsername_MouseMove(sender As Object, e As MouseEventArgs) Handles TxtUsername.MouseMove
        If TxtUsername.Text = "Kode Kriteria" Then
            TxtUsername.Text = ""
        End If
    End Sub

    Private Sub TxtPassword_MouseMove(sender As Object, e As MouseEventArgs) Handles TxtPassword.MouseMove
        If TxtPassword.Text = "Nama Kriteria" Then
            TxtPassword.Text = ""
        End If
    End Sub

    Private Sub Bobot_MouseMove(sender As Object, e As MouseEventArgs) Handles Bobot.MouseMove
        If Bobot.Text = "Bobot Kriteria" Then
            Bobot.Text = ""
        End If
    End Sub

    Private Sub ComboBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles ComboBox1.MouseMove
        If ComboBox1.Text = "-- Pilih Atribut --" Then
            ComboBox1.Text = ""
        End If
    End Sub

    'Variabel 1
    Sub Batal()
        TxtUsername.Enabled = True
        TxtUsername.Clear()
        TxtPassword.Clear()
        Bobot.Clear()
        ComboBox1.Text = ""
        TxtUsername.Focus()
    End Sub

    'Baru
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Batal()
    End Sub

    'Variabel 2
    Sub tampilData()
        DataAdapter = New OleDbDataAdapter("select * from kriteria order by `Kode Kriteria` Asc", Conn)
        DataSet = New DataSet
        DataAdapter.Fill(DataSet, "kriteria")
        DataGridView1.DataSource = DataSet.Tables(0)
    End Sub

    'Action
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        SqlQuery = "insert into kriteria(`Kode Kriteria`,`Nama Kriteria`,`Bobot`,`Atribut`) values('" & TxtUsername.Text & "','" & TxtPassword.Text & "','" & Bobot.Text & "','" & ComboBox1.Text & "')"
        If TxtUsername.Text <> "" And TxtPassword.Text <> "" And Val(Bobot.Text) > 0 Then
            Try
                PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                PerintahDatabase.ExecuteNonQuery()
                tampilData()
                Batal()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        SqlQuery = "update kriteria set `Nama Kriteria`='" & TxtPassword.Text & "',`Bobot`='" & Bobot.Text & "',`Atribut`='" & ComboBox1.Text & "' where `Kode Kriteria`='" & TxtUsername.Text & "'"
        If TxtUsername.Text <> "" And TxtPassword.Text <> "" And Val(Bobot.Text) > 0 Then
            Try
                PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                PerintahDatabase.ExecuteNonQuery()
                tampilData()
                Batal()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        SqlQuery = "delete from kriteria where `Kode Kriteria`='" & TxtUsername.Text & "'"
        If TxtUsername.Text <> "" And TxtPassword.Text <> "" And Val(Bobot.Text) > 0 Then
            Try
                PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                PerintahDatabase.ExecuteNonQuery()
                tampilData()
                Batal()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    'Dgv 1
    Private Sub TampilKeTextBoxKetikaKlikDataGrid(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TxtUsername.Enabled = False
        TxtUsername.Text = DataGridView1.Item(0, e.RowIndex).Value
        TxtPassword.Text = DataGridView1.Item(1, e.RowIndex).Value
        Bobot.Text = DataGridView1.Item(2, e.RowIndex).Value
        ComboBox1.Text = DataGridView1.Item(3, e.RowIndex).Value
    End Sub

    'Load
    Private Sub Showhow(sender As Object, e As EventArgs) Handles MyBase.Load
        konek()
        tampilData()
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
End Class