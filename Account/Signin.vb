Imports System.Data.OleDb
Public Class Signin
    'Cancel Move
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

    'Move
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
            TxtUsername.Text = "Full Name"
        End If

        If TxtPassword.Text = "" Then
            TxtPassword.Text = "Password"
        End If

        If TxtPassword.Text = "Password" Then
            TxtPassword.UseSystemPasswordChar = False
            Button4.Focus()
        End If
    End Sub

    Private Sub TxtUsername_MouseMove(sender As Object, e As MouseEventArgs) Handles TxtUsername.MouseMove
        If TxtUsername.Text = "Full Name" Then
            TxtUsername.Text = ""
        End If
    End Sub

    Private Sub TxtPassword_MouseMove(sender As Object, e As MouseEventArgs) Handles TxtPassword.MouseMove
        If TxtPassword.Text = "Password" Then
            TxtPassword.Text = ""
        End If
    End Sub

    'Entry Press
    Private Sub TxtUsername_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtUsername.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TxtPassword.Focus()
        End If
    End Sub

    Private Sub TxtPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtPassword.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Button4.PerformClick()
        End If
    End Sub

    'CKH
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If TxtPassword.UseSystemPasswordChar = True Then
            TxtPassword.UseSystemPasswordChar = False
            TxtPassword.PasswordChar = ""
            CheckBox1.Text = "Sembunyikan Password"
        Else
            TxtPassword.UseSystemPasswordChar = True
            CheckBox1.Text = "Tampilkan Password"
        End If
    End Sub

    Private Sub TxtPassword_Click(sender As Object, e As EventArgs) Handles TxtPassword.Click
        TxtPassword.UseSystemPasswordChar = True
    End Sub

    'Dashboard
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dashboard.Show()
        Dashboard.Panel7.Visible = False
        Dashboard.Panel6.Visible = False
        Dashboard.Panel8.Visible = False
        Dashboard.Panel9.Visible = False
        Dashboard.Label6.Text = "Sudah Memiliki Sebuah Akun ?"
        Dashboard.Button2.Text = "Sign In"
        Me.Close()
    End Sub

    'Varibel
    Sub Batal()
        TxtUsername.Text = ""
        TxtPassword.Text = ""
    End Sub

    'Login
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TxtUsername.Text = "" Then
            MsgBox("Anda Lupa Memasukkan Username, Silahkan Memasukkan Username Terlebih Dahulu!", , "Notice : Masukkan Inputan Username")
        ElseIf TxtPassword.Text = "" Then
            MsgBox("Anda Lupa Memasukkan Password, Silahkan Memasukkan Password Terlebih Dahulu!", , "Notice : Masukkan Inputan Password")
        Else
            Call konek()
            Cmdlogin = New OleDbCommand("Select * From login where Username='" & TxtUsername.Text & "' and Password='" & TxtPassword.Text & "'", Conn)
            'Hapus Drlogin Jika MySQL
            Drlogin = Cmdlogin.ExecuteReader
            Drlogin.Read()
            If Drlogin.HasRows Then
                Me.Hide()
                Dashboard.Panel7.Visible = True
                Dashboard.Panel9.Visible = True
                Dashboard.Panel8.Visible = True
                Dashboard.Panel6.Visible = True
                Dashboard.Show()
                Me.Close()
                Dashboard.Label6.Text = "Apakah Kamu Ingin Keluar Dari Akun ?"
                Dashboard.Button2.Text = "Log Out"
            Else
                MsgBox("Maaf Anda Tidak Bisa Melakukan Login, Periksa Kembali Username Dan Password Anda!", , "Notice : Username Dan Password Salah")
                Batal()
                TxtUsername.Focus()
                Return
            End If
        End If
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

    'Menu
    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        About.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        About.Show()
        Me.Close()
    End Sub

    'Information
    Private Sub Ldashboard_Click(sender As Object, e As EventArgs) Handles Ldashboard.Click
        MsgBox("Anda Belum Masuk Ke Akun, Silahkan Masuk Ke Akun Anda Terlebih Dahulu Dan Pastikan Anda Seorang Admin!", , "Notice : Silahkan Daftar Akun Anda Melalui Admin.")
    End Sub
End Class