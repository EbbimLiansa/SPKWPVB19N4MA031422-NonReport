Imports System.Data.OleDb
Public Class Laporan
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

    'Menu
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Landing.Show()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Landing.Show()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        WeightedProduct.Show()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        WeightedProduct.Show()
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        MsgBox("Anda Sudah Berada Di Halaman Laporan", , "Notice : Posisi Menu")
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        MsgBox("Anda Sudah Berada Di Halaman Laporan", , "Notice : Posisi Menu")
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        About.Show()
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        About.Show()
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

    'Load
    Private Sub Showhow(sender As Object, e As EventArgs) Handles MyBase.Load
        konek()
    End Sub

    'Show Laporan
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class