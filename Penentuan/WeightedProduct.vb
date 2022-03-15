Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Public Class WeightedProduct
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
    End Sub

    'Variabel 1
    Dim Perhitungan As New wp

    Private Function sumBobot(ByVal nilai As Bobot())
        Dim hasil As Double = 0
        For Each item As Bobot In nilai
            hasil += item.bobot
        Next
        Return hasil
    End Function

    Sub TampilkanPerbaikanBobot()
        DataGridView2.Columns.Clear()
        DataGridView2.Columns.Add("bobot", "Bobot")
        DataGridView2.Columns.Add("perbaikan", "Perbaikan Bobot")
        For i As Integer = 0 To UBound(Perhitungan.DataBobot)
            DataGridView2.Rows.Add()
            DataGridView2.Item("bobot", i).Value = Perhitungan.DataBobot(i).kode
            DataGridView2.Item("perbaikan", i).Value = Math.Round(Perhitungan.DataBobot(i).bobot / sumBobot(Perhitungan.DataBobot), 2)
        Next
    End Sub

    Sub TampilkanAlternatif()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("no", "No")
        DataGridView1.Columns.Add("kode", "Kode")
        DataGridView1.Columns.Add("alternatif", "Nama Alternatif")
        For j As Integer = 0 To UBound(Perhitungan.DataBobot)
            DataGridView1.Columns.Add(Perhitungan.DataBobot(j).kode, Perhitungan.DataBobot(j).nama & "(" & Perhitungan.DataBobot(j).bobot & ")")
        Next
        DataGridView1.Columns.Add("hasils", "Hasil S")
        DataGridView1.Columns.Add("hasilv", "Hasil V")
        DataGridView1.Columns.Add("hasil", "Hasil")

        Dim i As Integer = 0
        For Each item As Alternatif In Perhitungan.DataAlternatif
            DataGridView1.Rows.Add()
            DataGridView1.Item("no", i).Value = i + 1
            DataGridView1.Item("kode", i).Value = item.kode
            SqlQuery = "select * from Alternatif where `Kode Alternatif`='" & item.kode & "'"
            PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
            Try
                DataReader = PerintahDatabase.ExecuteReader
                While DataReader.Read
                    DataGridView1.Item("alternatif", i).Value = DataReader(1)
                End While
                DataReader.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            DataGridView1.Item("hasils", i).Value = item.hasilS
            DataGridView1.Item("hasilv", i).Value = item.hasilV
            DataGridView1.Item("hasil", i).Value = item.hasil * 100
            For j As Integer = 0 To UBound(Perhitungan.DataAlternatif(i).kriteria)
                With Perhitungan.DataAlternatif(i)
                    DataGridView1.Item(.kriteria(j).kode, i).Value = .kriteria(j).nilai
                End With
            Next

            i += 1
        Next
    End Sub

    'Action
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Mengisi Bobot
        Dim dataBobotSementara As New List(Of Bobot)
        SqlQuery = "select * from kriteria"
        PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
        DataReader = PerintahDatabase.ExecuteReader
        While DataReader.Read
            Dim bobotSementara As New Bobot
            bobotSementara.kode = DataReader(0)
            bobotSementara.nama = DataReader(1)
            bobotSementara.bobot = DataReader(2)
            If DataReader(3).ToString.ToLower = "biaya" Then
                bobotSementara.atribut = False
            Else
                bobotSementara.atribut = True
            End If

            dataBobotSementara.Add(bobotSementara)
        End While
        DataReader.Close()
        Perhitungan.JumlahKriteria = dataBobotSementara.Count
        Perhitungan.DataBobot = dataBobotSementara.ToArray

        'mengisi data alternatif
        Dim dataAlternatifSementara As New List(Of Alternatif)
        SqlQuery = "select * from Alternatif"
        PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
        DataReader = PerintahDatabase.ExecuteReader
        While DataReader.Read
            Dim AlternatifSementara As New Alternatif
            AlternatifSementara.kode = DataReader(0)
            SqlQuery = "select * from data_kriteria where `Kode Alternatif`='" & DataReader(0) & "'"
            PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
            Dim DataReader2 As OleDbDataReader = PerintahDatabase.ExecuteReader
            Dim dataKriteriaSementara As New List(Of Kriteria)
            While DataReader2.Read
                Dim satuanKriteria As New Kriteria
                satuanKriteria.kode = DataReader2(1)
                satuanKriteria.nilai = DataReader2(3)
                dataKriteriaSementara.Add(satuanKriteria)
            End While
            DataReader2.Close()
            AlternatifSementara.kriteria = dataKriteriaSementara.ToArray

            dataAlternatifSementara.Add(AlternatifSementara)
        End While
        DataReader.Close()
        Perhitungan.JumlahAlternatif = dataAlternatifSementara.Count
        Perhitungan.DataAlternatif = dataAlternatifSementara.ToArray
        Perhitungan.CariWP()

        TampilkanAlternatif()
        TampilkanPerbaikanBobot()
    End Sub

    'Load
    Private Sub Showhow(sender As Object, e As EventArgs) Handles MyBase.Load
        konek()
    End Sub

    'Action
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.RowCount > 1 Then
            Try
                SqlQuery = "delete from hasil"
                PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                PerintahDatabase.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            For i As Integer = 0 To DataGridView1.RowCount - 2
                SqlQuery = "insert into hasil(`Kode Alternatif`,`Hasil`) values('" & DataGridView1.Item("kode", i).Value & "','" & DataGridView1.Item("hasil", i).Value & "')"
                PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                Try
                    PerintahDatabase.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Next
            MsgBox("Data Tersimpan")
        End If
    End Sub

    'Variabel 2
    Sub TampilkanHasil()
        SqlQuery = "select * from hasil order by `Hasil` Desc"
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("", "Kode Alternatif")
        DataGridView1.Columns.Add("", "Nama Alternatif")
        DataGridView1.Columns.Add("", "Hasil")
        PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
        Try
            DataReader = PerintahDatabase.ExecuteReader
            Dim i As Integer = 0
            While DataReader.Read
                DataGridView1.Rows.Add()
                DataGridView1.Item(0, i).Value = DataReader(0)
                SqlQuery = "select * from Alternatif where `Kode Alternatif`='" & DataReader(0) & "'"
                PerintahDatabase = New OleDbCommand(SqlQuery, Conn)
                Try
                    Dim DataReader2 As OleDbDataReader = PerintahDatabase.ExecuteReader
                    While DataReader2.Read
                        DataGridView1.Item(1, i).Value = DataReader2(1)
                    End While
                    DataReader2.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                DataGridView1.Item(2, i).Value = DataReader(1)
                i += 1
            End While
            DataReader.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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
        MsgBox("Anda Sudah Berada Di Halaman Weighted Product", , "Notice : Posisi Menu")
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        MsgBox("Anda Sudah Berada Di Halaman Weighted Product", , "Notice : Posisi Menu")
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