'Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar
Imports System.Data.Odbc
'Imports System.Data.Common
'Imports System.Drawing
Imports Spire.Doc
'Imports Spire.Doc.Document
Imports Spire.Doc.Fields
Imports Spire.Doc.Documents
'Imports System.Drawing.Printing
Imports Spire.Pdf
'Imports MySqlConnector

Public Class Form1

    Dim conn As OdbcConnection
    Dim cmd As OdbcCommand
    Dim ds As DataSet
    Dim da As OdbcDataAdapter
    Dim rd As OdbcDataReader
    Dim MyDB As String

    Sub Connect()
        Try
            MyDB = "Driver={MySQL ODBC 8.0 Unicode Driver};Database=db_pasien;Server=localhost;uid=root"
            conn = New OdbcConnection(MyDB)
            If conn.State = ConnectionState.Closed Then conn.Open()
        Catch ex As Exception
            Dim Confirm As DialogResult = MsgBox("Tidak dapat terkoneksi dengan database. Cek xampp apakah database sudah dinyalakan", MessageBoxButtons.OK)
            For i As Integer = 0 To 1 Step 0
                If Confirm = DialogResult.OK Then
                    Application.Exit()
                End If
            Next
        End Try
    End Sub

    Sub Reset_textbox()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        RichTextBox1.Text = ""
    End Sub

    Sub Enter_press_on_TextBox()
        Call Connect()
        Dim Query_String As String = "Select * From tbl_pasien Where "
        Dim Add_String As String = ""
        'Dim Other_Params As Boolean = False
        Dim First_Params_Found As Boolean = False

        If TextBox1.Text <> "" Then
            First_Params_Found = True
            Dim words = TextBox1.Text.Split(" ")
            For i As Integer = 0 To words.Count - 1
                If i = 0 Then
                    Add_String = "Nama Like '%" + words(i) + "%' "
                Else
                    Add_String = Add_String + "And Nama Like '%" + words(i) + "%' "
                End If
            Next
            'Add_String = "Nama = '" & TextBox1.Text & "' "
            Query_String = Query_String + Add_String
            'If TextBox2.Text <> "" Then
            '    Query_String = Query_String + "And "
            'End If
        End If

        If TextBox2.Text <> "" Then
            'TextBox4.Text = CStr(First_Params_Found)
            Dim Umur As Integer
            If Int32.TryParse(TextBox2.Text, Umur) = False Then
                MsgBox("Field Umur harus diisi dengan angka")
            Else
                Select Case First_Params_Found
                    Case Is = True
                        Query_String = Query_String + "And "
                        Exit Select
                    Case Is = False
                        First_Params_Found = True
                End Select
                'If First_Params_Found = False Then
                '    First_Params_Found = True
                'Else
                '    Query_String = Query_String + "And "
                '    'If Other_Params = False Then
                '    '    Other_Params = True
                '    '    Query_String = Query_String + "And "
                '    'End If
                'End If
                'First_Params_Found = False
                Add_String = "Umur = '" & TextBox2.Text & "' "
                Query_String = Query_String + Add_String
                'If TextBox3.Text <> "" Then
                '    Query_String = Query_String + "And "
                'End If
            End If
        End If

        If TextBox3.Text <> "" Then
            Select Case First_Params_Found
                Case Is = True
                    Query_String = Query_String + "And "
                    Exit Select
                Case Is = False
                    First_Params_Found = True
            End Select
            'If First_Params_Found = False Then
            '    First_Params_Found = True
            'Else
            '    Query_String = Query_String + "And "
            '    'If Other_Params = False Then
            '    '    Other_Params = True
            '    '    Query_String = Query_String + "And "
            '    'End If
            'End If
            'First_Params = False
            'Add_String = "Alamat = '" & TextBox3.Text & "' "
            Dim words = TextBox3.Text.Split(" ")
            For i As Integer = 0 To words.Count - 1
                If i = 0 Then
                    Add_String = "Alamat Like '%" + words(i) + "%' "
                Else
                    Add_String = Add_String + "And Alamat Like '%" + words(i) + "%' "
                End If
            Next
            Query_String = Query_String + Add_String
            'If TextBox4.Text <> "" Then
            '    Query_String = Query_String + "And "
            'End If
        End If

        If TextBox4.Text <> "" Then
            Select Case First_Params_Found
                Case Is = True
                    Query_String = Query_String + "And "
                    Exit Select
                Case Is = False
                    First_Params_Found = True
            End Select
            'Add_String = "`No. Hp` = '" & TextBox4.Text & "' "
            Add_String = "`No. Hp` Like '%" + TextBox4.Text + "%' "
            Query_String = Query_String + Add_String
            If TextBox8.Text <> "" And TextBox9.Text <> "" And TextBox10.Text <> "" Then
                Query_String = Query_String
            ElseIf TextBox8.Text = "" And TextBox9.Text = "" And TextBox10.Text = "" Then
                Query_String = Query_String
            Else
                MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
                Exit Sub
            End If
        End If

        Dim tanggal_input As Boolean = False

        If TextBox8.Text <> "" And TextBox9.Text <> "" And TextBox10.Text <> "" Then
            Dim Temp As Integer

            If TextBox8.Text(0) = "0" And TextBox8.Text.Count = 2 Then
                TextBox8.Text = TextBox8.Text(1)
            End If

            If TextBox9.Text(0) = "0" And TextBox9.Text.Count = 2 Then
                TextBox9.Text = TextBox9.Text(1)
            End If

            If Integer.TryParse(TextBox8.Text, Temp) = False Or Integer.TryParse(TextBox9.Text, Temp) = False Or Integer.TryParse(TextBox10.Text, Temp) = False Then
                MsgBox("Input tanggal tidak valid")
                Exit Sub
            End If

            If TextBox8.Text.Trim().Length() < 1 Or TextBox8.Text.Trim().Length() > 2 Or TextBox9.Text.Trim().Length() < 1 Or TextBox9.Text.Trim().Length() > 2 Or TextBox10.Text.Trim().Length() <> 4 Then
                MsgBox("Input tanggal tidak valid")
                Exit Sub
            End If

            Dim Hari_Pada_Bulan_Awal As Integer = Hari_Pada_Bulan(TextBox9.Text, Integer.Parse(TextBox10.Text))
            If Integer.Parse(TextBox8.Text) > Hari_Pada_Bulan_Awal Or Integer.Parse(TextBox8.Text) <= 0 Then
                If Hari_Pada_Bulan_Awal = -1 Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If
                MsgBox("Tanggal Tidak Valid!!!")
                Exit Sub
            End If

            Select Case First_Params_Found
                Case Is = True
                    Query_String = Query_String + "And "
                    Exit Select
                Case Is = False
                    First_Params_Found = True
            End Select
            Add_String = "(`Tanggal Input` Between '" & TextBox10.Text & "-" & TextBox9.Text & "-" & TextBox8.Text & " 00:00:00 ' And '" & TextBox10.Text & "-" & TextBox9.Text & "-" & TextBox8.Text & " 23:59:59') "
            Query_String = Query_String + Add_String
            If TextBox13.Text <> "" And TextBox12.Text <> "" And TextBox11.Text <> "" Then
                Query_String = Query_String
                tanggal_input = True
            ElseIf TextBox13.Text = "" And TextBox12.Text = "" And TextBox11.Text = "" Then
                Query_String = Query_String
            Else
                MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
                Exit Sub
            End If
        ElseIf TextBox8.Text = "" And TextBox9.Text = "" And TextBox10.Text = "" Then
            Query_String = Query_String
        Else
            MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
            Exit Sub
        End If

        If TextBox13.Text <> "" And TextBox12.Text <> "" And TextBox11.Text <> "" Then
            Dim Temp As Integer

            If TextBox13.Text(0) = "0" And TextBox13.Text.Count = 2 Then
                TextBox13.Text = TextBox13.Text(1)
            End If

            If TextBox12.Text(0) = "0" And TextBox12.Text.Count = 2 Then
                TextBox12.Text = TextBox12.Text(1)
            End If

            If Integer.TryParse(TextBox13.Text, Temp) = False Or Integer.TryParse(TextBox12.Text, Temp) = False Or Integer.TryParse(TextBox11.Text, Temp) = False Then
                MsgBox("Input tanggal tidak valid")
                Exit Sub
            End If
            If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                MsgBox("Input tanggal tidak valid")
                Exit Sub
            End If
            If tanggal_input = True Then
                If Integer.Parse(TextBox10.Text) > Integer.Parse(TextBox11.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                ElseIf Integer.Parse(TextBox9.Text) > Integer.Parse(TextBox12.Text) And Integer.Parse(TextBox10.Text) = Integer.Parse(TextBox11.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Dim Hari_Pada_Bulan_Awal As Integer = Hari_Pada_Bulan(TextBox9.Text, Integer.Parse(TextBox10.Text))
                Dim Hari_Pada_Bulan_Akhir As Integer = Hari_Pada_Bulan(TextBox12.Text, Integer.Parse(TextBox11.Text))
                If (Integer.Parse(TextBox8.Text) > Hari_Pada_Bulan_Awal Or Integer.Parse(TextBox8.Text) <= 0) Or (Integer.Parse(TextBox13.Text) > Hari_Pada_Bulan_Akhir Or Integer.Parse(TextBox13.Text) <= 0) Then
                    If Hari_Pada_Bulan_Awal = -1 Or Hari_Pada_Bulan_Akhir = -1 Then
                        MsgBox("Tanggal Tidak Valid!!!")
                        Exit Sub
                    End If
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Dim Bulan_Awal As String = Nama_Bulan(TextBox9.Text)
                Dim Bulan_Akhir As String = Nama_Bulan(TextBox12.Text)
                If Bulan_Awal = "Nomor bulan tidak valid!!!" Or Bulan_Akhir = "Nomor bulan tidak valid!!!" Then
                    MsgBox("Nomor bulan tidak valid!!!")
                    Exit Sub
                End If

                Dim jumlah_hari As Integer = Hitung_Hari(Integer.Parse(TextBox8.Text), Integer.Parse(TextBox9.Text), Integer.Parse(TextBox10.Text), Integer.Parse(TextBox13.Text), Integer.Parse(TextBox12.Text), Integer.Parse(TextBox11.Text), False)
                If jumlah_hari <= 0 Then
                    MsgBox("Tanggal tidak valid")
                    Exit Sub
                End If
            Else
                Dim Hari_Pada_Bulan_Akhir As Integer = Hari_Pada_Bulan(TextBox12.Text, Integer.Parse(TextBox11.Text))
                If Integer.Parse(TextBox13.Text) > Hari_Pada_Bulan_Akhir Or Integer.Parse(TextBox13.Text) <= 0 Then
                    If Hari_Pada_Bulan_Akhir = -1 Then
                        MsgBox("Tanggal Tidak Valid!!!")
                        Exit Sub
                    End If
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If
            End If
            Select Case First_Params_Found
                Case Is = True
                    Query_String = Query_String + "And "
                    Exit Select
                Case Is = False
                    First_Params_Found = True
            End Select
            Add_String = "(`Update Terakhir` Between '" & TextBox11.Text & "-" & TextBox12.Text & "-" & TextBox13.Text & " 00:00:00 ' And '" & TextBox11.Text & "-" & TextBox12.Text & "-" & TextBox13.Text & " 23:59:59') "
            Query_String = Query_String + Add_String
            'If RichTextBox1.Text <> "" Then
            '    Query_String = Query_String + "And "
            'End If
        ElseIf TextBox13.Text = "" And TextBox12.Text = "" And TextBox11.Text = "" Then
            Query_String = Query_String
        Else
            MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
            Exit Sub
        End If

        If RichTextBox1.Text <> "" Then
            Select Case First_Params_Found
                Case Is = True
                    Query_String = Query_String + "And "
                    Exit Select
                Case Is = False
                    First_Params_Found = True
            End Select
            'Add_String = "Tindakan = '" & RichTextBox1.Text & "' "
            Dim words = RichTextBox1.Text.Split(" ")
            For i As Integer = 0 To words.Count - 1
                If i = 0 Then
                    Add_String = "Tindakan Like '%" + words(i) + "%' "
                ElseIf i = words.Count - 1 Then
                    words(i).Replace(ControlChars.Lf, "")
                    Add_String = Add_String + "And Tindakan Like '%" + words(i) + "%' "
                Else
                    Add_String = Add_String + "And Tindakan Like '%" + words(i) + "%' "
                End If
            Next
            Query_String = Query_String + Add_String
        End If

        'TextBox6.Text = Query_String
        'TextBox4.Text = CStr(Other_Params)
        'Label13.Text = Query_String

        da = New OdbcDataAdapter(Query_String, conn)
        ds = New DataSet
        da.Fill(ds, "tbl_pasien")
        DataGridView1.DataSource = ds.Tables("tbl_pasien")
        conn.Close()

        'Label13.Text = words(0)
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        DataGridView1.ClearSelection()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox6.Enabled = False
        TextBox5.Enabled = False
        TextBox7.Enabled = False
        TextBox5.Visible = True
        TextBox7.Visible = True
        RichTextBox1.Enabled = True
        TextBox8.Visible = False
        TextBox9.Visible = False
        TextBox10.Visible = False
        TextBox11.Visible = False
        TextBox12.Visible = False
        TextBox13.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        DataGridView2.Visible = False

        TextBox2.Font = New Font(TextBox2.Font, FontStyle.Regular)

        Label1.Visible = True
        Label2.Visible = True

        Button1.Visible = True
        Button1.Text = "Perbarui"

        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True

        Label6.Text = "Id"
        Label1.Text = "Nama Pasien"
        Label2.Text = "Umur"
        Label3.Text = "Alamat"
        Label4.Text = "No. HP"
        Label5.Text = "Tanggal Input"
        Label7.Text = "Update Terakhir"

        TextBox14.Visible = False
        TextBox15.Visible = False
        TextBox16.Visible = False
        TextBox17.Visible = False
        TextBox18.Visible = False
        TextBox19.Visible = False

        TextBox6.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = True
        RichTextBox1.Visible = True

        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False

        CheckBox1.Visible = False
        CheckBox2.Visible = False
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox6.Enabled = False
        TextBox5.Enabled = False
        TextBox7.Enabled = False
        TextBox5.Visible = True
        TextBox7.Visible = True
        RichTextBox1.Enabled = True
        TextBox8.Visible = False
        TextBox9.Visible = False
        TextBox10.Visible = False
        TextBox11.Visible = False
        TextBox12.Visible = False
        TextBox13.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        DataGridView2.Visible = False

        TextBox1.Visible = True
        TextBox2.Visible = True

        Call Reset_textbox()

        TextBox2.Font = New Font(TextBox2.Font, FontStyle.Regular)

        Label1.Visible = True
        Label2.Visible = True

        Button1.Visible = True
        Button1.Text = "Masukkan ke Database"

        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True

        Label6.Text = "Id"
        Label1.Text = "Nama Pasien"
        Label2.Text = "Umur"
        Label3.Text = "Alamat"
        Label4.Text = "No. HP"
        Label5.Text = "Tanggal Input"
        Label7.Text = "Update Terakhir"

        TextBox14.Visible = False
        TextBox15.Visible = False
        TextBox16.Visible = False
        TextBox17.Visible = False
        TextBox18.Visible = False
        TextBox19.Visible = False

        TextBox6.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = True
        RichTextBox1.Visible = True

        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False

        CheckBox1.Visible = False
        CheckBox2.Visible = False
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        DataGridView1.ClearSelection()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox6.Enabled = True
        TextBox5.Enabled = False
        TextBox7.Enabled = False
        TextBox5.Visible = False
        TextBox7.Visible = False
        TextBox8.Visible = True
        TextBox9.Visible = True
        TextBox10.Visible = True
        TextBox11.Visible = True
        TextBox12.Visible = True
        TextBox13.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = True
        Label12.Visible = True
        RichTextBox1.Enabled = True
        'Button1.Visible = False
        Button1.Text = "Cari Pasien"
        DataGridView2.Visible = False

        TextBox8.Enabled = True
        TextBox9.Enabled = True
        TextBox10.Enabled = True
        TextBox11.Enabled = True
        TextBox12.Enabled = True
        TextBox13.Enabled = True

        TextBox1.Visible = True
        TextBox2.Visible = True

        TextBox2.Font = New Font(TextBox2.Font, FontStyle.Regular)

        Label1.Visible = True
        Label2.Visible = True

        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True

        Label6.Text = "Id"
        Label1.Text = "Nama Pasien"
        Label2.Text = "Umur"
        Label3.Text = "Alamat"
        Label4.Text = "No. HP"
        Label5.Text = "Tanggal Input"
        Label7.Text = "Update Terakhir"

        TextBox14.Visible = False
        TextBox15.Visible = False
        TextBox16.Visible = False
        TextBox17.Visible = False
        TextBox18.Visible = False
        TextBox19.Visible = False

        TextBox6.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = True
        RichTextBox1.Visible = True

        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False

        CheckBox1.Visible = False
        CheckBox2.Visible = False
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Connect()
        da = New OdbcDataAdapter("Select * From tbl_pasien", conn)
        ds = New DataSet
        da.Fill(ds, "tbl_pasien")
        DataGridView1.DataSource = ds.Tables("tbl_pasien")
        conn.Close()

        DataGridView3.Visible = False

        Dim InstalledPrinters As String

        For Each InstalledPrinters In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            DataGridView2.Rows.Add(InstalledPrinters)
        Next InstalledPrinters
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Umur As Integer
        If RadioButton1.Checked = True Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Semua field harus diisi!!!")
            ElseIf Int32.TryParse(TextBox2.Text, Umur) = False Then
                MsgBox("Field Umur harus diisi dengan angka")
            Else
                Dim message As String
                If RichTextBox1.Text = "" Then
                    message = "Apakah anda yakin ingin menambah data pasien ini? (Karena isi tindakan kosong, tindakan ini tidak akan masuk ke tabel Riwayat Tindakan)"
                Else
                    message = "Apakah anda yakin ingin menambah data pasien ini?"
                End If
                Dim Confirm As DialogResult = MessageBox.Show(message, "Konfirmasi Penambahan Data", MessageBoxButtons.YesNo)
                If Confirm = DialogResult.Yes Then
                    Call Connect()
                    Dim InputData As String = "Insert into tbl_pasien (Nama, Umur, Alamat, `No. Hp`,`Tanggal Input`, `Update Terakhir`, Tindakan) Values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "', '" & TextBox4.Text & "', (Select NOW()), (Select NOW()),  '" & RichTextBox1.Text & "')"
                    cmd = New OdbcCommand(InputData, conn)
                    cmd.ExecuteNonQuery()

                    If RichTextBox1.Text <> "" Then
                        InputData = "Select Id From tbl_pasien Where Id = (Select MAX(Id) From tbl_pasien)"
                        cmd = New OdbcCommand(InputData, conn)
                        Dim dr As Integer = cmd.ExecuteScalar()

                        Dim id_tindakan As String
                        id_tindakan = "T1#" + dr.ToString
                        'Label13.Text = id_tindakan

                        InputData = "Select Tindakan From tbl_pasien Where Id = (Select MAX(Id) From tbl_pasien)"
                        cmd = New OdbcCommand(InputData, conn)
                        Dim dr_tindakan As String = cmd.ExecuteScalar()

                        InputData = "Insert into tbl_tindakan Values ('" & id_tindakan & "', (Select Id From tbl_pasien Where Id = '" & dr & "'), (Select Nama From tbl_pasien Where Id = '" & dr & "'), (Select NOW()), '" & dr_tindakan & "')"
                        cmd = New OdbcCommand(InputData, conn)
                        cmd.ExecuteNonQuery()
                    End If

                    MsgBox("Input Data Berhasil")
                    conn.Close()
                    Call Reset_textbox()
                    Call Connect()
                    da = New OdbcDataAdapter("Select * From tbl_pasien", conn)
                    ds = New DataSet
                    da.Fill(ds, "tbl_pasien")
                    DataGridView1.DataSource = ds.Tables("tbl_pasien")
                    conn.Close()
                End If
            End If
        ElseIf RadioButton3.Checked = True Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Semua field harus diisi!!!")
            ElseIf TextBox6.Text = "" Then
                MsgBox("Pilih terlebih dahulu data pasien yang akan diperbarui")
            ElseIf Int32.TryParse(TextBox2.Text, Umur) = False Then
                MsgBox("Field Umur harus diisi dengan angka")
            Else
                Dim message As String
                Dim Riwayat_sebelumnya As String = DataGridView1.SelectedRows(0).Cells(7).Value
                If RichTextBox1.Text = "" Then
                    message = "Apakah anda yakin ingin memperbarui data pasien ini? (Karena isi tindakan kosong, tindakan ini tidak akan masuk ke tabel Riwayat Tindakan. Tetapi kolom tindakan pada tabel Pasien akan kosong)"
                Else
                    message = "Apakah anda yakin ingin memperbarui data pasien ini?"
                End If
                Dim Confirm As DialogResult = MessageBox.Show(message, "Konfirmasi Perbaruan Data", MessageBoxButtons.YesNo)
                If Confirm = DialogResult.Yes Then
                    Call Connect()
                    Dim EditData As String = "Update tbl_pasien set Nama = '" & TextBox1.Text & "', Umur = '" & TextBox2.Text & "', Alamat = '" & TextBox3.Text & "', `No. Hp` = '" & TextBox4.Text & "', `Update Terakhir` = (Select NOW()), Tindakan = '" & RichTextBox1.Text & "' Where Id = '" & TextBox6.Text & "'"
                    cmd = New OdbcCommand(EditData, conn)
                    cmd.ExecuteNonQuery()

                    If RichTextBox1.Text <> "" And Riwayat_sebelumnya <> RichTextBox1.Text Then
                        EditData = "Select * From tbl_tindakan Where Id_pasien = '" & TextBox6.Text & "'"
                        cmd = New OdbcCommand(EditData, conn)
                        Dim dr As OdbcDataReader = cmd.ExecuteReader()
                        Dim id_tindakan As String
                        Dim count As Integer = 0

                        If dr.HasRows Then
                            While dr.Read()
                                count += 1
                            End While
                            count += 1
                            id_tindakan = "T" + count.ToString + "#" + TextBox6.Text
                        Else
                            id_tindakan = "T1#" + TextBox6.Text
                        End If

                        EditData = "Insert into tbl_tindakan Values ('" & id_tindakan & "', (Select Id From tbl_pasien Where Id = '" & TextBox6.Text & "'), (Select Nama From tbl_pasien Where Id = '" & TextBox6.Text & "'), (Select NOW()), '" & RichTextBox1.Text & "')"
                        cmd = New OdbcCommand(EditData, conn)
                        cmd.ExecuteNonQuery()
                    End If

                    MsgBox("Update Data Berhasil")
                    conn.Close()
                    Call Reset_textbox()
                    Call Connect()
                    da = New OdbcDataAdapter("Select * From tbl_pasien", conn)
                    ds = New DataSet
                    da.Fill(ds, "tbl_pasien")
                    DataGridView1.DataSource = ds.Tables("tbl_pasien")
                    conn.Close()
                End If
            End If
        ElseIf RadioButton4.Checked = True Then
            If TextBox2.Text <> "" Then
                If Int32.TryParse(TextBox2.Text, Umur) = False Then
                    MsgBox("Field Umur harus diisi dengan angka")
                    Exit Sub
                End If
            End If
            Call Enter_press_on_TextBox()
        ElseIf RadioButton2.Checked = True Then
            Dim Temp As Integer
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox14.Text = "" Or TextBox15.Text = "" Or TextBox16.Text = "" Or TextBox17.Text = "" Or TextBox18.Text = "" Or TextBox19.Text = "" Then
                MsgBox("Semua Field harus diisi!!!")
                Exit Sub
            End If

            If Int32.TryParse(TextBox2.Text, Umur) = False Then
                MsgBox("Field Umur harus diisi dengan angka")
                Exit Sub
            ElseIf Int32.TryParse(TextBox14.Text, Temp) = False Or Int32.TryParse(TextBox15.Text, Temp) = False Or Int32.TryParse(TextBox16.Text, Temp) = False Or Int32.TryParse(TextBox17.Text, Temp) = False Or Int32.TryParse(TextBox18.Text, Temp) = False Or Int32.TryParse(TextBox19.Text, Temp) = False Then
                MsgBox("Tanggal Tidak Valid!!!")
                Exit Sub
            ElseIf TextBox14.Text.Trim().Length <> 4 Or TextBox17.Text.Trim().Length <> 4 Or TextBox15.Text.Trim().Length < 1 Or TextBox15.Text.Trim().Length > 2 Or TextBox18.Text.Trim().Length < 1 Or TextBox18.Text.Trim().Length > 2 Or TextBox16.Text.Trim().Length < 1 Or TextBox16.Text.Trim().Length > 2 Or TextBox19.Text.Trim().Length < 1 Or TextBox19.Text.Trim().Length > 2 Then
                MsgBox("Tanggal Tidak Valid!!!")
                Exit Sub
            End If

            If TextBox19.Text(0) = "0" And TextBox19.Text.Count = 2 Then
                TextBox19.Text = TextBox19.Text(1)
            End If

            If TextBox16.Text(0) = "0" And TextBox16.Text.Count = 2 Then
                TextBox16.Text = TextBox16.Text(1)
            End If

            If TextBox18.Text(0) = "0" And TextBox18.Text.Count = 2 Then
                TextBox18.Text = TextBox18.Text(1)
            End If

            If TextBox15.Text(0) = "0" And TextBox15.Text.Count = 2 Then
                TextBox15.Text = TextBox15.Text(1)
            End If

            If Integer.Parse(TextBox17.Text) > Integer.Parse(TextBox14.Text) Then
                MsgBox("Tanggal Tidak Valid!!!")
                Exit Sub
            ElseIf Integer.Parse(TextBox18.Text) > Integer.Parse(TextBox15.Text) And Integer.Parse(TextBox17.Text) = Integer.Parse(TextBox14.Text) Then
                MsgBox("Tanggal Tidak Valid!!!")
                Exit Sub
            End If

            Dim Hari_Pada_Bulan_Awal As Integer = Hari_Pada_Bulan(TextBox18.Text, Integer.Parse(TextBox17.Text))
            Dim Hari_Pada_Bulan_Akhir As Integer = Hari_Pada_Bulan(TextBox15.Text, Integer.Parse(TextBox14.Text))
            If (Integer.Parse(TextBox19.Text) > Hari_Pada_Bulan_Awal Or Integer.Parse(TextBox19.Text) <= 0) Or (Integer.Parse(TextBox16.Text) > Hari_Pada_Bulan_Akhir Or Integer.Parse(TextBox16.Text) <= 0) Then
                If Hari_Pada_Bulan_Awal = -1 Or Hari_Pada_Bulan_Akhir = -1 Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If
                MsgBox("Tanggal Tidak Valid!!!")
                Exit Sub
            End If

            Dim Bulan_Awal As String = Nama_Bulan(TextBox18.Text)
            Dim Bulan_Akhir As String = Nama_Bulan(TextBox15.Text)
            If Bulan_Awal = "Nomor bulan tidak valid!!!" Or Bulan_Akhir = "Nomor bulan tidak valid!!!" Then
                MsgBox("Nomor bulan tidak valid!!!")
                Exit Sub
            End If

            Dim jumlah_hari As Integer = Hitung_Hari(Integer.Parse(TextBox19.Text), Integer.Parse(TextBox18.Text), Integer.Parse(TextBox17.Text), Integer.Parse(TextBox16.Text), Integer.Parse(TextBox15.Text), Integer.Parse(TextBox14.Text), True)
            If jumlah_hari <= 0 Or jumlah_hari > 30 Then
                MsgBox("Jumlah hari terlalu banyak atau tanggal tidak valid")
                Exit Sub
            End If

            Dim Surat_Sakit As New Document()
            Try
                Surat_Sakit.LoadFromFile("Template_Surat_Sakit.docx")
            Catch ex As Exception
                MsgBox("Template Surat Istirahat Sakit Tidak Ditemukan!!!")
                Exit Sub
            End Try

            Dim section As Section = Surat_Sakit.Sections(0)
            Dim para12 As Paragraph = section.Paragraphs(11)
            Dim para13 As Paragraph = section.Paragraphs(12)
            Dim para15 As Paragraph = section.Paragraphs(14)
            Dim para16 As Paragraph = section.Paragraphs(15)
            Dim para17 As Paragraph = section.Paragraphs(16)
            Dim para32 As Paragraph = section.Paragraphs(31)

            Dim tgl_hariini = Today.Day.ToString
            Dim bulan_ini = Today.Month.ToString
            Dim tahun_ini = Today.Year.ToString

            bulan_ini = Nama_Bulan(bulan_ini)

            Dim tr1 As TextRange = para12.AppendText(TextBox1.Text)
            Dim tr2 As TextRange = para13.AppendText(TextBox2.Text)
            Dim tr3 As TextRange = para15.AppendText("Selama " + CStr(jumlah_hari) + " hari")
            Dim tr4 As TextRange = para16.AppendText("Mulai Tanggal " + TextBox19.Text + " " + Bulan_Awal + " " + TextBox17.Text)
            Dim tr5 As TextRange = para17.AppendText("Sampai Tanggal " + TextBox16.Text + " " + Bulan_Akhir + " " + TextBox14.Text)
            Dim tr6 As TextRange = para32.AppendText(tgl_hariini + " " + bulan_ini + " " + tahun_ini)

            tr1.CharacterFormat.FontName = "Times New Roman"
            tr2.CharacterFormat.FontName = "Times New Roman"
            tr3.CharacterFormat.FontName = "Times New Roman"
            tr4.CharacterFormat.FontName = "Times New Roman"
            tr5.CharacterFormat.FontName = "Times New Roman"
            tr6.CharacterFormat.FontName = "Times New Roman"

            tr1.CharacterFormat.FontSize = 12
            tr2.CharacterFormat.FontSize = 12
            tr3.CharacterFormat.FontSize = 12
            tr4.CharacterFormat.FontSize = 12
            tr5.CharacterFormat.FontSize = 12
            tr6.CharacterFormat.FontSize = 12

            tr1.CharacterFormat.TextColor = Color.Black
            tr2.CharacterFormat.TextColor = Color.Black
            tr3.CharacterFormat.TextColor = Color.Black
            tr4.CharacterFormat.TextColor = Color.Black
            tr5.CharacterFormat.TextColor = Color.Black
            tr6.CharacterFormat.TextColor = Color.Black

            Dim file_name As String = "Surat Sakit\Surat Sakit " + TextBox1.Text + " " + DateTime.Now.ToString("dd MMMM yyyy") + ".pdf"

            Surat_Sakit.SaveToFile(file_name, Spire.Doc.FileFormat.PDF)
            Dim Confirm As DialogResult = MessageBox.Show("Dokumen tersimpan." & vbCrLf & "Apakah anda ingin mencetak dokumen ini dengan printer " + DataGridView2.SelectedRows(0).Cells(0).Value + "?", "Konfirmasi Pencetakan Dokumen", MessageBoxButtons.YesNo)

            If Confirm = DialogResult.Yes Then
                Dim SuratSakit_print As PdfDocument = New PdfDocument()
                SuratSakit_print.LoadFromFile(file_name)
                SuratSakit_print.PrintSettings.PrinterName = DataGridView2.SelectedRows(0).Cells(0).Value
                Try
                    SuratSakit_print.Print()
                Catch ex As Exception
                    MsgBox("Terjadi kesalahan pada pencetakan dokumen")
                End Try
            End If
        ElseIf RadioButton5.Checked = True Then
            'Dim Temp As Integer
            Dim Temp_uang As Int64
            If TextBox6.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Semua field harus diisi!!!")
                Exit Sub
            ElseIf TextBox6.Text.Length > 47 Then
                MsgBox("Field Id terlalu panjang (Maksimal 47 karakter)")
                Exit Sub
            ElseIf TextBox3.Text.Length > 100 Then
                MsgBox("Field Id terlalu panjang (Maksimal 100 karakter)")
                Exit Sub
            ElseIf TextBox4.Text.Length > 55 Then
                MsgBox("Nama ttd terlalu panjang (Maksimal 55 karakter)")
                Exit Sub
            ElseIf TextBox2.Text.Length > 15 Then
                MsgBox("Tidak dapat mencetak nominal uang tersebut (Maksimal Rp 999,999,999,999,999)")
                Exit Sub
            ElseIf TextBox2.Text(0) = "0" Then
                MsgBox("Jumlah nominal tidak valid")
                Exit Sub
            ElseIf Int64.TryParse(TextBox2.Text, Temp_uang) = False Then
                MsgBox("Field Banyaknya Uang harus diisi dengan angka")
                Exit Sub
            End If

            Dim Kwitansi As New Document()
            Try
                Kwitansi.LoadFromFile("Template_Kwitansi.docx")
            Catch ex As Exception
                MsgBox("Template Kwitansi Tidak Ditemukan!!!")
                Exit Sub
            End Try

            Dim len_char_uang As Integer = TextBox2.Text.Length

            Dim string_uang As String = "Rupiah"
            Dim count As Integer = 0

            Dim digit_ratus As New List(Of String)()
            Dim digit_ribu As New List(Of String)()
            Dim digit_juta As New List(Of String)()
            Dim digit_miliar As New List(Of String)()
            Dim digit_triliun As New List(Of String)()

            For i As Integer = len_char_uang - 1 To 0 Step -1
                If count < 3 Then
                    digit_ratus.Add(TextBox2.Text(i))
                ElseIf count < 6 Then
                    digit_ribu.Add(TextBox2.Text(i))
                ElseIf count < 9 Then
                    digit_juta.Add(TextBox2.Text(i))
                ElseIf count < 12 Then
                    digit_miliar.Add(TextBox2.Text(i))
                ElseIf count < 15 Then
                    digit_triliun.Add(TextBox2.Text(i))
                End If
                count += 1
            Next

            If digit_ratus.Count <= 3 Then
                For a As Integer = 0 To digit_ratus.Count - 1
                    If digit_ratus(a) <> "0" Then
                        If a = 0 Then
                            If digit_ratus.Count >= 2 Then
                                If digit_ratus(1) <> "1" Then
                                    string_uang = angka_ke_kata(digit_ratus(a)) + " " + string_uang
                                End If
                            Else
                                string_uang = angka_ke_kata(digit_ratus(a)) + " " + string_uang
                            End If
                        ElseIf a = 1 Then
                            If digit_ratus(a) = "1" Then
                                If digit_ratus(0) = "1" Then
                                    string_uang = "Sebelas " + string_uang
                                ElseIf digit_ratus(0) = "0" Then
                                    string_uang = "Sepuluh " + string_uang
                                Else
                                    string_uang = angka_ke_kata(digit_ratus(0)) + " Belas " + string_uang
                                End If
                            Else
                                string_uang = angka_ke_kata(digit_ratus(a)) + " Puluh " + string_uang
                            End If
                        ElseIf a = 2 Then
                            If digit_ratus(a) = "1" Then
                                string_uang = "Seratus " + string_uang
                                Exit For
                            Else
                                string_uang = angka_ke_kata(digit_ratus(a)) + " Ratus " + string_uang
                                Exit For
                            End If
                        End If
                        Continue For
                    End If
                Next
            End If

            If digit_ribu.Count <= 3 And digit_ribu.Count > 0 Then
                Dim ribu As Boolean = False
                For a As Integer = 0 To digit_ribu.Count - 1
                    If digit_ribu(a) <> "0" Then
                        If a = 0 Then
                            If digit_ribu.Count = 1 Then
                                If digit_ribu(a) = 1 Then
                                    string_uang = "Seribu " + string_uang
                                Else
                                    string_uang = angka_ke_kata(digit_ribu(a)) + " Ribu " + string_uang
                                End If
                                Exit For
                            ElseIf digit_ribu(1) <> "1" Then
                                string_uang = angka_ke_kata(digit_ribu(a)) + " Ribu " + string_uang
                                ribu = True
                            End If
                        ElseIf a = 1 Then
                            If digit_ribu(a) = "1" Then
                                If digit_ribu(0) = "0" Then
                                    string_uang = "Sepuluh" + " Ribu " + string_uang
                                    ribu = True
                                ElseIf digit_ribu(0) = "1" Then
                                    string_uang = "Sebelas" + " Ribu " + string_uang
                                    ribu = True
                                ElseIf digit_ribu(0) <> "0" And digit_ribu(0) <> "1" Then
                                    string_uang = angka_ke_kata(digit_ribu(0)) + " Belas " + "Ribu " + string_uang
                                    ribu = True
                                End If
                            Else
                                If ribu = False Then
                                    string_uang = angka_ke_kata(digit_ribu(a)) + " Puluh " + "Ribu " + string_uang
                                    ribu = True
                                Else
                                    string_uang = angka_ke_kata(digit_ribu(a)) + " Puluh " + string_uang
                                End If
                            End If
                        ElseIf a = 2 Then
                            If digit_ribu(a) = "1" Then
                                If ribu = False Then
                                    string_uang = "Seratus " + "Ribu " + string_uang
                                    ribu = True
                                    Exit For
                                Else
                                    string_uang = "Seratus " + string_uang
                                    Exit For
                                End If
                            Else
                                If ribu = False Then
                                    string_uang = angka_ke_kata(digit_ribu(a)) + " Ratus " + "Ribu " + string_uang
                                    Exit For
                                Else
                                    string_uang = angka_ke_kata(digit_ribu(a)) + " Ratus " + string_uang
                                    Exit For
                                End If
                            End If
                        End If
                        Continue For
                    End If
                Next
            End If

            If digit_juta.Count <= 3 And digit_juta.Count > 0 Then
                Dim juta As Boolean = False
                For a As Integer = 0 To digit_juta.Count - 1
                    If digit_juta(a) <> "0" Then
                        If a = 0 Then
                            If digit_juta.Count = 1 Then
                                string_uang = angka_ke_kata(digit_juta(a)) + " Juta " + string_uang
                                Exit For
                            ElseIf digit_juta(1) <> "1" Then
                                string_uang = angka_ke_kata(digit_juta(a)) + " Juta " + string_uang
                                juta = True
                            End If
                        ElseIf a = 1 Then
                            If digit_juta(a) = "1" Then
                                If digit_juta(0) = "0" Then
                                    string_uang = "Sepuluh" + " Juta " + string_uang
                                    juta = True
                                ElseIf digit_juta(0) = "1" Then
                                    string_uang = "Sebelas" + " Juta " + string_uang
                                    juta = True
                                ElseIf digit_juta(0) <> "0" And digit_juta(0) <> "1" Then
                                    string_uang = angka_ke_kata(digit_juta(0)) + " Belas " + "Juta " + string_uang
                                    juta = True
                                End If
                            Else
                                If juta = False Then
                                    string_uang = angka_ke_kata(digit_juta(a)) + " Puluh " + "Juta " + string_uang
                                    juta = True
                                Else
                                    string_uang = angka_ke_kata(digit_juta(a)) + " Puluh " + string_uang
                                End If
                            End If
                        ElseIf a = 2 Then
                            If digit_juta(a) = "1" Then
                                If juta = False Then
                                    string_uang = "Seratus " + "Juta " + string_uang
                                    juta = True
                                    Exit For
                                Else
                                    string_uang = "Seratus " + string_uang
                                    Exit For
                                End If
                            Else
                                If juta = False Then
                                    string_uang = angka_ke_kata(digit_juta(a)) + " Ratus " + "Juta " + string_uang
                                    Exit For
                                Else
                                    string_uang = angka_ke_kata(digit_juta(a)) + " Ratus " + string_uang
                                    Exit For
                                End If
                            End If
                        End If
                        Continue For
                    End If
                Next
            End If

            If digit_miliar.Count <= 3 And digit_miliar.Count > 0 Then
                Dim miliar As Boolean = False
                For a As Integer = 0 To digit_miliar.Count - 1
                    If digit_miliar(a) <> "0" Then
                        If a = 0 Then
                            If digit_miliar.Count = 1 Then
                                string_uang = angka_ke_kata(digit_miliar(a)) + " Miliar " + string_uang
                                Exit For
                            ElseIf digit_miliar(1) <> "1" Then
                                string_uang = angka_ke_kata(digit_miliar(a)) + " Miliar " + string_uang
                                miliar = True
                            End If
                        ElseIf a = 1 Then
                            If digit_miliar(a) = "1" Then
                                If digit_miliar(0) = "0" Then
                                    string_uang = "Sepuluh" + " Miliar " + string_uang
                                    miliar = True
                                ElseIf digit_miliar(0) = "1" Then
                                    string_uang = "Sebelas" + " Miliar " + string_uang
                                    miliar = True
                                ElseIf digit_miliar(0) <> "0" And digit_miliar(0) <> "1" Then
                                    string_uang = angka_ke_kata(digit_miliar(0)) + " Belas " + "Miliar " + string_uang
                                    miliar = True
                                End If
                            Else
                                If miliar = False Then
                                    string_uang = angka_ke_kata(digit_miliar(a)) + " Puluh " + "Miliar " + string_uang
                                    miliar = True
                                Else
                                    string_uang = angka_ke_kata(digit_miliar(a)) + " Puluh " + string_uang
                                End If
                            End If
                        ElseIf a = 2 Then
                            If digit_miliar(a) = "1" Then
                                If miliar = False Then
                                    string_uang = "Seratus " + "Miliar " + string_uang
                                    miliar = True
                                    Exit For
                                Else
                                    string_uang = "Seratus " + string_uang
                                    Exit For
                                End If
                            Else
                                If miliar = False Then
                                    string_uang = angka_ke_kata(digit_miliar(a)) + " Ratus " + "Miliar " + string_uang
                                    Exit For
                                Else
                                    string_uang = angka_ke_kata(digit_miliar(a)) + " Ratus " + string_uang
                                    Exit For
                                End If
                            End If
                        End If
                        Continue For
                    End If
                Next
            End If

            If digit_triliun.Count <= 3 And digit_triliun.Count > 0 Then
                Dim triliun As Boolean = False
                For a As Integer = 0 To digit_triliun.Count - 1
                    If digit_triliun(a) <> "0" Then
                        If a = 0 Then
                            If digit_triliun.Count = 1 Then
                                string_uang = angka_ke_kata(digit_triliun(a)) + " Triliun " + string_uang
                                Exit For
                            ElseIf digit_triliun(1) <> "1" Then
                                string_uang = angka_ke_kata(digit_triliun(a)) + " Triliun " + string_uang
                                triliun = True
                            End If
                        ElseIf a = 1 Then
                            If digit_triliun(a) = "1" Then
                                If digit_triliun(0) = "0" Then
                                    string_uang = "Sepuluh" + " Triliun " + string_uang
                                    triliun = True
                                ElseIf digit_triliun(0) = "1" Then
                                    string_uang = "Sebelas" + " Triliun " + string_uang
                                    triliun = True
                                ElseIf digit_triliun(0) <> "0" And digit_triliun(0) <> "1" Then
                                    string_uang = angka_ke_kata(digit_triliun(0)) + " Belas " + "Triliun " + string_uang
                                    triliun = True
                                End If
                            Else
                                If triliun = False Then
                                    string_uang = angka_ke_kata(digit_triliun(a)) + " Puluh " + "Triliun " + string_uang
                                    triliun = True
                                Else
                                    string_uang = angka_ke_kata(digit_triliun(a)) + " Puluh " + string_uang
                                End If
                            End If
                        ElseIf a = 2 Then
                            If digit_triliun(a) = "1" Then
                                If triliun = False Then
                                    string_uang = "Seratus " + "Triliun " + string_uang
                                    triliun = True
                                    Exit For
                                Else
                                    string_uang = "Seratus " + string_uang
                                    Exit For
                                End If
                            Else
                                If triliun = False Then
                                    string_uang = angka_ke_kata(digit_triliun(a)) + " Ratus " + "Triliun " + string_uang
                                    Exit For
                                Else
                                    string_uang = angka_ke_kata(digit_triliun(a)) + " Ratus " + string_uang
                                    Exit For
                                End If
                            End If
                        End If
                        Continue For
                    End If
                Next
            End If

            'Label13.Text = MaskedTextBox1.Text

            'Label13.Text = string_uang

            Dim section As Section = Kwitansi.Sections(0)
            Dim kwt_no As Paragraph = section.Paragraphs(1)
            Dim terima As Paragraph = section.Paragraphs(2)
            Dim uang As Paragraph = section.Paragraphs(3)
            Dim pembayaran As Paragraph = section.Paragraphs(4)

            Dim nominal As TextBox = Kwitansi.TextBoxes.Item(4)
            Dim nominal_para As Paragraph = nominal.Body.FirstParagraph
            Dim nominal_angka As String = ""
            Dim count_nominal As Integer = 1

            For h As Integer = len_char_uang - 1 To 0 Step -1
                nominal_angka = TextBox2.Text(h) + nominal_angka

                If count_nominal = 3 And h <> 0 Then
                    nominal_angka = "," + nominal_angka
                    count_nominal = 0
                End If
                count_nominal += 1
            Next

            Dim tgl_hariini = Today.Day.ToString
            Dim bulan_ini = Today.Month.ToString
            Dim tahun_ini = Today.Year.ToString

            bulan_ini = Nama_Bulan(bulan_ini)

            Dim tanggal As TextBox = Kwitansi.TextBoxes.Item(3)
            Dim tanggal_para As Paragraph = tanggal.Body.FirstParagraph

            Dim nama As TextBox = Kwitansi.TextBoxes.Item(2)
            Dim nama_para As Paragraph = nama.Body.FirstParagraph

            Dim tr1 As TextRange = kwt_no.AppendText(TextBox6.Text)
            Dim tr2 As TextRange = terima.AppendText(TextBox1.Text)
            Dim tr3 As TextRange = uang.AppendText(string_uang)
            Dim tr4 As TextRange = pembayaran.AppendText(TextBox3.Text)
            Dim tr5 As TextRange = tanggal_para.AppendText(tgl_hariini + " " + bulan_ini + " " + tahun_ini)
            Dim tr6 As TextRange = nominal_para.AppendText(" " + nominal_angka)
            Dim tr7 As TextRange = nama_para.AppendText(TextBox4.Text + ")")

            tr1.CharacterFormat.FontName = "Times New Roman"
            tr2.CharacterFormat.FontName = "Times New Roman"
            tr3.CharacterFormat.FontName = "Times New Roman"
            tr4.CharacterFormat.FontName = "Times New Roman"
            tr5.CharacterFormat.FontName = "Times New Roman"
            tr6.CharacterFormat.FontName = "Times New Roman"
            tr7.CharacterFormat.FontName = "Times New Roman"

            tr1.CharacterFormat.FontSize = 14

            If TextBox1.Text.Length < 60 Then
                tr2.CharacterFormat.FontSize = 14
            ElseIf TextBox1.Text.Length < 78 Then
                tr2.CharacterFormat.FontSize = 12
            ElseIf TextBox1.Text.Length < 90 Then
                tr2.CharacterFormat.FontSize = 11
            ElseIf TextBox1.Text.Length < 97 Then
                tr2.CharacterFormat.FontSize = 10
            Else
                tr2.CharacterFormat.FontSize = 9
            End If


            If string_uang.Length > 136 Then
                tr3.CharacterFormat.FontSize = 10
            ElseIf string_uang.Length > 121 Then
                tr3.CharacterFormat.FontSize = 11
            ElseIf string_uang.Length > 96 Then
                tr3.CharacterFormat.FontSize = 12
            Else
                tr3.CharacterFormat.FontSize = 14
            End If

            tr4.CharacterFormat.FontSize = 14
            tr5.CharacterFormat.FontSize = 14
            tr6.CharacterFormat.FontSize = 18
            tr7.CharacterFormat.FontSize = 14

            tr1.CharacterFormat.TextColor = Color.Black
            tr2.CharacterFormat.TextColor = Color.Black
            tr3.CharacterFormat.TextColor = Color.Black
            tr4.CharacterFormat.TextColor = Color.Black
            tr5.CharacterFormat.TextColor = Color.Black
            tr6.CharacterFormat.TextColor = Color.Black
            tr7.CharacterFormat.TextColor = Color.Black

            tr3.CharacterFormat.Bold = True
            tr6.CharacterFormat.Bold = True

            Dim file_name As String = "Kwitansi\Kwitansi " + TextBox1.Text + " " + DateTime.Now.ToString("dd MMMM yyyy HH.mm") + ".pdf"

            Kwitansi.SaveToFile(file_name, Spire.Doc.FileFormat.PDF)
            Dim Confirm As DialogResult = MessageBox.Show("Dokumen tersimpan." & vbCrLf & "Apakah anda ingin mencetak dokumen ini dengan printer " + DataGridView2.SelectedRows(0).Cells(0).Value + "?", "Konfirmasi Pencetakan Dokumen", MessageBoxButtons.YesNo)

            If Confirm = DialogResult.Yes Then
                Dim kwitansi_print As PdfDocument = New PdfDocument()
                kwitansi_print.LoadFromFile(file_name)
                kwitansi_print.PrintSettings.PrinterName = DataGridView2.SelectedRows(0).Cells(0).Value
                kwitansi_print.PrintSettings.SelectSinglePageLayout(Print.PdfSinglePageScalingMode.ActualSize)
                Try
                    kwitansi_print.Print()
                Catch ex As Exception
                    MsgBox("Terjadi kesalahan pada pencetakan dokumen")
                End Try
            End If
        ElseIf RadioButton6.Checked = True Then
            Dim Temp As Integer
            If CheckBox1.Checked = True And CheckBox2.Checked = False Then
                If TextBox17.Text = "" Or TextBox18.Text = "" Or TextBox19.Text = "" Or TextBox14.Text = "" Or TextBox15.Text = "" Or TextBox16.Text = "" Then
                    MsgBox("Semua Field harus diisi!!!")
                    Exit Sub
                End If

                If TextBox19.Text(0) = "0" And TextBox19.Text.Count = 2 Then
                    TextBox19.Text = TextBox19.Text(1)
                End If

                If TextBox16.Text(0) = "0" And TextBox16.Text.Count = 2 Then
                    TextBox16.Text = TextBox16.Text(1)
                End If

                If TextBox18.Text(0) = "0" And TextBox18.Text.Count = 2 Then
                    TextBox18.Text = TextBox18.Text(1)
                End If

                If TextBox15.Text(0) = "0" And TextBox15.Text.Count = 2 Then
                    TextBox15.Text = TextBox15.Text(1)
                End If

                If Int32.TryParse(TextBox14.Text, Temp) = False Or Int32.TryParse(TextBox15.Text, Temp) = False Or Int32.TryParse(TextBox16.Text, Temp) = False Or Int32.TryParse(TextBox17.Text, Temp) = False Or Int32.TryParse(TextBox18.Text, Temp) = False Or Int32.TryParse(TextBox19.Text, Temp) = False Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                ElseIf TextBox14.Text.Trim().Length <> 4 Or TextBox17.Text.Trim().Length <> 4 Or TextBox15.Text.Trim().Length < 1 Or TextBox15.Text.Trim().Length > 2 Or TextBox18.Text.Trim().Length < 1 Or TextBox18.Text.Trim().Length > 2 Or TextBox16.Text.Trim().Length < 1 Or TextBox16.Text.Trim().Length > 2 Or TextBox19.Text.Trim().Length < 1 Or TextBox19.Text.Trim().Length > 2 Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                If Integer.Parse(TextBox17.Text) > Integer.Parse(TextBox14.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                ElseIf Integer.Parse(TextBox18.Text) > Integer.Parse(TextBox15.Text) And Integer.Parse(TextBox17.Text) = Integer.Parse(TextBox14.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Dim Hari_Pada_Bulan_Awal As Integer = Hari_Pada_Bulan(TextBox18.Text, Integer.Parse(TextBox17.Text))
                Dim Hari_Pada_Bulan_Akhir As Integer = Hari_Pada_Bulan(TextBox15.Text, Integer.Parse(TextBox14.Text))
                If (Integer.Parse(TextBox19.Text) > Hari_Pada_Bulan_Awal Or Integer.Parse(TextBox19.Text) <= 0) Or (Integer.Parse(TextBox16.Text) > Hari_Pada_Bulan_Akhir Or Integer.Parse(TextBox16.Text) <= 0) Then
                    If Hari_Pada_Bulan_Awal = -1 Or Hari_Pada_Bulan_Akhir = -1 Then
                        MsgBox("Tanggal Tidak Valid!!!")
                        Exit Sub
                    End If
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Dim Bulan_Awal As String = Nama_Bulan(TextBox18.Text)
                Dim Bulan_Akhir As String = Nama_Bulan(TextBox15.Text)
                If Bulan_Awal = "Nomor bulan tidak valid!!!" Or Bulan_Akhir = "Nomor bulan tidak valid!!!" Then
                    MsgBox("Nomor bulan tidak valid!!!")
                    Exit Sub
                End If

                Dim jumlah_hari As Integer = Hitung_Hari(Integer.Parse(TextBox19.Text), Integer.Parse(TextBox18.Text), Integer.Parse(TextBox17.Text), Integer.Parse(TextBox16.Text), Integer.Parse(TextBox15.Text), Integer.Parse(TextBox14.Text), False)
                If jumlah_hari <= 0 Then
                    MsgBox("Tanggal tidak valid")
                    Exit Sub
                End If

                Call Connect()
                da = New OdbcDataAdapter("Select * From tbl_pasien Where `Tanggal Input` Between '" & TextBox17.Text & "-" & TextBox18.Text & "-" & TextBox19.Text & " 00:00:00 ' And '" & TextBox14.Text & "-" & TextBox15.Text & "-" & TextBox16.Text & " 23:59:59'", conn)
                ds = New DataSet
                da.Fill(ds, "tbl_pasien")
                DataGridView1.DataSource = ds.Tables("tbl_pasien")
                conn.Close()
            ElseIf CheckBox1.Checked = False And CheckBox2.Checked = True Then
                If TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Or TextBox13.Text = "" Then
                    MsgBox("Semua Field harus diisi!!!")
                    Exit Sub
                End If

                If TextBox8.Text(0) = "0" And TextBox8.Text.Count = 2 Then
                    TextBox8.Text = TextBox8.Text(1)
                End If

                If TextBox9.Text(0) = "0" And TextBox9.Text.Count = 2 Then
                    TextBox9.Text = TextBox9.Text(1)
                End If

                If TextBox13.Text(0) = "0" And TextBox13.Text.Count = 2 Then
                    TextBox13.Text = TextBox13.Text(1)
                End If

                If TextBox12.Text(0) = "0" And TextBox12.Text.Count = 2 Then
                    TextBox12.Text = TextBox12.Text(1)
                End If

                If Int32.TryParse(TextBox8.Text, Temp) = False Or Int32.TryParse(TextBox9.Text, Temp) = False Or Int32.TryParse(TextBox10.Text, Temp) = False Or Int32.TryParse(TextBox11.Text, Temp) = False Or Int32.TryParse(TextBox12.Text, Temp) = False Or Int32.TryParse(TextBox13.Text, Temp) = False Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                ElseIf TextBox10.Text.Trim().Length <> 4 Or TextBox11.Text.Trim().Length <> 4 Or TextBox9.Text.Trim().Length < 1 Or TextBox9.Text.Trim().Length > 2 Or TextBox12.Text.Trim().Length < 1 Or TextBox12.Text.Trim().Length > 2 Or TextBox8.Text.Trim().Length < 1 Or TextBox8.Text.Trim().Length > 2 Or TextBox13.Text.Trim().Length < 1 Or TextBox13.Text.Trim().Length > 2 Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                If Integer.Parse(TextBox10.Text) > Integer.Parse(TextBox11.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                ElseIf Integer.Parse(TextBox9.Text) > Integer.Parse(TextBox12.Text) And Integer.Parse(TextBox10.Text) = Integer.Parse(TextBox11.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Dim Hari_Pada_Bulan_Awal As Integer = Hari_Pada_Bulan(TextBox9.Text, Integer.Parse(TextBox10.Text))
                Dim Hari_Pada_Bulan_Akhir As Integer = Hari_Pada_Bulan(TextBox12.Text, Integer.Parse(TextBox11.Text))
                If (Integer.Parse(TextBox8.Text) > Hari_Pada_Bulan_Awal Or Integer.Parse(TextBox8.Text) <= 0) Or (Integer.Parse(TextBox13.Text) > Hari_Pada_Bulan_Akhir Or Integer.Parse(TextBox13.Text) <= 0) Then
                    If Hari_Pada_Bulan_Awal = -1 Or Hari_Pada_Bulan_Akhir = -1 Then
                        MsgBox("Tanggal Tidak Valid!!!")
                        Exit Sub
                    End If
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Dim Bulan_Awal As String = Nama_Bulan(TextBox9.Text)
                Dim Bulan_Akhir As String = Nama_Bulan(TextBox12.Text)
                If Bulan_Awal = "Nomor bulan tidak valid!!!" Or Bulan_Akhir = "Nomor bulan tidak valid!!!" Then
                    MsgBox("Nomor bulan tidak valid!!!")
                    Exit Sub
                End If

                Dim jumlah_hari As Integer = Hitung_Hari(Integer.Parse(TextBox8.Text), Integer.Parse(TextBox9.Text), Integer.Parse(TextBox10.Text), Integer.Parse(TextBox13.Text), Integer.Parse(TextBox12.Text), Integer.Parse(TextBox11.Text), False)
                If jumlah_hari <= 0 Then
                    MsgBox("Tanggal tidak valid")
                    Exit Sub
                End If

                Call Connect()
                da = New OdbcDataAdapter("Select * From tbl_pasien Where `Update Terakhir` Between '" & TextBox10.Text & "-" & TextBox9.Text & "-" & TextBox8.Text & " 00:00:00 ' And '" & TextBox11.Text & "-" & TextBox12.Text & "-" & TextBox13.Text & " 23:59:59'", conn)
                ds = New DataSet
                da.Fill(ds, "tbl_pasien")
                DataGridView1.DataSource = ds.Tables("tbl_pasien")
                conn.Close()
            Else
                If TextBox17.Text = "" Or TextBox18.Text = "" Or TextBox19.Text = "" Or TextBox14.Text = "" Or TextBox15.Text = "" Or TextBox16.Text = "" Then
                    MsgBox("Semua Field harus diisi!!!")
                    Exit Sub
                End If

                If TextBox19.Text(0) = "0" And TextBox19.Text.Count = 2 Then
                    TextBox19.Text = TextBox19.Text(1)
                End If

                If TextBox16.Text(0) = "0" And TextBox16.Text.Count = 2 Then
                    TextBox16.Text = TextBox16.Text(1)
                End If

                If TextBox18.Text(0) = "0" And TextBox18.Text.Count = 2 Then
                    TextBox18.Text = TextBox18.Text(1)
                End If

                If TextBox15.Text(0) = "0" And TextBox15.Text.Count = 2 Then
                    TextBox15.Text = TextBox15.Text(1)
                End If

                If Int32.TryParse(TextBox14.Text, Temp) = False Or Int32.TryParse(TextBox15.Text, Temp) = False Or Int32.TryParse(TextBox16.Text, Temp) = False Or Int32.TryParse(TextBox17.Text, Temp) = False Or Int32.TryParse(TextBox18.Text, Temp) = False Or Int32.TryParse(TextBox19.Text, Temp) = False Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                ElseIf TextBox14.Text.Trim().Length <> 4 Or TextBox17.Text.Trim().Length <> 4 Or TextBox15.Text.Trim().Length < 1 Or TextBox15.Text.Trim().Length > 2 Or TextBox18.Text.Trim().Length < 1 Or TextBox18.Text.Trim().Length > 2 Or TextBox16.Text.Trim().Length < 1 Or TextBox16.Text.Trim().Length > 2 Or TextBox19.Text.Trim().Length < 1 Or TextBox19.Text.Trim().Length > 2 Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                If Integer.Parse(TextBox17.Text) > Integer.Parse(TextBox14.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                ElseIf Integer.Parse(TextBox18.Text) > Integer.Parse(TextBox15.Text) And Integer.Parse(TextBox17.Text) = Integer.Parse(TextBox14.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Dim Hari_Pada_Bulan_Awal As Integer = Hari_Pada_Bulan(TextBox18.Text, Integer.Parse(TextBox17.Text))
                Dim Hari_Pada_Bulan_Akhir As Integer = Hari_Pada_Bulan(TextBox15.Text, Integer.Parse(TextBox14.Text))
                If (Integer.Parse(TextBox19.Text) > Hari_Pada_Bulan_Awal Or Integer.Parse(TextBox19.Text) <= 0) Or (Integer.Parse(TextBox16.Text) > Hari_Pada_Bulan_Akhir Or Integer.Parse(TextBox16.Text) <= 0) Then
                    If Hari_Pada_Bulan_Awal = -1 Or Hari_Pada_Bulan_Akhir = -1 Then
                        MsgBox("Tanggal Tidak Valid!!!")
                        Exit Sub
                    End If
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Dim Bulan_Awal As String = Nama_Bulan(TextBox18.Text)
                Dim Bulan_Akhir As String = Nama_Bulan(TextBox15.Text)
                If Bulan_Awal = "Nomor bulan tidak valid!!!" Or Bulan_Akhir = "Nomor bulan tidak valid!!!" Then
                    MsgBox("Nomor bulan tidak valid!!!")
                    Exit Sub
                End If

                Dim jumlah_hari As Integer = Hitung_Hari(Integer.Parse(TextBox19.Text), Integer.Parse(TextBox18.Text), Integer.Parse(TextBox17.Text), Integer.Parse(TextBox16.Text), Integer.Parse(TextBox15.Text), Integer.Parse(TextBox14.Text), False)
                If jumlah_hari <= 0 Then
                    MsgBox("Tanggal tidak valid")
                    Exit Sub
                End If

                If TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Or TextBox13.Text = "" Then
                    MsgBox("Semua Field harus diisi!!!")
                    Exit Sub
                End If

                If TextBox8.Text(0) = "0" And TextBox8.Text.Count = 2 Then
                    TextBox8.Text = TextBox8.Text(1)
                End If

                If TextBox9.Text(0) = "0" And TextBox9.Text.Count = 2 Then
                    TextBox9.Text = TextBox9.Text(1)
                End If

                If TextBox13.Text(0) = "0" And TextBox13.Text.Count = 2 Then
                    TextBox13.Text = TextBox13.Text(1)
                End If

                If TextBox12.Text(0) = "0" And TextBox12.Text.Count = 2 Then
                    TextBox12.Text = TextBox12.Text(1)
                End If

                If Int32.TryParse(TextBox8.Text, Temp) = False Or Int32.TryParse(TextBox9.Text, Temp) = False Or Int32.TryParse(TextBox10.Text, Temp) = False Or Int32.TryParse(TextBox11.Text, Temp) = False Or Int32.TryParse(TextBox12.Text, Temp) = False Or Int32.TryParse(TextBox13.Text, Temp) = False Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                ElseIf TextBox10.Text.Trim().Length <> 4 Or TextBox11.Text.Trim().Length <> 4 Or TextBox9.Text.Trim().Length < 1 Or TextBox9.Text.Trim().Length > 2 Or TextBox12.Text.Trim().Length < 1 Or TextBox12.Text.Trim().Length > 2 Or TextBox8.Text.Trim().Length < 1 Or TextBox8.Text.Trim().Length > 2 Or TextBox13.Text.Trim().Length < 1 Or TextBox13.Text.Trim().Length > 2 Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                If Integer.Parse(TextBox10.Text) > Integer.Parse(TextBox11.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                ElseIf Integer.Parse(TextBox9.Text) > Integer.Parse(TextBox12.Text) And Integer.Parse(TextBox10.Text) = Integer.Parse(TextBox11.Text) Then
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Hari_Pada_Bulan_Awal = Hari_Pada_Bulan(TextBox9.Text, Integer.Parse(TextBox10.Text))
                Hari_Pada_Bulan_Akhir = Hari_Pada_Bulan(TextBox12.Text, Integer.Parse(TextBox11.Text))
                If (Integer.Parse(TextBox8.Text) > Hari_Pada_Bulan_Awal Or Integer.Parse(TextBox8.Text) <= 0) Or (Integer.Parse(TextBox13.Text) > Hari_Pada_Bulan_Akhir Or Integer.Parse(TextBox13.Text) <= 0) Then
                    If Hari_Pada_Bulan_Awal = -1 Or Hari_Pada_Bulan_Akhir = -1 Then
                        MsgBox("Tanggal Tidak Valid!!!")
                        Exit Sub
                    End If
                    MsgBox("Tanggal Tidak Valid!!!")
                    Exit Sub
                End If

                Bulan_Awal = Nama_Bulan(TextBox9.Text)
                Bulan_Akhir = Nama_Bulan(TextBox12.Text)
                If Bulan_Awal = "Nomor bulan tidak valid!!!" Or Bulan_Akhir = "Nomor bulan tidak valid!!!" Then
                    MsgBox("Nomor bulan tidak valid!!!")
                    Exit Sub
                End If

                jumlah_hari = Hitung_Hari(Integer.Parse(TextBox8.Text), Integer.Parse(TextBox9.Text), Integer.Parse(TextBox10.Text), Integer.Parse(TextBox13.Text), Integer.Parse(TextBox12.Text), Integer.Parse(TextBox11.Text), False)
                If jumlah_hari <= 0 Then
                    MsgBox("Tanggal tidak valid")
                    Exit Sub
                End If

                Call Connect()
                da = New OdbcDataAdapter("Select * From tbl_pasien Where (`Tanggal Input` Between '" & TextBox17.Text & "-" & TextBox18.Text & "-" & TextBox19.Text & " 00:00:00 ' And '" & TextBox14.Text & "-" & TextBox15.Text & "-" & TextBox16.Text & " 23:59:59') Or (`Update Terakhir` Between '" & TextBox10.Text & "-" & TextBox9.Text & "-" & TextBox8.Text & " 00:00:00 ' And '" & TextBox11.Text & "-" & TextBox12.Text & "-" & TextBox13.Text & " 23:59:59')", conn)
                ds = New DataSet
                da.Fill(ds, "tbl_pasien")
                DataGridView1.DataSource = ds.Tables("tbl_pasien")
                conn.Close()
            End If
        End If
    End Sub

    Function angka_ke_kata(i As String) As String
        Select Case i
            Case Is = "1"
                Return "Satu"
            Case Is = "2"
                Return "Dua"
            Case Is = "3"
                Return "Tiga"
            Case Is = "4"
                Return "Empat"
            Case Is = "5"
                Return "Lima"
            Case Is = "6"
                Return "Enam"
            Case Is = "7"
                Return "Tujuh"
            Case Is = "8"
                Return "Delapan"
            Case Is = "9"
                Return "Sembilan"
            Case Else
                Return "Something's wrong"
        End Select
    End Function

    Function Hari_Pada_Bulan(i As String, tahun As Integer) As Integer
        Select Case i
            Case Is = "01" Or "1"
                Return 31
            Case Is = "02" Or "2"
                If tahun Mod 4 = 0 Then
                    Return 29
                End If
                Return 28
            Case Is = "03" Or "3"
                Return 31
            Case Is = "04" Or "4"
                Return 30
            Case Is = "05" Or "5"
                Return 31
            Case Is = "06" Or "6"
                Return 30
            Case Is = "07" Or "7"
                Return 31
            Case Is = "08" Or "8"
                Return 31
            Case Is = "09" Or "9"
                Return 30
            Case Is = "10"
                Return 31
            Case Is = "11"
                Return 30
            Case Is = "12"
                Return 31
            Case Else
                Return -1
        End Select
    End Function

    Function Nama_Bulan(i As String) As String
        Select Case i
            Case Is = "01" Or "1"
                Return "Januari"
            Case Is = "02" Or "2"
                Return "Februari"
            Case Is = "03" Or "3"
                Return "Maret"
            Case Is = "04" Or "4"
                Return "April"
            Case Is = "05" Or "5"
                Return "Mei"
            Case Is = "06" Or "6"
                Return "Juni"
            Case Is = "07" Or "7"
                Return "Juli"
            Case Is = "08" Or "8"
                Return "Agustus"
            Case Is = "09" Or "9"
                Return "September"
            Case Is = "10"
                Return "Oktober"
            Case Is = "11"
                Return "November"
            Case Is = "12"
                Return "Desember"
            Case Else
                Return "Nomor bulan tidak valid!!!"
        End Select
    End Function

    Function Hitung_Hari(tgl_awal As Integer, bulan_awal As Integer, tahun_awal As Integer, tgl_akhir As Integer, bulan_akhir As Integer, tahun_akhir As Integer, limit As Boolean) As Integer
        If (bulan_awal = bulan_akhir) And (tahun_awal = tahun_akhir) Then
            Return tgl_akhir - tgl_awal + 1
        ElseIf tahun_awal = tahun_akhir Then
            Dim selisih_bulan As Integer = bulan_akhir - bulan_awal
            If selisih_bulan > 1 Then
                Return -1
            Else
                Dim jumlah_hari As Integer
                jumlah_hari += Hari_Pada_Bulan(CStr(bulan_awal), tahun_awal)
                jumlah_hari = jumlah_hari - tgl_awal + 1
                jumlah_hari += tgl_akhir
                Return jumlah_hari
            End If
            'Dim i As Integer = 1
            'Dim jumlah_hari As Integer
            'For i = 1 To selisih_bulan
            '    jumlah_hari += Hari_Pada_Bulan(CStr(bulan_awal), tahun_awal)
            '    If i = 1 Then
            '        jumlah_hari = jumlah_hari - tgl_awal + 1
            '    ElseIf i = selisih_bulan Then
            '        jumlah_hari += tgl_akhir
            '    End If
            '    bulan_awal += 1
            'Next
            'Return jumlah_hari
        ElseIf tahun_awal > tahun_akhir Then
            Return -1
        ElseIf tahun_awal < tahun_akhir Then
            Dim selisih_tahun As Integer = tahun_akhir - tahun_awal
            If limit = True Then
                If selisih_tahun > 1 Then
                    Return -1
                ElseIf bulan_akhir + 12 - bulan_awal > 1 Then
                    Return -1
                Else
                    Dim jumlah_hari As Integer
                    jumlah_hari = 31 - tgl_awal + 1
                    jumlah_hari += tgl_akhir
                    Return jumlah_hari
                End If
            Else
                Return 1
            End If
        End If
    End Function

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If RadioButton4.Checked = True Then
            If e.KeyChar = Chr(13) Then
                Call Connect()
                cmd = New OdbcCommand("Select * From tbl_pasien Where Id = '" & TextBox6.Text & "'", conn)
                rd = cmd.ExecuteReader
                If rd.HasRows Then
                    TextBox1.Text = rd.Item("Nama")
                    TextBox2.Text = rd.Item("Umur")
                    TextBox3.Text = rd.Item("Alamat")
                    TextBox4.Text = rd.Item("No. Hp")
                    RichTextBox1.Text = rd.Item("Tindakan")
                End If
                da = New OdbcDataAdapter("Select * From tbl_pasien Where Id = '" & TextBox6.Text & "'", conn)
                ds = New DataSet
                da.Fill(ds, "tbl_pasien")
                DataGridView1.DataSource = ds.Tables("tbl_pasien")
                conn.Close()
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If RadioButton1.Checked = False Then
            If e.KeyChar = Chr(13) Then
                Call Connect()
                Call Enter_press_on_TextBox()
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If RadioButton1.Checked = False Then
            If e.KeyChar = Chr(13) Then
                Call Connect()
                Call Enter_press_on_TextBox()
            End If
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If RadioButton1.Checked = False Then
            If e.KeyChar = Chr(13) Then
                Call Connect()
                Call Enter_press_on_TextBox()
            End If
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If RadioButton1.Checked = False Then
            If e.KeyChar = Chr(13) Then
                Call Connect()
                Call Enter_press_on_TextBox()
            End If
        End If
    End Sub

    'Private Sub RichTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RichTextBox1.KeyPress
    '    If RadioButton1.Checked = False Then
    '        If e.KeyChar = Chr(13) Then
    '            Call Connect()
    '            Call Enter_press_on_TextBox()
    '        End If
    '    End If
    'End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Button3.Enabled = True
        Dim Selected_Row_Count As Integer = DataGridView1.SelectedRows.Count()
        If Selected_Row_Count >= 1 Then
            Dim Data_Id As String = DataGridView1.SelectedRows(0).Cells(0).Value
            Dim Data_Nama As String = DataGridView1.SelectedRows(0).Cells(1).Value
            Dim Data_Umur As String = DataGridView1.SelectedRows(0).Cells(2).Value.ToString
            Dim Data_Alamat As String = DataGridView1.SelectedRows(0).Cells(3).Value
            Dim Data_Hp As String = DataGridView1.SelectedRows(0).Cells(4).Value
            Dim Data_TglInput As String = DataGridView1.SelectedRows(0).Cells(5).Value
            Dim Data_TglUpdate As String = DataGridView1.SelectedRows(0).Cells(6).Value
            Dim Data_Tindakan As String = DataGridView1.SelectedRows(0).Cells(7).Value

            Dim Data_TglInput_Tahun As String = Data_TglInput(6) + Data_TglInput(7) + Data_TglInput(8) + Data_TglInput(9)
            Dim Data_TglInput_Bulan As String = Data_TglInput(3) + Data_TglInput(4)
            Dim Data_TglInput_Tanggal As String = Data_TglInput(0) + Data_TglInput(1)

            Dim Data_TglUpdate_Tahun As String = Data_TglUpdate(6) + Data_TglUpdate(7) + Data_TglUpdate(8) + Data_TglUpdate(9)
            Dim Data_TglUpdate_Bulan As String = Data_TglUpdate(3) + Data_TglUpdate(4)
            Dim Data_TglUpdate_Tanggal As String = Data_TglUpdate(0) + Data_TglUpdate(1)

            If Not (RadioButton1.Checked = False Xor RadioButton5.Checked = False) Then
                TextBox6.Text = Data_Id
                TextBox1.Text = Data_Nama
                TextBox2.Text = Data_Umur
                TextBox3.Text = Data_Alamat
                TextBox4.Text = Data_Hp
                TextBox5.Text = Data_TglInput
                TextBox7.Text = Data_TglUpdate
                If RadioButton6.Checked = False Then
                    TextBox8.Text = Data_TglInput_Tanggal
                    TextBox9.Text = Data_TglInput_Bulan
                    TextBox10.Text = Data_TglInput_Tahun
                    TextBox13.Text = Data_TglUpdate_Tanggal
                    TextBox12.Text = Data_TglUpdate_Bulan
                    TextBox11.Text = Data_TglUpdate_Tahun
                End If
                RichTextBox1.Text = Data_Tindakan
                End If
            End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        RichTextBox1.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim Selected_Row_Count As Integer = DataGridView1.SelectedRows.Count()
        Dim Confirm As DialogResult = MessageBox.Show("Apakah anda yakin ingin menghapus data pasien ini? Semua riwayat tindakan pasien ini akan dihapus juga.", "Konfirmasi Penghapusan Data", MessageBoxButtons.YesNo)
        If Selected_Row_Count >= 1 Then
            If Confirm = DialogResult.Yes Then
                Dim Data_Id As String = DataGridView1.SelectedRows(0).Cells(0).Value
                Call Connect()
                Dim HapusData As String = "Delete From tbl_tindakan Where Id_pasien = '" & Data_Id & "'"
                cmd = New OdbcCommand(HapusData, conn)
                cmd.ExecuteNonQuery()

                HapusData = "Delete From tbl_pasien Where Id = '" & Data_Id & "'"
                cmd = New OdbcCommand(HapusData, conn)
                cmd.ExecuteNonQuery()

                MsgBox("Hapus Data Berhasil")
                conn.Close()
                Call Reset_textbox()
                Call Connect()
                da = New OdbcDataAdapter("Select * From tbl_pasien", conn)
                ds = New DataSet
                da.Fill(ds, "tbl_pasien")
                DataGridView1.DataSource = ds.Tables("tbl_pasien")
                conn.Close()
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DataGridView3.Visible = False
        Call Connect()
        da = New OdbcDataAdapter("Select * From tbl_pasien", conn)
        ds = New DataSet
        da.Fill(ds, "tbl_pasien")
        DataGridView1.DataSource = ds.Tables("tbl_pasien")
        conn.Close()
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If RadioButton4.Checked = True Then
            If e.KeyChar = Chr(13) Then
                Call Enter_press_on_TextBox()
                'If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
                '    MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
                'ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
                '    If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                '        MsgBox("Input tanggal tidak valid")
                '    Else
                '
                '    End If
                'End If
            End If
        End If
    End Sub

    Private Sub DataGridView1_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles DataGridView1.DataBindingComplete
        DataGridView1.ClearSelection()
        Call Reset_textbox()
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If RadioButton4.Checked = True Then
            If e.KeyChar = Chr(13) Then
                Call Enter_press_on_TextBox()
                'If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
                '    MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
                'ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
                '    If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                '        MsgBox("Input tanggal tidak valid")
                '    Else
                '
                '    End If
                'End If
            End If
        End If
    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        If RadioButton4.Checked = True Then
            If e.KeyChar = Chr(13) Then
                Call Enter_press_on_TextBox()
                'If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
                '    MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
                'ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
                '    If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                '        MsgBox("Input tanggal tidak valid")
                '    Else
                '
                '    End If
                'End If
            End If
        End If
    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox13.KeyPress
        If RadioButton4.Checked = True Then
            If e.KeyChar = Chr(13) Then
                Call Enter_press_on_TextBox()
                'If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
                '    MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
                'ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
                '    If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                '        MsgBox("Input tanggal tidak valid")
                '    Else
                '
                '    End If
                'End If
            End If
        End If
    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        If RadioButton4.Checked = True Then
            If e.KeyChar = Chr(13) Then
                Call Enter_press_on_TextBox()
                'If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
                '    MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
                'ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
                '    If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                '        MsgBox("Input tanggal tidak valid")
                '    Else
                '
                '    End If
                'End If
            End If
        End If
    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        If RadioButton4.Checked = True Then
            If e.KeyChar = Chr(13) Then
                Call Enter_press_on_TextBox()
                'If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
                '    MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
                'ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
                '    If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                '        MsgBox("Input tanggal tidak valid")
                '    Else
                '
                '    End If
                'End If
            End If
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        DataGridView1.ClearSelection()
        Call Reset_textbox()
        TextBox6.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        TextBox7.Visible = False
        TextBox8.Visible = False
        TextBox9.Visible = False
        TextBox10.Visible = False
        TextBox11.Visible = False
        TextBox12.Visible = False
        TextBox13.Visible = False

        TextBox1.Visible = True
        TextBox2.Visible = True

        DataGridView2.Visible = True

        TextBox2.Font = New Font(TextBox2.Font, FontStyle.Regular)

        Label1.Visible = True
        Label2.Visible = True

        Label5.Visible = True
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False

        Label1.Text = "Nama"
        Label2.Text = "Umur"
        Label3.Text = "Tanggal Awal"
        Label4.Text = "Tanggal Akhir"
        Label5.Text = "Printer"

        TextBox14.Visible = True
        TextBox15.Visible = True
        TextBox16.Visible = True
        TextBox17.Visible = True
        TextBox18.Visible = True
        TextBox19.Visible = True

        TextBox14.Enabled = True
        TextBox15.Enabled = True
        TextBox16.Enabled = True
        TextBox17.Enabled = True
        TextBox18.Enabled = True
        TextBox19.Enabled = True

        Label14.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label17.Visible = True

        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False

        RichTextBox1.Visible = False

        Button1.Text = "Save and Print"

        CheckBox1.Visible = False
        CheckBox2.Visible = False
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        DataGridView1.ClearSelection()
        Call Reset_textbox()
        TextBox6.Enabled = True
        TextBox6.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = True
        TextBox5.Visible = False
        TextBox5.Enabled = True
        TextBox7.Visible = False
        TextBox8.Visible = False
        TextBox9.Visible = False
        TextBox10.Visible = False
        TextBox11.Visible = False
        TextBox12.Visible = False
        TextBox13.Visible = False

        TextBox1.Visible = True
        TextBox2.Visible = True

        DataGridView2.Visible = True

        TextBox2.Font = New Font(TextBox2.Font, FontStyle.Bold)

        Label1.Visible = True
        Label2.Visible = True

        Label5.Visible = True
        Label6.Visible = True
        Label7.Visible = False
        Label8.Visible = False

        Label6.Text = "Kwt No"
        Label1.Text = "Sudah Terima Dari"
        Label2.Text = "Banyaknya Uang"
        Label3.Text = "Untuk Pembayaran"
        Label4.Text = "Nama TTD"
        Label5.Text = "Printer"

        TextBox14.Visible = False
        TextBox15.Visible = False
        TextBox16.Visible = False
        TextBox17.Visible = False
        TextBox18.Visible = False
        TextBox19.Visible = False

        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False

        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label12.Visible = False

        RichTextBox1.Visible = False

        Button1.Text = "Save and Print"

        CheckBox1.Visible = False
        CheckBox2.Visible = False
    End Sub

    Dim Daftar_Pasien As Boolean = True
    Dim Riwayat_Tindakan As Boolean = False

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Daftar_Pasien = True Then
            If DataGridView1.SelectedRows.Count() > 0 Then
                Call Connect()
                da = New OdbcDataAdapter("Select Id_tindakan, Nama, Tanggal, Tindakan From tbl_tindakan Where Id_pasien = '" & DataGridView1.SelectedRows(0).Cells(0).Value & "'", conn)
                ds = New DataSet
                da.Fill(ds, "tbl_tindakan")
                DataGridView3.DataSource = ds.Tables("tbl_tindakan")
                conn.Close()

                DataGridView3.Visible = True

                Daftar_Pasien = False
                Riwayat_Tindakan = True
                Button5.Text = "Daftar Pasien"
                Button3.Enabled = False
            End If
        ElseIf Riwayat_Tindakan = True Then
            DataGridView3.Visible = False

            Daftar_Pasien = True
            Riwayat_Tindakan = False
            Button5.Text = "Riwayat Tindakan"
            Button3.Enabled = True
        End If
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        DataGridView1.ClearSelection()
        Call Reset_textbox()
        TextBox6.Enabled = True
        TextBox6.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        TextBox5.Enabled = False
        TextBox7.Visible = False
        TextBox8.Visible = True
        TextBox9.Visible = True
        TextBox10.Visible = True
        TextBox11.Visible = True
        TextBox12.Visible = True
        TextBox13.Visible = True

        TextBox1.Visible = False
        TextBox2.Visible = False

        DataGridView2.Visible = False

        TextBox2.Font = New Font(TextBox2.Font, FontStyle.Regular)

        Label1.Visible = False
        Label2.Visible = False

        Label5.Visible = True
        Label6.Visible = False
        Label7.Visible = True
        Label8.Visible = False

        Label3.Text = "Input Dari"
        Label4.Text = "Input Sampai"
        Label5.Text = "Update Dari"
        Label7.Text = "Update Sampai"

        TextBox14.Visible = True
        TextBox15.Visible = True
        TextBox16.Visible = True
        TextBox17.Visible = True
        TextBox18.Visible = True
        TextBox19.Visible = True

        Label14.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label17.Visible = True

        Label9.Visible = True
        Label10.Visible = True
        Label11.Visible = True
        Label12.Visible = True

        RichTextBox1.Visible = False

        Button1.Text = "Cari Pasien"

        CheckBox1.Visible = True
        CheckBox2.Visible = True

        If CheckBox1.Checked = True Then
            TextBox14.Enabled = True
            TextBox15.Enabled = True
            TextBox16.Enabled = True
            TextBox17.Enabled = True
            TextBox18.Enabled = True
            TextBox19.Enabled = True
        Else
            TextBox14.Enabled = False
            TextBox15.Enabled = False
            TextBox16.Enabled = False
            TextBox17.Enabled = False
            TextBox18.Enabled = False
            TextBox19.Enabled = False
        End If

        If CheckBox2.Checked = True Then
            TextBox8.Enabled = True
            TextBox9.Enabled = True
            TextBox10.Enabled = True
            TextBox11.Enabled = True
            TextBox12.Enabled = True
            TextBox13.Enabled = True
        Else
            TextBox8.Enabled = False
            TextBox9.Enabled = False
            TextBox10.Enabled = False
            TextBox11.Enabled = False
            TextBox12.Enabled = False
            TextBox13.Enabled = False
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox14.Enabled = True
            TextBox15.Enabled = True
            TextBox16.Enabled = True
            TextBox17.Enabled = True
            TextBox18.Enabled = True
            TextBox19.Enabled = True
        Else
            TextBox14.Enabled = False
            TextBox15.Enabled = False
            TextBox16.Enabled = False
            TextBox17.Enabled = False
            TextBox18.Enabled = False
            TextBox19.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox8.Enabled = True
            TextBox9.Enabled = True
            TextBox10.Enabled = True
            TextBox11.Enabled = True
            TextBox12.Enabled = True
            TextBox13.Enabled = True
        Else
            TextBox8.Enabled = False
            TextBox9.Enabled = False
            TextBox10.Enabled = False
            TextBox11.Enabled = False
            TextBox12.Enabled = False
            TextBox13.Enabled = False
        End If
    End Sub

    'Private Sub TextBox14_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox14.KeyPress
    '    If e.KeyChar = Chr(13) Then
    '        If TextBox14.Text = "" Or TextBox15.Text = "" Or TextBox16.Text = "" Then
    '            MsgBox("Tahun, bulan, dan hari harus terisi semua atau dikosongkan semua")
    '        ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
    '            If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
    '                MsgBox("Input tanggal tidak valid")
    '            Else
    '                Call Enter_press_on_TextBox()
    '            End If
    '        End If
    '    End If
    'End Sub
End Class
