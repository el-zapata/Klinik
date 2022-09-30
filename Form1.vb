Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar
Imports System.Data.Odbc
Imports System.Data.Common

Public Class Form1

    Dim conn As OdbcConnection
    Dim cmd As OdbcCommand
    Dim ds As DataSet
    Dim da As OdbcDataAdapter
    Dim rd As OdbcDataReader
    Dim MyDB As String

    Sub Connect()
        MyDB = "Driver={MySQL ODBC 8.0 Unicode Driver};Database=db_pasien;Server=localhost;uid=root"
        conn = New OdbcConnection(MyDB)
        If conn.State = ConnectionState.Closed Then conn.Open()
    End Sub

    Sub Reset_textbox()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        RichTextBox1.Text = ""
    End Sub

    Sub Enter_press_on_TextBox()
        Call Connect()
        Dim Query_String As String = "Select * From tbl_pasien Where "
        Dim Add_String As String
        'Dim Other_Params As Boolean = False
        Dim First_Params_Found As Boolean = False

        If TextBox1.Text <> "" Then
            First_Params_Found = True
            Add_String = "Nama = '" & TextBox1.Text & "' "
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
            Add_String = "Alamat = '" & TextBox3.Text & "' "
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
            Add_String = "`No. Hp` = '" & TextBox4.Text & "' "
            Query_String = Query_String + Add_String
            If TextBox8.Text <> "" And TextBox9.Text <> "" And TextBox10.Text <> "" Then
                Query_String = Query_String
            ElseIf TextBox8.Text = "" And TextBox9.Text = "" And TextBox10.Text = "" Then
                Query_String = Query_String
            Else
                MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
                Exit Sub
            End If
        End If

        If TextBox8.Text <> "" And TextBox9.Text <> "" And TextBox10.Text <> "" Then
            If TextBox8.Text.Trim().Length() < 1 Or TextBox8.Text.Trim().Length() > 2 Or TextBox9.Text.Trim().Length() < 1 Or TextBox9.Text.Trim().Length() > 2 Or TextBox10.Text.Trim().Length() <> 4 Then
                MsgBox("Input tanggal tidak valid")
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
            ElseIf TextBox13.Text = "" And TextBox12.Text = "" And TextBox11.Text = "" Then
                Query_String = Query_String
            Else
                MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
                Exit Sub
            End If
        ElseIf TextBox8.Text = "" And TextBox9.Text = "" And TextBox10.Text = "" Then
            Query_String = Query_String
        Else
            MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
            Exit Sub
        End If

        If TextBox13.Text <> "" And TextBox12.Text <> "" And TextBox11.Text <> "" Then
            If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                MsgBox("Input tanggal tidak valid")
                Exit Sub
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
            MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
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
            Add_String = "Tindakan = '" & RichTextBox1.Text & "' "
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
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs)
        Button1.Text = "Hapus"
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
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

        Button1.Visible = True
        Button1.Text = "Perbarui"
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

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        RichTextBox1.Text = ""

        Button1.Visible = True
        Button1.Text = "Masukkan ke Database"
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
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


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Connect()
        da = New OdbcDataAdapter("Select * From tbl_pasien", conn)
        ds = New DataSet
        da.Fill(ds, "tbl_pasien")
        DataGridView1.DataSource = ds.Tables("tbl_pasien")
        conn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Umur As Integer
        If RadioButton1.Checked = True Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
                MsgBox("Semua field harus diisi!!!")
            ElseIf Int32.TryParse(TextBox2.Text, Umur) = False Then
                MsgBox("Field Umur harus diisi dengan angka")
            Else
                Dim Confirm As DialogResult = MessageBox.Show("Apakah anda yakin ingin menambah data pasien ini?", "Konfirmasi Penambahan Data", MessageBoxButtons.YesNo)
                If Confirm = DialogResult.Yes Then
                    Call Connect()
                    Dim InputData As String = "Insert into tbl_pasien (Nama, Umur, Alamat, `No. Hp`, Tindakan) Values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & RichTextBox1.Text & "')"
                    cmd = New OdbcCommand(InputData, conn)
                    cmd.ExecuteNonQuery()
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
            ElseIf Int32.TryParse(TextBox2.Text, Umur) = False Then
                MsgBox("Field Umur harus diisi dengan angka")
            Else
                Dim Confirm As DialogResult = MessageBox.Show("Apakah anda yakin ingin memperbarui data pasien ini?", "Konfirmasi Perbaruan Data", MessageBoxButtons.YesNo)
                If Confirm = DialogResult.Yes Then
                    Call Connect()
                    Dim EditData As String = "Update tbl_pasien set Nama = '" & TextBox1.Text & "', Umur = '" & TextBox2.Text & "', Alamat = '" & TextBox3.Text & "', `No. Hp` = '" & TextBox4.Text & "', `Update Terakhir` = CURRENT_TIMESTAMP, Tindakan = '" & RichTextBox1.Text & "' Where Id = '" & TextBox6.Text & "'"
                    cmd = New OdbcCommand(EditData, conn)
                    cmd.ExecuteNonQuery()
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
        End If
    End Sub

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

    Private Sub RichTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RichTextBox1.KeyPress
        If RadioButton1.Checked = False Then
            If e.KeyChar = Chr(13) Then
                Call Connect()
                Call Enter_press_on_TextBox()
            End If
        End If
    End Sub

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

            If RadioButton1.Checked = False Then
                TextBox6.Text = Data_Id
                TextBox1.Text = Data_Nama
                TextBox2.Text = Data_Umur
                TextBox3.Text = Data_Alamat
                TextBox4.Text = Data_Hp
                TextBox5.Text = Data_TglInput
                TextBox7.Text = Data_TglUpdate
                TextBox8.Text = Data_TglInput_Tanggal
                TextBox9.Text = Data_TglInput_Bulan
                TextBox10.Text = Data_TglInput_Tahun
                TextBox13.Text = Data_TglUpdate_Tanggal
                TextBox12.Text = Data_TglUpdate_Bulan
                TextBox11.Text = Data_TglUpdate_Tahun
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
        RichTextBox1.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim Selected_Row_Count As Integer = DataGridView1.SelectedRows.Count()
        Dim Confirm As DialogResult = MessageBox.Show("Apakah anda yakin ingin menghapus data pasien ini?", "Konfirmasi Penghapusan Data", MessageBoxButtons.YesNo)
        If Selected_Row_Count >= 1 Then
            If Confirm = DialogResult.Yes Then
                Dim Data_Id As String = DataGridView1.SelectedRows(0).Cells(0).Value
                Call Connect()
                Dim HapusData As String = "Delete From tbl_pasien Where Id = '" & Data_Id & "'"
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
        Call Connect()
        da = New OdbcDataAdapter("Select * From tbl_pasien", conn)
        ds = New DataSet
        da.Fill(ds, "tbl_pasien")
        DataGridView1.DataSource = ds.Tables("tbl_pasien")
        conn.Close()
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If e.KeyChar = Chr(13) Then
            If TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Then
                MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
            ElseIf IsNumeric(TextBox8.Text) And IsNumeric(TextBox9.Text) And IsNumeric(TextBox10.Text) Then
                If TextBox8.Text.Trim().Length() < 1 Or TextBox8.Text.Trim().Length() > 2 Or TextBox9.Text.Trim().Length() < 1 Or TextBox9.Text.Trim().Length() > 2 Or TextBox10.Text.Trim().Length() <> 4 Then
                    MsgBox("Input tanggal tidak valid")
                Else
                    Call Enter_press_on_TextBox()
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView1_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles DataGridView1.DataBindingComplete
        DataGridView1.ClearSelection()
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If e.KeyChar = Chr(13) Then
            If TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Then
                MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
            ElseIf IsNumeric(TextBox8.Text) And IsNumeric(TextBox9.Text) And IsNumeric(TextBox10.Text) Then
                If TextBox8.Text.Trim().Length() < 1 Or TextBox8.Text.Trim().Length() > 2 Or TextBox9.Text.Trim().Length() < 1 Or TextBox9.Text.Trim().Length() > 2 Or TextBox10.Text.Trim().Length() <> 4 Then
                    MsgBox("Input tanggal tidak valid")
                Else
                    Call Enter_press_on_TextBox()
                End If
            End If
        End If
    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        If e.KeyChar = Chr(13) Then
            If TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Then
                MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
            ElseIf IsNumeric(TextBox8.Text) And IsNumeric(TextBox9.Text) And IsNumeric(TextBox10.Text) Then
                If TextBox8.Text.Trim().Length() < 1 Or TextBox8.Text.Trim().Length() > 2 Or TextBox9.Text.Trim().Length() < 1 Or TextBox9.Text.Trim().Length() > 2 Or TextBox10.Text.Trim().Length() <> 4 Then
                    MsgBox("Input tanggal tidak valid")
                Else
                    Call Enter_press_on_TextBox()
                End If
            End If
        End If
    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox13.KeyPress
        If e.KeyChar = Chr(13) Then
            If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
                MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
            ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
                If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                    MsgBox("Input tanggal tidak valid")
                Else
                    Call Enter_press_on_TextBox()
                End If
            End If
        End If
    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        If e.KeyChar = Chr(13) Then
            If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
                MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
            ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
                If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                    MsgBox("Input tanggal tidak valid")
                Else
                    Call Enter_press_on_TextBox()
                End If
            End If
        End If
    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        If e.KeyChar = Chr(13) Then
            If TextBox13.Text = "" Or TextBox12.Text = "" Or TextBox11.Text = "" Then
                MsgBox("Tanggal, bulan, dan hari harus terisi semua atau dikosongkan semua")
            ElseIf IsNumeric(TextBox13.Text) And IsNumeric(TextBox12.Text) And IsNumeric(TextBox11.Text) Then
                If TextBox13.Text.Trim().Length() < 1 Or TextBox13.Text.Trim().Length() > 2 Or TextBox12.Text.Trim().Length() < 1 Or TextBox12.Text.Trim().Length() > 2 Or TextBox11.Text.Trim().Length() <> 4 Then
                    MsgBox("Input tanggal tidak valid")
                Else
                    Call Enter_press_on_TextBox()
                End If
            End If
        End If
    End Sub
End Class
