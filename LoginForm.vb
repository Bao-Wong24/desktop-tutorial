Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading

Public Class LoginForm
    Private Sub btnlogin_Click(sender As Object, e As EventArgs) Handles btnlogin.Click

        Dim con As SqlConnection = New SqlConnection("Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True")
        Dim cmd As SqlCommand = New SqlCommand("select * from USERS where UserName='" + txtuserfield.Text + "' and Password='" + txtpassfield.Text + "'", con)
        Dim da As SqlDataAdapter = New SqlDataAdapter(cmd)
        Dim dt As DataTable = New DataTable()
        da.Fill(dt)
        If (dt.Rows.Count > 0) Then
            MessageBox.Show("Login Success", "information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            WelcomeForm.Show()
            Hide()
        Else
            MessageBox.Show("Error", "information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            txtpassfield.UseSystemPasswordChar = True
        Else
            txtpassfield.UseSystemPasswordChar = False
        End If
    End Sub

    Private Sub btnlogoff_Click(sender As Object, e As EventArgs) Handles btnlogoff.Click
        Close()
    End Sub

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Thread.Sleep(4500)
    End Sub
End Class