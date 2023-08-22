Imports System.Data.SqlClient

Public Class Users_Main_Form

    'declare connection
    'Dim con1 As New SqlConnection
    'Dim cmd1 As New SqlCommand

    Dim con2 As New SqlConnection
    Dim cmd2 As New SqlCommand
    Dim n As Integer

    Private Sub Users_Main_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'TODO: This line of code loads data into the 'VbimsdbDataSet.Users' table. You can move, or remove it, as needed.
        Me.UsersTableAdapter.Fill(Me.VbimsdbDataSet.Users)


        'check database connection 

        'con1.ConnectionString = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\wongb.LAPTOP-QSA8B0CQ\Desktop\216 DevSoftVB Development\Inventory System\Solution Software\CANTON IMS\CANTON IMS\vbimsdb.mdf;Integrated Security=True"
        'If con1.State = ConnectionState.Open Then
        'con1.Close()

        'End If
        'con1.Open()

        con2.ConnectionString = "Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True"

        If con2.State = ConnectionState.Open Then
            con2.Close()

        End If
        con2.Open()

        display_data()
        con2.Close()

    End Sub

    Public Sub display_data()

        cmd2 = con2.CreateCommand()
        cmd2.CommandType = CommandType.Text
        cmd2.CommandText = "select * from USERS"
        cmd2.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(cmd2)
        da.Fill(dt)

        DataGridView1.DataSource = dt
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        WelcomeForm.Show()
        Hide()
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UPDATEToolStripMenuItem.Click


        'n = Convert.ToInt32(DataGridView1.SelectedCells.Item(0).Value.ToString())
        'txtTin.Text = n.ToString()

        'keep connection
        If con2.State = ConnectionState.Open Then
            con2.Close()

        End If
        con2.Open()

        'write update query
        cmd2 = con2.CreateCommand()
        cmd2.CommandType = CommandType.Text
        cmd2.CommandText = "update USERS set FirstName= @B, LastName= @C, UserName= @D, Password= @E, Email= @F, Role= @G,  city= @H, Phone = @I WHERE  UserTINno= @A"

        cmd2.Parameters.AddWithValue("@A", txtTin.Text)
        cmd2.Parameters.AddWithValue("@B", txtFN.Text)
        cmd2.Parameters.AddWithValue("@C", txtLN.Text)
        cmd2.Parameters.AddWithValue("@D", txtusername.Text)
        cmd2.Parameters.AddWithValue("@E", txtpass.Text)
        cmd2.Parameters.AddWithValue("@F", txtEmail.Text)
        cmd2.Parameters.AddWithValue("@G", txtRole.Text)
        cmd2.Parameters.AddWithValue("@H", txtcity.Text)
        cmd2.Parameters.AddWithValue("@I", txtphone.Text)


        cmd2.ExecuteNonQuery()

        display_data()
        con2.Close()

    End Sub

    Private Sub CLEARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLEARToolStripMenuItem.Click
        txtTin.Text = String.Empty
        txtcity.Text = String.Empty
        txtEmail.Text = String.Empty
        txtFN.Text = String.Empty
        txtLN.Text = String.Empty
        txtpass.Text = String.Empty
        txtphone.Text = String.Empty
        txtRole.Text = String.Empty
        txtusername.Text = String.Empty
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLOSEToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub CATEGORYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CATEGORYToolStripMenuItem.Click
        Category_Main_Form.Show()
        Hide()
    End Sub

    Private Sub PRODUCTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PRODUCTToolStripMenuItem.Click
        ProductForm.Show()
        Hide()
    End Sub

    Private Sub SUPPLIERToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SUPPLIERToolStripMenuItem.Click
        Supply_Main_Form.Show()
        Hide()
    End Sub

    Private Sub PURCHASEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PURCHASEToolStripMenuItem.Click
        Purchase_Main_Form.Show()
        Hide()
    End Sub

    Private Sub SEARCHToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles SEARCHToolStripMenuItem.Click

        cmd2 = con2.CreateCommand()
        cmd2.CommandType = CommandType.Text
        cmd2.CommandText = "select * from USERS where FirstName='" + txtFN.Text + "' OR LastName='" + txtLN.Text + "' OR City='" + txtcity.Text + "' OR Phone='" + txtphone.Text + "'
        AND Role='" + txtRole.Text + "' OR Email='" + txtEmail.Text + "' OR Username='" + txtusername.Text + "' OR Password='" + txtpass.Text + "'"
        con2.Open()
        cmd2.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(cmd2)
        da.Fill(dt)

        DataGridView1.DataSource = dt
        con2.Close()

    End Sub

    Private Sub ADDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDToolStripMenuItem.Click

        cmd2 = con2.CreateCommand()
        cmd2.CommandType = CommandType.Text
        cmd2.CommandText = "insert into USERS (UserTINno, FirstName, LastName, UserName, Password, Email, Role, City, Phone) 
        values('" + txtTin.Text + "','" + txtFN.Text + "','" + txtLN.Text + "','" + txtusername.Text + "','" + txtpass.Text + "','" + txtEmail.Text + "','" + txtRole.Text + "',
        '" + txtcity.Text + "','" + txtphone.Text + "')"
        con2.Open()
        cmd2.ExecuteNonQuery()

        MessageBox.Show("Record Inserted Successful")

        display_data()
        con2.Close()
    End Sub

    Private Sub DELETEToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles DELETEToolStripMenuItem2.Click
        If con2.State = ConnectionState.Open Then
            con2.Close()

        End If
        con2.Open()

        'for sql database/set only
        cmd2 = con2.CreateCommand()
        cmd2.CommandType = CommandType.Text
        cmd2.CommandText = "delete from USERS where FirstName='" + txtFN.Text + "' OR LastName='" + txtLN.Text + "' OR City='" + txtcity.Text + "' OR Phone='" + txtphone.Text + "'"
        cmd2.ExecuteNonQuery()


        MessageBox.Show("Record Deleted In Database Successful")

        'for data grid only
        If DataGridView1.SelectedRows.Count >= 1 Then
            For i As Integer = DataGridView1.SelectedRows.Count - 1 To 0 Step -1
                DataGridView1.Rows.RemoveAt(DataGridView1.SelectedRows(i).Index)
            Next


            MessageBox.Show("Record Deleted in Data Table Successful")

        Else
            MessageBox.Show("No rows to select")
        End If
    End Sub
End Class