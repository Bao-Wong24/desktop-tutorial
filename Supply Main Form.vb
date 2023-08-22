Imports System.Data.SqlClient
Imports System.Security.Cryptography

Public Class Supply_Main_Form


    Dim con3 As New SqlConnection
    Dim cmd3 As New SqlCommand
    Dim n As Integer
    Private Sub Supply_Main_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'VbimsdbDataSet.Supplier' table. You can move, or remove it, as needed.
        Me.SupplierTableAdapter.Fill(Me.VbimsdbDataSet.Supplier)

        con3.ConnectionString = "Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True"
        If con3.State = ConnectionState.Open Then
            con3.Close()

        End If
        con3.Open()

        display_data1()
        con3.Close()
    End Sub

    Private Sub ADDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDToolStripMenuItem.Click

        If con3.State = ConnectionState.Open Then
            con3.Close()

        End If
        con3.Open()

        cmd3 = con3.CreateCommand()
        cmd3.CommandType = CommandType.Text
        cmd3.CommandText = "insert into SUPPLIER (SuppNo, SuppName, SuppPhone, SuppEmail, SuppAddress) 
        values('" + txtsuppno.Text + "','" + txtsuppname.Text + "','" + txtsuppPh.Text + "','" + txtsuppEmail.Text + "', '" + txtaddr.Text + "')"
        cmd3.ExecuteNonQuery()

        MessageBox.Show("Record Inserted Successful")


        display_data1()
        con3.Close()
    End Sub

    Public Sub display_data1()

        cmd3 = con3.CreateCommand()
        cmd3.CommandType = CommandType.Text
        cmd3.CommandText = "select * from SUPPLIER"
        cmd3.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(cmd3)
        da.Fill(dt)

        DataGridView2.DataSource = dt
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        WelcomeForm.Show()
        Hide()
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UPDATEToolStripMenuItem.Click

        'n = Convert.ToInt32(DataGridView2.SelectedCells.Item(0).Value.ToString())
        'txtsuppno.Text = n.ToString()

        'keep connection
        If con3.State = ConnectionState.Open Then
            con3.Close()

        End If
        con3.Open()

        'write update query
        cmd3 = con3.CreateCommand()
        cmd3.CommandType = CommandType.Text
        cmd3.CommandText = "update SUPPLIER set SuppName= @B, SuppPhone= @C, SuppEmail= @D, SuppAddress= @E WHERE SuppNo= @A"

        cmd3.Parameters.AddWithValue("@A", txtsuppno.Text)
        cmd3.Parameters.AddWithValue("@B", txtsuppname.Text)
        cmd3.Parameters.AddWithValue("@C", txtsuppPh.Text)
        cmd3.Parameters.AddWithValue("@D", txtsuppEmail.Text)
        cmd3.Parameters.AddWithValue("@E", txtaddr.Text)

        cmd3.ExecuteNonQuery()


        MessageBox.Show("Update SuccessFul")

        display_data1()
        con3.Close()
    End Sub

    Private Sub SEARCHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SEARCHToolStripMenuItem.Click

        cmd3 = con3.CreateCommand()
        cmd3.CommandType = CommandType.Text
        cmd3.CommandText = "select * from SUPPLIER where SuppNo='" + txtsuppno.Text + "' OR SuppName='" + txtsuppname.Text + "' OR SuppPhone='" + txtsuppPh.Text + "'
        OR SuppEmail='" + txtsuppEmail.Text + "' OR SuppAddress='" + txtaddr.Text + "'"
        con3.Open()
        cmd3.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(cmd3)
        da.Fill(dt)

        DataGridView2.DataSource = dt
        con3.Close()
    End Sub

    Private Sub CLEARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLEARToolStripMenuItem.Click
        txtaddr.Text = String.Empty
        txtsuppEmail.Text = String.Empty
        txtsuppname.Text = String.Empty
        txtsuppno.Text = String.Empty
        txtsuppPh.Text = String.Empty
    End Sub

    Private Sub CategoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CategoryToolStripMenuItem.Click
        Category_Main_Form.Show()
        Hide()
    End Sub

    Private Sub NewUsersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewUsersToolStripMenuItem.Click
        Users_Main_Form.Show()
        Hide()
    End Sub

    Private Sub NewSupplierToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewSupplierToolStripMenuItem.Click
        Purchase_Main_Form.Show()
        Hide()
    End Sub

    Private Sub NewPurchaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewPurchaseToolStripMenuItem.Click
        ProductForm.Show()
        Hide()
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLOSEToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub DELETEToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles DELETEToolStripMenuItem2.Click
        If con3.State = ConnectionState.Open Then
            con3.Close()

        End If
        con3.Open()

        'for sql database/set only
        cmd3 = con3.CreateCommand()
        cmd3.CommandType = CommandType.Text
        cmd3.CommandText = "delete from SUPPLIER where SuppNo='" + txtsuppno.Text + "' OR SuppName='" + txtsuppname.Text + "' OR SuppPhone='" + txtsuppPh.Text + "'
        AND SuppEmail='" + txtsuppEmail.Text + "' OR SuppAddress='" + txtaddr.Text + "'"
        cmd3.ExecuteNonQuery()


        MessageBox.Show("Record Deleted In Database Successful")

        'for data grid only 
        If DataGridView2.SelectedRows.Count >= 1 Then
            For j As Integer = DataGridView2.SelectedRows.Count - 1 To 0 Step -1
                DataGridView2.Rows.RemoveAt(DataGridView2.SelectedRows(j).Index)
            Next


            MessageBox.Show("Record Deleted in Data Table Successful")

        Else
            MessageBox.Show("No rows to select")
        End If

    End Sub

    Private Sub NewOrderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewOrderToolStripMenuItem.Click
        OrderForm.Show()
        Hide()
    End Sub
End Class