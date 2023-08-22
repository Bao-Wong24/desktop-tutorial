Imports System.Data.SqlClient
Public Class Purchase_Main_Form

    Dim command3 As New SqlCommand
    Dim connection3 As New SqlConnection("Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True")

    Private Sub lblordph_Click(sender As Object, e As EventArgs) Handles lblpuruserph.Click

    End Sub

    Private Sub Purchase_Main_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.PRODUCT' table. You can move, or remove it, as needed.
        Me.PRODUCTTableAdapter.Fill(Me.CANTONIMS_DBDataSet.PRODUCT)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.SUPPLIER' table. You can move, or remove it, as needed.
        Me.SUPPLIERTableAdapter.Fill(Me.CANTONIMS_DBDataSet.SUPPLIER)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.ORDERS' table. You can move, or remove it, as needed.
        Me.ORDERSTableAdapter.Fill(Me.CANTONIMS_DBDataSet.ORDERS)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.USERS' table. You can move, or remove it, as needed.
        Me.USERSTableAdapter.Fill(Me.CANTONIMS_DBDataSet.USERS)

        display_data()
        connection3.Close()


    End Sub

    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        WelcomeForm.Show()
        Hide()
    End Sub

    Private Sub ADDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDToolStripMenuItem.Click

        Dim addSql As String = "INSERT INTO PURCHASE VALUES (@A, @B, @C, @D, @E, @F, @G, @H, @I, @J, @K)"
        Dim connStr As String = "Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True"
        Using connection3 As New SqlConnection(connStr),
                command3 As New SqlCommand(addSql, connection3)

            command3.Parameters.AddWithValue("@A", txtpurchno.Text)
            command3.Parameters.AddWithValue("@B", cbpurchuser.SelectedValue.ToString())
            command3.Parameters.AddWithValue("@C", cbpurchPhone.SelectedValue.ToString())
            command3.Parameters.AddWithValue("@D", cbpurchordNum.SelectedValue.ToString())
            command3.Parameters.AddWithValue("@E", cbpurchsuppnum.SelectedValue.ToString())
            command3.Parameters.AddWithValue("@F", cbpurprodnum.SelectedValue.ToString())
            command3.Parameters.AddWithValue("@G", cbpurchUOM.SelectedValue.ToString())
            command3.Parameters.AddWithValue("@H", cbpurchqty.SelectedValue.ToString())
            command3.Parameters.AddWithValue("@I", cbpurchordcost.SelectedValue.ToString())
            command3.Parameters.AddWithValue("@J", cbpurchTotalCost.SelectedValue.ToString())
            command3.Parameters.AddWithValue("@K", DateTimePicker3.Value.ToString())



            connection3.Open()
            command3.ExecuteNonQuery()


        End Using

        MessageBox.Show("Record Inserted Successful")

        display_data()
        connection3.Close()

    End Sub

    Private Sub display_data()

        If connection3.State = ConnectionState.Open Then
            connection3.Close()

        End If
        connection3.Open()


        command3 = connection3.CreateCommand()
        command3.CommandType = CommandType.Text
        command3.CommandText = "select PurchNo, USERS.UserTINno, PurchPhone, ORDERS.OrdNo, SUPPLIER.SuppNo, PRODUCT.ProdNo, PRODUCT.ProdUOM, PurchQty, PurchCost, PurchTotalCost, PurchDate FROM PURCHASE, ORDERS, PRODUCT, SUPPLIER, USERS"
        command3.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(command3)
        da.Fill(dt)

        DataGridView6.DataSource = dt
        connection3.Close()
        DataGridView6.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLOSEToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub CATEGORYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CATEGORYToolStripMenuItem.Click
        Category_Main_Form.Show()
        Hide()
    End Sub

    Private Sub SUPPLIERToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SUPPLIERToolStripMenuItem.Click
        Supply_Main_Form.Show()
        Hide()
    End Sub

    Private Sub USERSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles USERSToolStripMenuItem.Click
        Users_Main_Form.Show()
        Hide()
    End Sub

    Private Sub PRODUCTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PRODUCTToolStripMenuItem.Click
        ProductForm.Show()
        Hide()
    End Sub

    Private Sub ORDERSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ORDERSToolStripMenuItem.Click
        OrderForm.Show()
        Hide()
    End Sub

    Private Sub CLEARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLEARToolStripMenuItem.Click
        txtpurchno.Text = String.Empty
        cbpurchordcost.SelectedIndex = -1
        cbpurchordNum.SelectedIndex = -1
        cbpurchPhone.SelectedIndex = -1
        cbpurchqty.SelectedIndex = -1
        cbpurchsuppnum.SelectedIndex = -1
        cbpurchTotalCost.SelectedIndex = -1
        cbpurchUOM.SelectedIndex = -1
        cbpurchuser.SelectedIndex = -1
        cbpurprodnum.SelectedIndex = -1
    End Sub

    Private Sub DELETEToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles DELETEToolStripMenuItem2.Click

        If connection3.State = ConnectionState.Open Then
            connection3.Close()

        End If
        connection3.Open()

        'for sql database/set only
        command3 = connection3.CreateCommand()
        command3.CommandType = CommandType.Text
        command3.CommandText = "delete from PURCHASE where PurchNo='" & txtpurchno.Text & "'"

        command3.ExecuteNonQuery()

        MessageBox.Show("Record Deleted In Database Successful")


        'for data grid only 
        If DataGridView6.SelectedRows.Count >= 1 Then
            For j As Integer = DataGridView6.SelectedRows.Count - 1 To 0 Step -1
                DataGridView6.Rows.RemoveAt(DataGridView6.SelectedRows(j).Index)
            Next


            MessageBox.Show("Record Deleted in Data Table Successful")

        Else
            MessageBox.Show("No rows to select")
        End If


    End Sub

    Private Sub SEARCHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SEARCHToolStripMenuItem.Click

        If connection3.State = ConnectionState.Open Then
            connection3.Close()

        End If
        connection3.Open()

        command3 = connection3.CreateCommand()
        command3.CommandType = CommandType.Text
        command3.CommandText = "select * from PURCHASE where PurchNo='" & txtpurchno.Text & "' OR PurchPhone='" & cbpurchPhone.SelectedValue.ToString() & "'"
        command3.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(command3)
        da.Fill(dt)

        DataGridView6.DataSource = dt
        connection3.Close()
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UPDATEToolStripMenuItem.Click

        If connection3.State = ConnectionState.Open Then
            connection3.Close()

        End If
        connection3.Open()

        command3 = connection3.CreateCommand()
        command3.CommandType = CommandType.Text
        command3.CommandText = "update PURCHASE set UserTINno= @B, PurchPhone= @C, OrdNo= @D, SuppNo= @E, ProdNo= @F, ProdUOM= @G, PurchQty= @H, PurchCost= @I, PurchTotalCost= @J, PurchDate= @K WHERE PurchNo= @A"


        command3.Parameters.AddWithValue("@A", txtpurchno.Text)
        command3.Parameters.AddWithValue("@B", cbpurchuser.SelectedValue.ToString())
        command3.Parameters.AddWithValue("@C", cbpurchPhone.SelectedValue.ToString())
        command3.Parameters.AddWithValue("@D", cbpurchordNum.SelectedValue.ToString())
        command3.Parameters.AddWithValue("@E", cbpurchsuppnum.SelectedValue.ToString())
        command3.Parameters.AddWithValue("@F", cbpurprodnum.SelectedValue.ToString())
        command3.Parameters.AddWithValue("@G", cbpurchUOM.SelectedValue.ToString())
        command3.Parameters.AddWithValue("@H", cbpurchqty.SelectedValue.ToString())
        command3.Parameters.AddWithValue("@I", cbpurchordcost.SelectedValue.ToString())
        command3.Parameters.AddWithValue("@J", cbpurchTotalCost.SelectedValue.ToString())
        command3.Parameters.AddWithValue("@K", DateTimePicker3.Value.ToString())


        command3.ExecuteNonQuery()


        MessageBox.Show("Update SuccessFul")

        display_data()
        connection3.Close()
    End Sub
End Class