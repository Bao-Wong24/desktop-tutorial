Imports System.Collections.ObjectModel
Imports System.Data.Common
Imports System.Data.SqlClient

Public Class OrderForm

    Dim command2 As New SqlCommand
    Dim connection2 As New SqlConnection("Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True")


    Private Sub OrderForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.USERS' table. You can move, or remove it, as needed.
        Me.USERSTableAdapter.Fill(Me.CANTONIMS_DBDataSet.USERS)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.CATEGORY' table. You can move, or remove it, as needed.
        Me.CATEGORYTableAdapter.Fill(Me.CANTONIMS_DBDataSet.CATEGORY)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.PRODUCT' table. You can move, or remove it, as needed.
        Me.PRODUCTTableAdapter.Fill(Me.CANTONIMS_DBDataSet.PRODUCT)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.SUPPLIER' table. You can move, or remove it, as needed.
        Me.SUPPLIERTableAdapter.Fill(Me.CANTONIMS_DBDataSet.SUPPLIER)


        display_data()
        connection2.Close()


    End Sub

    Private Sub ADDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDToolStripMenuItem.Click


        Dim addSql As String = "INSERT INTO ORDERS VALUES (@A, @B, @C, @D, @E, @F, @G, @H, @I, @J, @K, @L)"
        Dim connStr As String = "Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True"
        Using connection2 As New SqlConnection(connStr),
                command2 As New SqlCommand(addSql, connection2)

            command2.Parameters.AddWithValue("@A", TbOrdNum.Text)
            command2.Parameters.AddWithValue("@B", cbordsuppnum.SelectedValue.ToString())
            command2.Parameters.AddWithValue("@C", cbordprodnum.SelectedValue.ToString())
            command2.Parameters.AddWithValue("@D", cbordcatnum.SelectedValue.ToString())
            command2.Parameters.AddWithValue("@E", TbOrdQty.Text)
            command2.Parameters.AddWithValue("@F", TbVat.Text)
            command2.Parameters.AddWithValue("@G", TbST.Text)
            command2.Parameters.AddWithValue("@H", TbTP.Text)
            command2.Parameters.AddWithValue("@I", cbordPhone.Text)
            command2.Parameters.AddWithValue("@J", cborduseTIN.SelectedValue.ToString())
            command2.Parameters.AddWithValue("@K", DateTimePicker2.Value.ToString())
            command2.Parameters.AddWithValue("@L", cbordprice.Text)


            connection2.Open()
            command2.ExecuteNonQuery()


        End Using

        MessageBox.Show("Record Inserted Successful")

        display_data()
        connection2.Close()

    End Sub

    Private Sub display_data()

        If connection2.State = ConnectionState.Open Then
            connection2.Close()

        End If
        connection2.Open()


        command2 = connection2.CreateCommand()
        command2.CommandType = CommandType.Text
        command2.CommandText = "select OrdNo, SUPPLIER.SuppNo, PRODUCT.ProdNO, CATEGORY.CatNo, OrdQty, OrdPrice, OrdVAT, OrdSubTotal, OrdTotalPrice,
 SUPPLIER.SuppPhone, USERS.UserTINno, OrdDate FROM ORDERS, PRODUCT, SUPPLIER, CATEGORY, USERS"
        command2.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(command2)
        da.Fill(dt)

        DataGridView5.DataSource = dt
        connection2.Close()
        DataGridView5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UPDATEToolStripMenuItem.Click

        If connection2.State = ConnectionState.Open Then
            connection2.Close()

        End If
        connection2.Open()

        command2 = connection2.CreateCommand()
        command2.CommandType = CommandType.Text
        command2.CommandText = "update ORDERS set SuppNo= @B, ProdNo= @C, CatNo= @D, OrdQty= @E, OrdVAT= @F, OrdSubTotal= @G, OrdTotalPrice= @H, SuppPhone= @I, UserTINno= @J, OrdDate= @K, OrdPrice= @L WHERE OrdNo= @A"

        command2.Parameters.AddWithValue("@A", TbOrdNum.Text)
        command2.Parameters.AddWithValue("@B", cbordsuppnum.SelectedValue.ToString())
        command2.Parameters.AddWithValue("@C", cbordprodnum.SelectedValue.ToString())
        command2.Parameters.AddWithValue("@D", cbordcatnum.SelectedValue.ToString())
        command2.Parameters.AddWithValue("@E", TbOrdQty.Text)
        command2.Parameters.AddWithValue("@F", TbVat.Text)
        command2.Parameters.AddWithValue("@G", TbST.Text)
        command2.Parameters.AddWithValue("@H", TbTP.Text)
        command2.Parameters.AddWithValue("@I", cbordPhone.SelectedValue.ToString())
        command2.Parameters.AddWithValue("@J", cborduseTIN.SelectedValue.ToString())
        command2.Parameters.AddWithValue("@K", DateTimePicker2.Value.ToString())
        command2.Parameters.AddWithValue("@L", cbordprice.SelectedValue.ToString())

        command2.ExecuteNonQuery()


        MessageBox.Show("Update SuccessFul")

        display_data()
        connection2.Close()
    End Sub

    Private Sub SEARCHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SEARCHToolStripMenuItem.Click


        If connection2.State = ConnectionState.Open Then
            connection2.Close()

        End If
        connection2.Open()

        command2 = connection2.CreateCommand()
        command2.CommandType = CommandType.Text
        command2.CommandText = "select * from ORDERS where OrdNo='" + TbOrdNum.Text + "' OR SuppPhone='" + cbordPhone.SelectedValue.ToString() + "' OR OrdPrice='" + cbordprice.Text + "'"
        command2.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(command2)
        da.Fill(dt)

        DataGridView5.DataSource = dt
        connection2.Close()
    End Sub

    Private Sub DELETEToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles DELETEToolStripMenuItem2.Click
        If connection2.State = ConnectionState.Open Then
            connection2.Close()

        End If
        connection2.Open()

        'for sql database/set only
        command2 = connection2.CreateCommand()
        command2.CommandType = CommandType.Text
        command2.CommandText = "delete from ORDERS where OrdNo='" & TbOrdNum.Text & "' OR SuppNo='" & cbordsuppnum.SelectedValue.ToString() & "' OR ProdNo='" & cbordprodnum.SelectedValue.ToString() & "' OR CatNo='" & cbordcatnum.SelectedValue.ToString() & "' 
        OR OrdQty='" & TbOrdQty.Text & "' OR OrdPrice='" & cbordprice.SelectedValue.ToString() & "' OR OrdVAT='" & TbVat.Text & "' OR OrdSubTotal='" & TbST.Text & "' OR OrdTotalPrice='" & TbTP.Text & "' OR SuppPhone='" + cbordPhone.SelectedValue.ToString() + "' OR
UserTINno='" + cborduseTIN.SelectedValue.ToString() + "' OR OrdDate='" + DateTimePicker2.Value + "'"

        command2.ExecuteNonQuery()

        MessageBox.Show("Record Deleted In Database Successful")

        'for data grid only 
        If DataGridView5.SelectedRows.Count >= 1 Then
            For j As Integer = DataGridView5.SelectedRows.Count - 1 To 0 Step -1
                DataGridView5.Rows.RemoveAt(DataGridView5.SelectedRows(j).Index)
            Next


            MessageBox.Show("Record Deleted in Data Table Successful")

        Else
            MessageBox.Show("No rows to select")
        End If

    End Sub

    Private Sub CLEARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLEARToolStripMenuItem.Click
        TbOrdNum.Text = String.Empty
        TbOrdQty.Text = String.Empty
        TbST.Text = String.Empty
        TbTP.Text = String.Empty
        TbVat.Text = String.Empty
        cbordcatnum.SelectedIndex = -1
        cbordPhone.SelectedIndex = -1
        cbordprice.SelectedIndex = -1
        cbordprodnum.SelectedIndex = -1
        cbordsuppnum.SelectedIndex = -1
        cborduseTIN.SelectedIndex = -1

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

    Private Sub PURCHASEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PURCHASEToolStripMenuItem.Click
        Purchase_Main_Form.Show()
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

    Private Sub CALCULATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CALCULATEToolStripMenuItem.Click

        'Subtotal without tax rate
        Dim dblprice As Double
        Dim intqty As Integer
        Dim dblrate As Double
        Dim dblSubTotal As Double
        Dim dblTotalPrice As Double


        dblprice = Convert.ToDouble(cbordprice.SelectedValue)
        intqty = Convert.ToInt32(TbOrdQty.Text)
        dblrate = Convert.ToDouble(TbVat.Text)


        dblSubTotal = dblprice * intqty
        TbST.Text = dblSubTotal.ToString("C2")

        dblTotalPrice = dblSubTotal * dblrate + dblSubTotal
        TbTP.Text = dblTotalPrice.ToString("C2")

        'Subot
    End Sub

    Private Sub btnProdHome_Click(sender As Object, e As EventArgs) Handles btnProdHome.Click
        WelcomeForm.Show()
        Hide()
    End Sub
End Class