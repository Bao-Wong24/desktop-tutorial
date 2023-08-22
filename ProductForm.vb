Imports System.Data.SqlClient
Imports System.Windows.Input

Public Class ProductForm

    Dim command As New SqlCommand
    Dim connection As New SqlConnection("Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True")


    Private Sub ProductForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet2.SUPPLIER' table. You can move, or remove it, as needed.
        Me.SUPPLIERTableAdapter1.Fill(Me.CANTONIMS_DBDataSet2.SUPPLIER)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet2.CATEGORY' table. You can move, or remove it, as needed.
        Me.CATEGORYTableAdapter2.Fill(Me.CANTONIMS_DBDataSet2.CATEGORY)

        display_data()
        connection.Close()
    End Sub
    Private Sub ADDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDToolStripMenuItem.Click

        Dim addSql As String = "INSERT INTO PRODUCT VALUES (@A, @B, @C, @D, @E, @F, @G, @H, @I)"
        Dim connStr As String = "Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True"
        Using connection As New SqlConnection(connStr),
                command As New SqlCommand(addSql, connection)
            'command = connection.CreateCommand()
            'command.CommandType = CommandType.Text
            'command.CommandText = "insert into PRODUCT(ProdNo, ProdName, ProdQty, ProdUnitPrice, ProdUOM, ProdDate, CatNo, SuppNo, ProdTotalPrice) values('" & TbProdNum.Text & "','" & Tbprodname.Text & "','" & Tbquantity.Text & "','" & TbUP.Text & "','" & TbUOM.Text & "','" & DateTimePicker1.Value & "','" & CbcategoryNum.SelectedItem & "','" & CbSuppNum.SelectedItem & "','" & TbTP.Text & "')"
            'command.CommandText = "INSERT INTO PRODUCT (ProdNo, ProdName, ProdQty, ProdUnitPrice, ProdUOM, ProdDate, CatNo, SuppNo, ProdTotalPrice) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"

            command.Parameters.AddWithValue("@A", TbProdNum.Text)
            command.Parameters.AddWithValue("@B", Tbprodname.Text)
            command.Parameters.AddWithValue("@C", Tbquantity.Text)
            command.Parameters.AddWithValue("@D", TbUP.Text)
            command.Parameters.AddWithValue("@E", TbUOM.Text)
            command.Parameters.AddWithValue("@F", DateTimePicker1.Value.ToString())
            command.Parameters.AddWithValue("@G", CbcategoryNum.SelectedValue.ToString())
            command.Parameters.AddWithValue("@H", CbSuppNum.SelectedValue.ToString())
            command.Parameters.AddWithValue("@I", TbTP.Text)

            connection.Open()
            command.ExecuteNonQuery()

        End Using

        MessageBox.Show("Record Inserted Successful")

        display_data()
        connection.Close()

    End Sub

    Private Sub display_data()

        command = connection.CreateCommand()
        command.CommandType = CommandType.Text
        command.CommandText = "select ProdNo, ProdName, ProdQty, ProdUnitPrice, ProdUOM, ProdDate, ProdTotalPrice, CATEGORY.CatNo, SUPPLIER.SuppNo FROM PRODUCT, SUPPLIER, CATEGORY"
        connection.Open()
        command.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(command)
        da.Fill(dt)

        DataGridView4.DataSource = dt
        DataGridView4.Refresh()
        connection.Close()
        DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


    End Sub

    Private Sub CALCULATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CALCULATEToolStripMenuItem.Click

        Dim dblprice As Double
        Dim intqty As Integer
        Dim dbltotalcost As Double

        dblprice = Convert.ToDouble(TbUP.Text)
        intqty = Convert.ToInt32(Tbquantity.Text)

        dbltotalcost = dblprice * intqty
        TbTP.Text = dbltotalcost.ToString("C2")

    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UPDATEToolStripMenuItem.Click

        If connection.State = ConnectionState.Open Then
            connection.Close()

        End If
        connection.Open()

        command = connection.CreateCommand()
        command.CommandType = CommandType.Text
        command.CommandText = "update PRODUCT set ProdName= @B, ProdQty= @C, ProdUnitPrice= @D, ProdUOM= @E, ProdDate= @F, CatNo= @G, SuppNo= @H, ProdTotalPrice= @I WHERE ProdNo= @A"

        command.Parameters.AddWithValue("@A", TbProdNum.Text)
        command.Parameters.AddWithValue("@B", Tbprodname.Text)
        command.Parameters.AddWithValue("@C", Tbquantity.Text)
        command.Parameters.AddWithValue("@D", TbUP.Text)
        command.Parameters.AddWithValue("@E", TbUOM.Text)
        command.Parameters.AddWithValue("@F", DateTimePicker1.Value.ToString())
        command.Parameters.AddWithValue("@G", CbcategoryNum.SelectedItem)
        command.Parameters.AddWithValue("@H", CbSuppNum.SelectedItem)
        command.Parameters.AddWithValue("@I", TbTP.Text)

        command.ExecuteNonQuery()


        MessageBox.Show("Update SuccessFul")

        display_data()
        connection.Close()

    End Sub

    Private Sub SEARCHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SEARCHToolStripMenuItem.Click


        If connection.State = ConnectionState.Open Then
            connection.Close()

        End If
        connection.Open()

        command = connection.CreateCommand()
        command.CommandType = CommandType.Text
        command.CommandText = "select * from PRODUCT where ProdNo='" + TbProdNum.Text + "' OR ProdName='" + Tbprodname.Text + "'"
        command.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(command)
        da.Fill(dt)

        DataGridView4.DataSource = dt
        connection.Close()

    End Sub

    Private Sub btnProdHome_Click(sender As Object, e As EventArgs) Handles btnProdHome.Click
        WelcomeForm.Show()
        Hide()
    End Sub

    Private Sub DELETEToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles DELETEToolStripMenuItem2.Click

        If connection.State = ConnectionState.Open Then
            connection.Close()

        End If
        connection.Open()

        'for sql database/set only
        command = connection.CreateCommand()
        command.CommandType = CommandType.Text
        command.CommandText = "delete from PRODUCT where ProdNo='" & TbProdNum.Text & "' OR ProdName='" & Tbprodname.Text & "' OR ProdQty='" & Tbquantity.Text & "' OR ProdUnitPrice='" & TbUP.Text & "' 
        OR ProdUOM='" & TbUOM.Text & "' OR ProdDate='" & DateTimePicker1.Value & "' OR CatNo='" & 1 & "' OR SuppNo='" & 1 & "' OR ProdTotalPrice='" & TbTP.Text & "'"

        command.ExecuteNonQuery()

        MessageBox.Show("Record Deleted In Database Successful")

        'for data grid only 
        If DataGridView4.SelectedRows.Count >= 1 Then
            For j As Integer = DataGridView4.SelectedRows.Count - 1 To 0 Step -1
                DataGridView4.Rows.RemoveAt(DataGridView4.SelectedRows(j).Index)
            Next


            MessageBox.Show("Record Deleted in Data Table Successful")

        Else
            MessageBox.Show("No rows to select")
        End If


    End Sub

    Private Sub CLEARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLEARToolStripMenuItem.Click
        TbProdNum.Text = String.Empty
        Tbprodname.Text = String.Empty
        Tbquantity.Text = String.Empty
        TbTP.Text = String.Empty
        TbUOM.Text = String.Empty
        TbUP.Text = String.Empty
        CbcategoryNum.SelectedIndex = -1
        CbSuppNum.SelectedIndex = -1
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

    Private Sub ORDERToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ORDERToolStripMenuItem.Click
        OrderForm.Show()
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
End Class