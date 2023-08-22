Imports System.Data.SqlClient
Public Class Category_Main_Form

    Dim con4 As New SqlConnection
    Dim cmd4 As New SqlCommand
    Dim k As Integer
    'Dim n As Integer
    'Dim strcat As String

    Dim intcatnum(3) As Integer
    Dim strcatnum As String
    Dim intnum As Integer
    Dim intcatcounter As Integer = 1
    Dim intreader As Integer
    Private Sub Category_Main_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet1.CATEGORY' table. You can move, or remove it, as needed.
        Me.CATEGORYTableAdapter1.Fill(Me.CANTONIMS_DBDataSet1.CATEGORY)

        con4.ConnectionString = "Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True"
        If con4.State = ConnectionState.Open Then
            con4.Close()

        End If
        con4.Open()

        Set_Categoryvalue()
        Set_category()
        display_data()

    End Sub

    Private Sub Set_Categoryvalue()


        For intnum = 0 To 3
            intcatnum(intnum) = intcatcounter
            intcatcounter += 1
        Next


        For Each intreader In intcatnum
            strcatnum = intreader.ToString()
            cbcategorynum.Items.Add(strcatnum)
        Next

    End Sub

    Private Sub Set_category()

        Dim strcategory() As String = {"Food", "Drink", "Dairy/Beverage", "Other"}
        Dim strReader As String

        For Each strReader In strcategory
            cbcategory.Items.Add(strReader)
        Next

    End Sub


    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        WelcomeForm.Show()
        Hide()
    End Sub

    Private Sub ADDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ADDToolStripMenuItem.Click

        If con4.State = ConnectionState.Open Then
            con4.Close()

        End If
        con4.Open()

        cmd4 = con4.CreateCommand()
        cmd4.CommandType = CommandType.Text
        cmd4.CommandText = "insert into CATEGORY (CatNo, CatName) 
           values('" + cbcategorynum.SelectedItem + "','" + cbcategory.SelectedItem + "')"
        cmd4.ExecuteNonQuery()


        MessageBox.Show("Record Inserted Successful")


        display_data()
        con4.Close()


    End Sub

    Public Sub display_data()

        cmd4 = con4.CreateCommand()
        cmd4.CommandType = CommandType.Text
        cmd4.CommandText = "select * from CATEGORY"
        cmd4.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(cmd4)
        da.Fill(dt)

        DataGridView3.DataSource = dt
        DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UPDATEToolStripMenuItem.Click

        k = Convert.ToInt32(DataGridView3.SelectedCells.Item(0).Value.ToString())
        cbcategorynum.SelectedItem = k.ToString()

        If con4.State = ConnectionState.Open Then
            con4.Close()

        End If
        con4.Open()


        'write update query
        cmd4 = con4.CreateCommand()
        cmd4.CommandType = CommandType.Text
        cmd4.CommandText = "update CATEGORY set CatNo= @A, CatName= @B WHERE CatNo= @A"

        cmd4.Parameters.AddWithValue("@A", cbcategorynum.SelectedItem)
        cmd4.Parameters.AddWithValue("@B", cbcategory.SelectedItem)

        cmd4.ExecuteNonQuery()

        MessageBox.Show("Update SuccessFul")

        display_data()
        con4.Close()

    End Sub

    Private Sub SEARCHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SEARCHToolStripMenuItem.Click

        cmd4 = con4.CreateCommand()
        cmd4.CommandType = CommandType.Text
        cmd4.CommandText = "select * from CATEGORY where CatNo='" + cbcategorynum.SelectedItem + "' OR  CatName='" + cbcategory.SelectedItem + "'"
        con4.Open()
        cmd4.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(cmd4)
        da.Fill(dt)

        DataGridView3.DataSource = dt
        con4.Close()
    End Sub

    Private Sub CLEARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLEARToolStripMenuItem.Click

        cbcategorynum.SelectedIndex = -1
        cbcategory.SelectedIndex = -1
    End Sub

    Private Sub CLOSEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CLOSEToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub PRODUCTToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PRODUCTToolStripMenuItem.Click
        ProductForm.Show()
        Hide()
    End Sub

    Private Sub DELETEToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles DELETEToolStripMenuItem2.Click
        If con4.State = ConnectionState.Open Then
            con4.Close()

        End If
        con4.Open()

        'for sql database/set only
        cmd4 = con4.CreateCommand()
        cmd4.CommandType = CommandType.Text
        cmd4.CommandText = "delete from CATEGORY where CatNo='" + cbcategorynum.SelectedItem + "' OR CatName='" + cbcategory.SelectedItem + "'"
        cmd4.ExecuteNonQuery()


        MessageBox.Show("Record Deleted In Database Successful")
        'for data grid only 
        If DataGridView3.SelectedRows.Count >= 1 Then
            For i As Integer = DataGridView3.SelectedRows.Count - 1 To 0 Step -1
                DataGridView3.Rows.RemoveAt(DataGridView3.SelectedRows(i).Index)
            Next


            MessageBox.Show("Record Deleted in Data Table Successful")

        Else
            MessageBox.Show("No rows to select")
        End If
    End Sub

    Private Sub NewUsersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewUsersToolStripMenuItem.Click
        Users_Main_Form.Show()
        Hide()
    End Sub

    Private Sub NewOrderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewOrderToolStripMenuItem.Click
        OrderForm.Show()
        Hide()
    End Sub

    Private Sub NewSupplierToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewSupplierToolStripMenuItem.Click
        Supply_Main_Form.Show()
        Hide()
    End Sub

    Private Sub NewPurchaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewPurchaseToolStripMenuItem.Click
        Purchase_Main_Form.ShowDialog()
        Hide()
    End Sub
End Class