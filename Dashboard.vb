Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices

Public Class Dashboard


    Dim command4 As New SqlCommand
    Dim connection4 As New SqlConnection("Data Source=LAPTOP-QSA8B0CQ\SQLEXPRESS;Initial Catalog=CANTONIMS_DB;Integrated Security=True")

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Chart1.Series("ORDERS").Points.AddXY("Orders", DataGridView1.Rows.Count)
        Me.Chart1.Series("PURCHASES").Points.AddXY("Purchases", DataGridView2.Rows.Count)
        Me.Chart1.Series("PRODUCTS").Points.AddXY("Products", DataGridView3.Rows.Count)

    End Sub

    Private Sub Button2_MouseEnter(sender As Object, e As EventArgs) Handles btnoverall.MouseEnter

        btnoverall.BackColor = Color.FromArgb(67, 83, 98)

    End Sub

    Private Sub Button2_MouseLeave(sender As Object, e As EventArgs) Handles btnoverall.MouseLeave

        btnoverall.BackColor = Color.FromArgb(67, 83, 98)

    End Sub


    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.PURCHASE' table. You can move, or remove it, as needed.
        Me.PURCHASETableAdapter.Fill(Me.CANTONIMS_DBDataSet.PURCHASE)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.ORDERS' table. You can move, or remove it, as needed.
        Me.ORDERSTableAdapter.Fill(Me.CANTONIMS_DBDataSet.ORDERS)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.PRODUCT' table. You can move, or remove it, as needed.
        Me.PRODUCTTableAdapter.Fill(Me.CANTONIMS_DBDataSet.PRODUCT)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.SUPPLIER' table. You can move, or remove it, as needed.
        Me.SUPPLIERTableAdapter.Fill(Me.CANTONIMS_DBDataSet.SUPPLIER)
        'TODO: This line of code loads data into the 'CANTONIMS_DBDataSet.USERS' table. You can move, or remove it, as needed.
        Me.USERSTableAdapter.Fill(Me.CANTONIMS_DBDataSet.USERS)

        If connection4.State = ConnectionState.Open Then
            connection4.Close()

        End If
        connection4.Open()

        display_Box_data()
        connection4.Close()

    End Sub


    Private Sub display_Box_data()

        If connection4.State = ConnectionState.Open Then
            connection4.Close()

        End If
        connection4.Open()

        ListBox1.Items.Add(cbuserlist.SelectedValue.ToString())
        ListBox2.Items.Add(cbproductlist.SelectedValue.ToString())
        ListBox3.Items.Add(cbsupplierlist.SelectedValue.ToString())
        connection4.Close()



    End Sub

    Private Sub display_PanelGridData1()
        If connection4.State = ConnectionState.Open Then
            connection4.Close()

        End If
        connection4.Open()

        command4 = connection4.CreateCommand()
        command4.CommandType = CommandType.Text
        command4.CommandText = "select  OrdNo FROM ORDERS"
        command4.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(command4)
        da.Fill(dt)

        DataGridView1.DataSource = dt
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Private Sub display_PanelGridData2()
        If connection4.State = ConnectionState.Open Then
            connection4.Close()

        End If
        connection4.Open()

        command4 = connection4.CreateCommand()
        command4.CommandType = CommandType.Text
        command4.CommandText = "select  PurchNo FROM PURCHASE"
        command4.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(command4)
        da.Fill(dt)

        DataGridView2.DataSource = dt
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Private Sub display_PanelGridData3()
        If connection4.State = ConnectionState.Open Then
            connection4.Close()

        End If
        connection4.Open()

        command4 = connection4.CreateCommand()
        command4.CommandType = CommandType.Text
        command4.CommandText = "select  ProdNo FROM PRODUCT"
        command4.ExecuteNonQuery()

        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(command4)
        da.Fill(dt)

        DataGridView2.DataSource = dt
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Private Sub btnrefresh_Click(sender As Object, e As EventArgs) Handles btnrefresh.Click
        'delete old data and put new data to replace it 
        DataGridView1.Refresh()
        DataGridView2.Refresh()
        DataGridView3.Refresh()
    End Sub

    Private Sub btnoverall_Click(sender As Object, e As EventArgs) Handles btnoverall.Click
        display_PanelGridData1()
        display_Box_data()

    End Sub

    Private Sub btnClean_Click(sender As Object, e As EventArgs) Handles btnClean.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
    End Sub

    Private Sub btnHome_Click(sender As Object, e As EventArgs) Handles btnHome.Click
        WelcomeForm.Show()
        Hide()
    End Sub
End Class