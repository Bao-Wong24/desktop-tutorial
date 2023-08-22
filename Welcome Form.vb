Public Class WelcomeForm

    Private Sub btnproduct_Click(sender As Object, e As EventArgs) Handles btnproduct.Click
        ProductForm.Show()
        Hide()
    End Sub

    Private Sub btnuser_Click(sender As Object, e As EventArgs) Handles btnuser.Click
        Users_Main_Form.Show()
        Hide()
    End Sub

    Private Sub btnsupplier_Click(sender As Object, e As EventArgs) Handles btnsupplier.Click
        Supply_Main_Form.Show()
        Hide()
    End Sub

    Private Sub btnpurchase_Click(sender As Object, e As EventArgs) Handles btnpurchase.Click
        Purchase_Main_Form.Show()
        Hide()
    End Sub

    Private Sub btndashboard_Click(sender As Object, e As EventArgs) Handles btndashboard.Click
        Dashboard.Show()
        Hide()
    End Sub

    Private Sub btnexit_Click_1(sender As Object, e As EventArgs) Handles btnexit.Click
        'exit out of the application and have to run it again 
        Close()
    End Sub

    Private Sub btnlogout_Click(sender As Object, e As EventArgs) Handles btnlogout.Click
        'close and show logon form
        Close()
        LoginForm.Show()
    End Sub

    Private Sub btncategory_Click(sender As Object, e As EventArgs) Handles btncategory.Click
        Category_Main_Form.Show()
        Hide()
    End Sub

    Private Sub btnorder_Click(sender As Object, e As EventArgs) Handles btnorder.Click
        OrderForm.Show()
        Hide()
    End Sub
End Class
