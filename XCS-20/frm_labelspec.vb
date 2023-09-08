Public Class frm_labelspec
    Private Sub frm_labelspec_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LOADREFERENCE()
    End Sub
    Private Sub LOADREFERENCE()
        Dim query = "SELECT ModelName FROM Label"
        Dim dt = koneksiDB.bacaData(query).Tables(0)

        For i As Integer = 0 To dt.Rows.Count - 1
            Cmbo_Select.Items.Add(dt.Rows(i).Item(0))
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub Cmbo_Select_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmbo_Select.SelectedIndexChanged
        On Error Resume Next
        If Cmbo_Select.Text = "" Then Exit Sub
        CLEARDATA()
        Dim query = "Select * FROM Label WHERE ModelName = '" & Cmbo_Select.Text & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)
        lbl_Detail1.Text = dt.Rows(0).Item("Product_Detail1")
        lbl_Detail2.Text = dt.Rows(0).Item("Product_Detail2")
        lbl_Detail3.Text = dt.Rows(0).Item("Product_Detail3")
        lbl_Detail4.Text = dt.Rows(0).Item("Product_Detail4")
        lbl_Detail5.Text = dt.Rows(0).Item("Product_Detail5")
        lbl_Detail6.Text = dt.Rows(0).Item("Product_Detail6")
        lbl_Detail7.Text = dt.Rows(0).Item("Product_Detail7")
        lbl_Detail8.Text = dt.Rows(0).Item("Product_Detail8")
        lbl_Detail9.Text = dt.Rows(0).Item("Product_Torque")
        lbl_Detail10.Text = dt.Rows(0).Item("Product_Detail10")
        TextBox3.Text = dt.Rows(0).Item("ArticleNos")
        TextBox1.Text = dt.Rows(0).Item("Product_Template")
        TextBox2.Text = dt.Rows(0).Item("Product_Reference")
        lbl_Detail0.Text = dt.Rows(0).Item("Product_Reference")
    End Sub
    Private Sub CLEARDATA()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub
End Class