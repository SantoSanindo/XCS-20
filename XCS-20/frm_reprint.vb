Imports System.Threading
Public Class frm_reprint
    Dim reprintBuf As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MSComm.Close()
        'frmMain.Barcode_Comm.Open()
        Me.Close()
    End Sub
    Private Sub frm_reprint_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MSComm.Open()
    End Sub
    Private Sub DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles MSComm.DataReceived
        reprintBuf = MSComm.ReadExisting()
        If InStr(1, reprintBuf, vbCrLf) <> 0 Then
            Me.Invoke(Sub()
                          reprintBuf = Mid(reprintBuf, 1, InStr(1, reprintBuf, vbCr) - 1)
                          TextBox1.Text = reprintBuf

                          If Dir(INIPSNFOLDERPATH & reprintBuf & ".Txt") = "" Then
                              If Dir(INIACHIEVEPATH & reprintBuf & ".Txt") = "" Then
                                  Label1.Text = "--> Unable to find" & reprintBuf & ".Txt" & vbCrLf
                                  reprintBuf = ""
                                  Exit Sub
                              End If
                          End If
                          rprintlabel.JobArticle = Microsoft.VisualBasic.Strings.Left(reprintBuf, 6)
                          rprintlabel.JobModelName = Article2Model(rprintlabel.JobArticle)
                          ReprintlabelVar = LoadLabelParameter(rprintlabel.JobModelName)

                          If Not OpenDocument(INITEMPLATEPATH & ReprintlabelVar.UnitLabelTemplate) Then
                              Label1.Text = "--> Unable to open label template" & vbCrLf
                              reprintBuf = ""
                              Exit Sub
                          End If

                          If Not SetPrinter("Zebra 110XiIII Plus (600 dpi)", "USB001") Then
                              'If Not PrinterSelect("Zebra 110XiIII Plus (600 dpi)", "LPT1:") Then
                              Label1.Text = "--> Unable to select Printer" & vbCrLf
                              reprintBuf = ""
                              Exit Sub
                          End If

                          ActiveDoc.Variables.FormVariables.Item("Var14").Value = INIPHOTOPATH & ReprintlabelVar.UnitImage
                          ActiveDoc.Variables.FormVariables.Item("Var10").Value = reprintBuf
                          ActiveDoc.Variables.FormVariables.Item("Var8").Value = "20" & Mid(reprintBuf, 10, 2)
                          ActiveDoc.Variables.FormVariables.Item("Var9").Value = Mid(reprintBuf, 12, 2)

                          If PrintLab(1) = False Then     'necessary for printing
                              MsgBox("Error", vbCritical)
                          End If

                          CloseCodesoft()
                          reprintBuf = ""
                      End Sub)
        End If
    End Sub
    Private Function LoadLabelParameter(csmodel As String) As LabelData
        On Error Resume Next
        Dim query = "Select * FROM Label WHERE ModelName = '" & csmodel & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)

        LoadLabelParameter.UnitModelName = "" & dt.Rows(0).Item("Product_Reference")
        LoadLabelParameter.UnitArticleNos = "" & dt.Rows(0).Item("ArticleNos")
        LoadLabelParameter.UnitImage = "" & dt.Rows(0).Item("Product_Img")
        LoadLabelParameter.UnitLabelTemplate = "" & dt.Rows(0).Item("Product_Template")
        LoadLabelParameter.UnitLabelSymb1 = "" & dt.Rows(0).Item("Product_Symbol1")
        LoadLabelParameter.UnitLabelSymb2 = "" & dt.Rows(0).Item("Product_Symbol2")
        LoadLabelParameter.UnitLabelTor = "" & dt.Rows(0).Item("Product_Torque")
ErrorHandler:
    End Function
End Class