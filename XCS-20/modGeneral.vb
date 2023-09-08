Imports System.Data.SqlClient
Module modGeneral
    Public Sub GetLastConfig()
        Dim Filenum As Integer
        Filenum = FreeFile()
        Dim tempcode As String
        Dim pos1, pos2, pos3, pos4, pos5 As Integer

        FileOpen(Filenum, INISTATUSPATH, OpenMode.Input)

        tempcode = LineInput(Filenum)
        FileClose(Filenum)

        pos1 = InStr(1, tempcode, ",")
        pos2 = InStr(pos1 + 1, tempcode, ",")
        pos3 = InStr(pos2 + 1, tempcode, ",")
        pos4 = InStr(pos3 + 1, tempcode, ",")

        LoadWOfrRFID.JobNos = Mid(tempcode, 1, pos1 - 1)
        LoadWOfrRFID.JobModelName = Mid(tempcode, pos1 + 1, (pos2 - pos1) - 1)
        LoadWOfrRFID.JobQTy = Mid(tempcode, pos2 + 1, (pos3 - pos2) - 1)
        LoadWOfrRFID.JobArticle = Model2Article(LoadWOfrRFID.JobModelName)
        LoadWOfrRFID.JobRFIDTag = Mid(tempcode, pos3 + 1, (pos4 - pos3) - 1)
        LoadWOfrRFID.JobUnitaryCount = Mid(tempcode, pos4 + 1)
    End Sub
    Public Function Model2Article(csmodel As String) As String
        Dim Ref As String
        On Error GoTo ErrorHandler
        Dim query = "Select * FROM Parameter WHERE ModelName = '" & csmodel & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)

        Ref = dt.Rows(0).Item("ArticleNos")
        Return Ref
ErrorHandler:
    End Function
    Public Function Article2Model(ArtNos As String) As String
        Dim Ref As String
        On Error GoTo ErrorHandler

        Dim query = "Select * FROM Parameter WHERE ArticleNOS = '" & ArtNos & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)

        Ref = dt.Rows(0).Item("MODELNAME")
        Return Ref
        Exit Function
ErrorHandler:
    End Function
    Public Function UpdateStnStatus() As Boolean
        Dim Filenum1 As Integer

        Filenum1 = FreeFile()
        FileOpen(Filenum1, INISTATUSPATH, OpenMode.Output)
        PrintLine(Filenum1, LoadWOfrRFID.JobNos & "," & LoadWOfrRFID.JobModelName & "," & LoadWOfrRFID.JobQTy & "," & LoadWOfrRFID.JobRFIDTag & "," & LoadWOfrRFID.JobUnitaryCount)
        FileClose(Filenum1)
    End Function
    Public Function CheckWOExist(WO As String) As Boolean
        Dim temp As String
        On Error GoTo ErrorHandler

        Dim query = "Select * From STN5 Where WONOS = '" & WO & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)

        temp = dt.Rows(0).Item("WONOS")
        If temp <> "" Then
            Return True
        End If
ErrorHandler:
    End Function
    Public Function UpdateWO(WO As String, updateqty As String) As Boolean
        On Error GoTo ErrorHandler
        Dim konek As New SqlConnection(Database)
        Dim sql As String = "Update STN5 Set OUTPUT = '" & updateqty & "' Where WONOS = '" & WO & "'"
        konek.Open()
        Dim sc As New SqlCommand(sql, konek)
        Dim adapter As New SqlDataAdapter(sc)
        If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
            'MsgBox("Saving succeed!")
            konek.Close()
        Else
            MsgBox("Saving Failed!")
        End If
        Return True
ErrorHandler:
        MsgBox("Error Update WO!")
        Return False
    End Function
    Public Function AddWO(WO As String) As Boolean
        On Error GoTo ErrorHandler
        Dim konek As New SqlConnection(Database)
        Dim sql As String = "INSERT INTO STN5 (WONOS,OUTPUT) VALUES ('" & WO & "', 0)"
        konek.Open()
        Dim sc As New SqlCommand(sql, konek)
        Dim adapter As New SqlDataAdapter(sc)
        If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
            MsgBox("Saving succeed!")
            konek.Close()
        Else
            MsgBox("Saving Failed!")
        End If
        Return True
ErrorHandler:
        MsgBox("Error AddWO!")
        Return False
    End Function
    Public Function RetrieveWOQty(WO As String) As String
        Dim readqty As String
        Dim query = "Select * FROM STN5 WHERE WONOS = '" & WO & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)
        readqty = dt.Rows(0).Item("OUTPUT")
        Return readqty
    End Function
    Public Function Bin162Dec(Data As String) As Long
        Dim Expon As Integer
        Dim Sumtotal As Double

        For i = 1 To 16
            If Mid(Data, i, 1) = "1" Then
                Expon = Math.Abs(i - 16)
                Sumtotal = Sumtotal + 2 ^ Expon
            End If
        Next
        Return Sumtotal
    End Function
End Module
