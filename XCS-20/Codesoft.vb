Imports Tkx.Lppa
Module Codesoft
    Public PrinterApp As Tkx.Lppa.Application = Nothing
    Public ActiveDoc As Tkx.Lppa.Document = Nothing

    Public Sub OpenCodesoft()
        PrinterApp = New Tkx.Lppa.Application
    End Sub

    Public Sub CloseCodesoft()
        If PrinterApp IsNot Nothing Then
            If PrinterApp.Documents IsNot Nothing Then
                PrinterApp.Documents.CloseAll(False)
                PrinterApp.Quit()
            End If
        End If
    End Sub

    Public Function OpenDocument(location As String) As Boolean
        Try
            ActiveDoc = PrinterApp.Documents.Open(location, True)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub CloseDocument()
        If ActiveDoc IsNot Nothing Then
            ActiveDoc.Close(False)
        End If
    End Sub

    Public Function SetPrinter(PrinterName As String, Portname As String) As Boolean
        Try
            If ActiveDoc IsNot Nothing Then
                ActiveDoc.Printer.SwitchTo(PrinterName, Portname, False)
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function PrintLab(Qty As Integer) As Boolean
        Try
            If ActiveDoc IsNot Nothing Then
                ActiveDoc.FullClippingMode = True
                If ActiveDoc IsNot Nothing Then
                    ActiveDoc.PrintDocument(Qty)
                End If
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub killLPPA()
        For Each p As Process In System.Diagnostics.Process.GetProcessesByName("Lppa")
            p.Kill()
        Next
    End Sub
End Module
