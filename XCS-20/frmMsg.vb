Public Class frmMsg
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        KeyCylFree
        HoldDUTCylup()
        UnlockDoor
        PushOpenDoor()
        frmMain.ScanPSnFlag = False
        frmMain.InterruptFlag = False
        frmMain.TestAction = 0
    End Sub
End Class