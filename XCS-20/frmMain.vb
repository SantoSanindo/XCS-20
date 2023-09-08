Imports System.Threading
Public Class frmMain
    Dim ChromaFBbuf As String
    Dim ChromaBuf As String
    Dim AssyBuf As String
    Dim CheckWO As String
    Dim currentDate As DateTime = DateTime.Now
    Dim ReadTagFlag As Boolean 'True when Timer1 is reading Tag
    Public Interruptread As Integer
    Public TestAction As Integer
    Public ScanPSnFlag As Boolean
    Public InterruptFlag As Boolean
    Dim TeststatusFlag As Boolean
    Dim TimeoutCount As Integer
    Public AssyAction As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fullPath As String = System.AppDomain.CurrentDomain.BaseDirectory
        Dim projectFolder As String = fullPath.Replace("\XCS-20\bin\Debug\", "").Replace("\XCS-20\bin\Release\", "")
        If Dir(projectFolder & "\Config\Config.INI") = "" Then 'This is to initalize the program during start up
            MsgBox("Config.INI is missing")
            End
        End If

        ReadINI(projectFolder & "\Config\Config.INI")
        GetLastConfig()

        killLPPA()
        OpenCodesoft()

        ' --> Need PCI_L112.dll file
        If Not INIT_PCI112() Then
            MsgBox("Unable to initialize Syntek")
            End
        End If

        Chroma_Comm.Open()
        If Not Init_Chroma() Then
            MsgBox("Unable to communicate with Chroma 19053")
            End
        End If

        If Not Set_Chroma_HipotTest() Then
            MsgBox("Unable to configure Chroma 19053")
            End
        End If

        lbl_WOnos.Text = LoadWOfrRFID.JobNos
        lbl_currentref.Text = LoadWOfrRFID.JobModelName
        lbl_wocounter.Text = LoadWOfrRFID.JobQTy
        lbl_unitaryCount.Text = LoadWOfrRFID.JobUnitaryCount
        lbl_tagnos.Text = LoadWOfrRFID.JobRFIDTag
        lbl_ArticleNos.Text = LoadWOfrRFID.JobArticle

        'Load Parameter from database - Server
        If Not LoadParameter(LoadWOfrRFID.JobModelName) Then
            MsgBox("Unable to load parameters from Database")
            End
        End If

        If Not UpdateParameter2Tester() Then
            MsgBox("Unable to update parameters to tester")
            End
        End If

        'Load label parameter
        LabelVar = LoadLabelParameter(LoadWOfrRFID.JobModelName)

        'Load Rack Material List
        If Not LoadRackMaterial() Then
            MsgBox("Unable to load Rack Materials")
            End
        End If

        'Load Model Material
        If Not LoadModelMaterial(LoadWOfrRFID.JobModelName) Then
            MsgBox("Unable to load Model Material")
            End
        End If

        'Update Rack indicator --> Need MotionNet.dll file
        If Not ActivateRackLED() Then
            MsgBox("Unable to communicate with PLC")
            End
        End If

        RFID_Comm.Open()
        Barcode_Comm.Open()
        Timer1.Enabled = True
    End Sub
    Private Function Init_Chroma() As Boolean
        If Not Set_chroma("*IDN?", "Chroma,19053,0,4.30") Then
            Return False
            Exit Function
        End If
        If Not Set_chroma("SYST:LOCK:REQ?", "1") Then
            Return False
            Exit Function
        End If
        Return True
    End Function
    Public Function Set_chroma(Command As String, Reply As String) As Boolean
        Dim Retry As Integer
        Dim Timeout As Double

Retry:
        Timeout = 0
        Chroma_Comm.Write(Command & vbCrLf)
        If Reply <> "" Then
            Do While ChromaFBbuf <> Reply
                Timeout = Timeout + 1
                If Timeout > 60000 Then
                    Retry = Retry + 1
                    If Retry > 4 Then
                        Return False
                        Exit Function
                    Else
                        GoTo Retry
                    End If
                End If
                System.Windows.Forms.Application.DoEvents()
            Loop
        End If
        Return True
    End Function
    Public Function Set_Chroma_HipotTest() As Boolean
        Dim Timeout As Double
        Dim Retry As Integer
100:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC 1000" & vbCrLf)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC?" & vbCrLf)
        Do While ChromaFBbuf <> "+1.000000E+03"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 100
                End If
            End If
        Loop

        Retry = 0
200:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:LIM 0.005" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:LIM?" & vbCrLf)
        Do While ChromaFBbuf <> "+5.000001E-03"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 200
                End If
            End If
        Loop

        Retry = 0
300:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME:RAMP 0.5" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME:RAMP?" & vbCrLf)
        Do While ChromaFBbuf <> "+5.000000E-01"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 300
                End If
            End If
        Loop

        Retry = 0
400:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME 1" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:TIME?" & vbCrLf)
        Do While ChromaFBbuf <> "+1.000000E+00"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 400
                End If
            End If
        Loop

500:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""

        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:CHAN(@(1))" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:CHAN?" & vbCrLf)

        Do While ChromaFBbuf <> "(@(1))"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 500
                End If
            End If
        Loop

        Retry = 0
600:
        Timeout = 0
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""

        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:CHAN:LOW(@(2))" & vbCrLf)
        Thread.Sleep(10)
        Chroma_Comm.Write("SOUR:SAFE:STEP 1:AC:CHAN:LOW?" & vbCrLf)
        Do While ChromaFBbuf <> "(@(2))"
            System.Windows.Forms.Application.DoEvents()
            Timeout = Timeout + 1
            If Timeout > 60000 Then
                Retry = Retry + 1
                If Retry > 4 Then
                    GoTo FailComm
                Else
                    Chroma_Comm.DiscardInBuffer()
                    GoTo 600
                End If
            End If
        Loop

        Return True
        Exit Function
FailComm:
        Return False
    End Function
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Barcode_Comm.Close()
        Chroma_Comm.Close()
        RFID_Comm.Close()
        killLPPA()

        Me.Close()
    End Sub
    Dim Unit As ProductData
    Private Function LoadParameter(csmodel As String) As Boolean
        On Error Resume Next
        Dim query = "Select * FROM Parameter WHERE ModelName = '" & csmodel & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)
        ReDim Unit.UnitContact_original(6)
        ReDim Unit.UnitContactOriginaltemp(6)
        ReDim Unit.UnitContact_Wkey(6)
        ReDim Unit.UnitcontactWkeytemp(6)
        ReDim Unit.UnitContact_WkeyTension(6)
        ReDim Unit.UnitContactWkeyTensiontemp(6)

        LoadWOfrRFID.JobArticle = dt.Rows(0).Item("ArticleNos")
        LoadWOfrRFID.JobProductThread = "" & dt.Rows(0).Item("BodyType")
        LoadWOfrRFID.JobButtonType = "" & dt.Rows(0).Item("ButtonOpt")
        LoadWOfrRFID.JobProductMaterial = "" & dt.Rows(0).Item("MaterialType")

        DUTParameter.UnitTension = "" & dt.Rows(0).Item("TensionType")
        DUTParameter.UnitFunctionType = "" & dt.Rows(0).Item("FunctionType")
        DUTParameter.UnitS11ContactNos = "" & dt.Rows(0).Item("S11")
        DUTParameter.UnitS12ContactNos = "" & dt.Rows(0).Item("S12")
        DUTParameter.UnitS21ContactNos = "" & dt.Rows(0).Item("S21")
        DUTParameter.UnitS22ContactNos = "" & dt.Rows(0).Item("S22")
        DUTParameter.UnitS31ContactNos = "" & dt.Rows(0).Item("S31")
        DUTParameter.UnitS32ContactNos = "" & dt.Rows(0).Item("S32")
        DUTParameter.UnitS41ContactNos = "" & dt.Rows(0).Item("S41")
        DUTParameter.UnitS42ContactNos = "" & dt.Rows(0).Item("S42")
        DUTParameter.UnitS51ContactNos = "" & dt.Rows(0).Item("S51")
        DUTParameter.UnitS52ContactNos = "" & dt.Rows(0).Item("S52")
        DUTParameter.UnitS61ContactNos = "" & dt.Rows(0).Item("S61")
        DUTParameter.UnitS62ContactNos = "" & dt.Rows(0).Item("S62")
        DUTParameter.UnitS11PN = "" & dt.Rows(0).Item("S11PN")
        DUTParameter.UnitS12PN = "" & dt.Rows(0).Item("S12PN")
        DUTParameter.UnitS21PN = "" & dt.Rows(0).Item("S21PN")
        DUTParameter.UnitS22PN = "" & dt.Rows(0).Item("S22PN")
        DUTParameter.UnitS31PN = "" & dt.Rows(0).Item("S31PN")
        DUTParameter.UnitS32PN = "" & dt.Rows(0).Item("S32PN")
        DUTParameter.UnitS41PN = "" & dt.Rows(0).Item("S41PN")
        DUTParameter.UnitS42PN = "" & dt.Rows(0).Item("S42PN")
        DUTParameter.UnitS51PN = "" & dt.Rows(0).Item("S51PN")
        DUTParameter.UnitS52PN = "" & dt.Rows(0).Item("S52PN")
        DUTParameter.UnitS61PN = "" & dt.Rows(0).Item("S61PN")
        DUTParameter.UnitS62PN = "" & dt.Rows(0).Item("S62PN")
        DUTParameter.UnitX1PN = "" & dt.Rows(0).Item("X1PN")
        DUTParameter.UnitX2PN = "" & dt.Rows(0).Item("X2PN")
        DUTParameter.UnitE1PN = "" & dt.Rows(0).Item("E1PN")
        DUTParameter.UnitE2PN = "" & dt.Rows(0).Item("E2PN")
        DUTParameter.UnitS1CC = "" & dt.Rows(0).Item("S1CC")
        DUTParameter.UnitS2CC = "" & dt.Rows(0).Item("S2CC")
        DUTParameter.UnitS3CC = "" & dt.Rows(0).Item("S3CC")
        DUTParameter.UnitS4CC = "" & dt.Rows(0).Item("S4CC")
        DUTParameter.UnitS5CC = "" & dt.Rows(0).Item("S5CC")
        DUTParameter.UnitS6CC = "" & dt.Rows(0).Item("S6CC")
        DUTParameter.UnitX1CC = "" & dt.Rows(0).Item("X1CC")
        DUTParameter.UnitX2CC = "" & dt.Rows(0).Item("X2CC")
        DUTParameter.UnitE1CC = "" & dt.Rows(0).Item("E1CC")
        DUTParameter.UnitE2CC = "" & dt.Rows(0).Item("E2CC")
        Unit.UnitContact_original(1) = "" & dt.Rows(0).Item("Contact1Type")
        Unit.UnitContact_original(2) = "" & dt.Rows(0).Item("Contact2Type")
        Unit.UnitContact_original(3) = "" & dt.Rows(0).Item("Contact3Type")
        Unit.UnitContact_original(4) = "" & dt.Rows(0).Item("Contact4Type")
        Unit.UnitContact_original(5) = "" & dt.Rows(0).Item("Contact5Type")
        Unit.UnitContact_original(6) = "" & dt.Rows(0).Item("Contact6Type")
        DUTParameter.UnitOriginalSate = ""

        For i = 1 To 6
            If Unit.UnitContact_Wkey(i) = "None" Then
                DUTParameter.UnitStatewkey = DUTParameter.UnitStatewkey & "0"
                Unit.UnitcontactWkeytemp(i) = "0"
            Else
                DUTParameter.UnitStatewkey = DUTParameter.UnitStatewkey & Unit.UnitContact_Wkey(i)
                Unit.UnitcontactWkeytemp(i) = Unit.UnitContact_Wkey(i)
            End If
        Next

        Unit.UnitContact_WkeyTension(1) = "" & dt.Rows(0).Item("Contact1_W_Key_Ten")
        Unit.UnitContact_WkeyTension(2) = "" & dt.Rows(0).Item("Contact2_W_Key_Ten")
        Unit.UnitContact_WkeyTension(3) = "" & dt.Rows(0).Item("Contact3_W_Key_Ten")
        Unit.UnitContact_WkeyTension(4) = "" & dt.Rows(0).Item("Contact4_W_Key_Ten")
        Unit.UnitContact_WkeyTension(5) = "" & dt.Rows(0).Item("Contact5_W_Key_Ten")
        Unit.UnitContact_WkeyTension(6) = "" & dt.Rows(0).Item("Contact6_W_Key_Ten")
        Unit.UnitStatewkeyntension = ""

        For i = 1 To 6
            If Unit.UnitContact_WkeyTension(i) = "None" Then
                DUTParameter.UnitStatewkeyntension = DUTParameter.UnitStatewkeyntension & "0"
                Unit.UnitContactWkeyTensiontemp(i) = "0"
            Else
                DUTParameter.UnitStatewkeyntension = DUTParameter.UnitStatewkeyntension & Unit.UnitContact_WkeyTension(i)
                Unit.UnitContactWkeyTensiontemp(i) = Unit.UnitContact_WkeyTension(i)
            End If
        Next

        Return True
ErrorHandler:
    End Function
    Private Function UpdateParameter2Tester() As Boolean
        If Unit.UnitContact_original(1) = "None" Then
            Line1.Visible = False
            Line2.Visible = False
        Else
            Line1.Visible = True
            Line2.Visible = True
        End If
        If Unit.UnitContact_original(2) = "None" Then
            Line3.Visible = False
            Line4.Visible = False
        Else
            Line3.Visible = True
            Line4.Visible = True
        End If
        If Unit.UnitContact_original(3) = "None" Then
            Line5.Visible = False
            Line6.Visible = False
        Else
            Line5.Visible = True
            Line6.Visible = True
        End If
        If Unit.UnitContact_original(4) = "None" Then
            Line7.Visible = False
            Line8.Visible = False
        Else
            Line7.Visible = True
            Line8.Visible = True
        End If
        If Unit.UnitContact_original(5) = "None" Then
            Line9.Visible = False
            Line10.Visible = False
        Else
            Line9.Visible = True
            Line10.Visible = True
        End If
        If Unit.UnitContact_original(6) = "None" Then
            Line11.Visible = False
            Line12.Visible = False
        Else
            Line11.Visible = True
            Line12.Visible = True
        End If
        Txt_contactNos1.Text = DUTParameter.UnitS11ContactNos
        Txt_contactNos2.Text = DUTParameter.UnitS12ContactNos
        Txt_contactNos3.Text = DUTParameter.UnitS21ContactNos
        Txt_contactNos4.Text = DUTParameter.UnitS22ContactNos
        Txt_contactNos5.Text = DUTParameter.UnitS31ContactNos
        Txt_contactNos6.Text = DUTParameter.UnitS32ContactNos
        Txt_contactNos7.Text = DUTParameter.UnitS41ContactNos
        Txt_contactNos8.Text = DUTParameter.UnitS42ContactNos
        Txt_contactNos9.Text = DUTParameter.UnitS51ContactNos
        Txt_contactNos10.Text = DUTParameter.UnitS52ContactNos
        Txt_contactNos11.Text = DUTParameter.UnitS61ContactNos
        Txt_contactNos12.Text = DUTParameter.UnitS62ContactNos
        Txt_ContactPN1.Text = DUTParameter.UnitS11PN
        Txt_ContactPN2.Text = DUTParameter.UnitS12PN
        Txt_ContactPN3.Text = DUTParameter.UnitS21PN
        Txt_ContactPN4.Text = DUTParameter.UnitS22PN
        Txt_ContactPN5.Text = DUTParameter.UnitS31PN
        Txt_ContactPN6.Text = DUTParameter.UnitS32PN
        Txt_ContactPN7.Text = DUTParameter.UnitS41PN
        Txt_ContactPN8.Text = DUTParameter.UnitS42PN
        Txt_ContactPN9.Text = DUTParameter.UnitS51PN
        Txt_ContactPN10.Text = DUTParameter.UnitS52PN
        Txt_ContactPN11.Text = DUTParameter.UnitS61PN
        Txt_ContactPN12.Text = DUTParameter.UnitS62PN
        Txt_ContactPN13.Text = DUTParameter.UnitX1PN
        Txt_ContactPN14.Text = DUTParameter.UnitX2PN
        Txt_ContactPN15.Text = DUTParameter.UnitE1PN
        Txt_ContactPN16.Text = DUTParameter.UnitE2PN
        Line1.BackColor = CableColorCode(DUTParameter.UnitS1CC)
        Line2.BackColor = CableColorCode(DUTParameter.UnitS1CC)
        Line3.BackColor = CableColorCode(DUTParameter.UnitS2CC)
        Line4.BackColor = CableColorCode(DUTParameter.UnitS2CC)
        Line5.BackColor = CableColorCode(DUTParameter.UnitS3CC)
        Line6.BackColor = CableColorCode(DUTParameter.UnitS3CC)
        Line7.BackColor = CableColorCode(DUTParameter.UnitS4CC)
        Line8.BackColor = CableColorCode(DUTParameter.UnitS4CC)
        Line9.BackColor = CableColorCode(DUTParameter.UnitS5CC)
        Line10.BackColor = CableColorCode(DUTParameter.UnitS5CC)
        Line11.BackColor = CableColorCode(DUTParameter.UnitS6CC)
        Line12.BackColor = CableColorCode(DUTParameter.UnitS6CC)
        Line13.BackColor = CableColorCode(DUTParameter.UnitX1CC)
        Line14.BackColor = CableColorCode(DUTParameter.UnitX2CC)
        Line15.BackColor = CableColorCode(DUTParameter.UnitE1CC)
        Line16.BackColor = CableColorCode(DUTParameter.UnitE2CC)
        Return True
    End Function
    Private Function CableColorCode(Colorcode As String) As Color
        Select Case Colorcode
            Case "Black"
                Return Color.Black
            Case "White"
                Return Color.White
            Case "Blue"
                Return Color.Blue
            Case "Green"
                Return Color.Green
            Case "Orange"
                Return Color.Orange
            Case "Pink"
                Return Color.Pink
            Case "Red"
                Return Color.Red
            Case "Yellow"
                Return Color.Yellow
        End Select

    End Function
    Private Function LoadLabelParameter(csmodel As String) As LabelData
        On Error Resume Next
        Dim query = "Select * FROM Label WHERE ModelName = '" & csmodel & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)

        LoadLabelParameter.UnitModelName = "" & dt.Rows(0).Item("Product_Reference")
        LoadLabelParameter.UnitArticleNos = "" & dt.Rows(0).Item("ArticleNos")
        LoadLabelParameter.UnitImage = "" & dt.Rows(0).Item("Conn_Img")
        LoadLabelParameter.UnitLabelTemplate = "" & dt.Rows(0).Item("Conn_Template")

        Exit Function
    End Function
    Dim Part As RackConfig
    Dim Job As JobOrder
    Private Function LoadRackMaterial() As Boolean
        Dim FNum As Integer
        Dim LineStr As String
        Dim i As Integer
        ReDim Part.PartPLCWord(45)
        ReDim Part.PartNos(45)

        On Error GoTo ErrorHandler

        FNum = FreeFile()

        FileOpen(FNum, INIMATERIALPATH & "Rack\" & "SA_Connector", OpenMode.Input)
        Do While Not EOF(FNum)
            LineStr = LineInput(FNum)
            Part.PartNos(i) = LineStr
            Part.PartPLCWord(i) = 40200 + i
            i = i + 1
        Loop
        FileClose(FNum)
        Return True
ErrorHandler:
        Return False
    End Function
    Private Function LoadModelMaterial(Unitname As String) As Boolean
        Dim FNum As Integer
        Dim LineStr As String
        Dim i As Integer
        ReDim Job.JobModelMaterial(45)

        On Error GoTo ErrorHandler

        FNum = FreeFile()

        FileOpen(FNum, INIMATERIALPATH & "SA_Connector\" & Unitname & ".Txt", OpenMode.Input)
        Do While Not EOF(FNum)
            LineStr = LineInput(FNum)
            Job.JobModelMaterial(i) = LineStr
            i = i + 1
        Loop
        FileClose(FNum)
        Return True
ErrorHandler:
        Return False

    End Function
    Private Function ActivateRackLED() As Boolean
        Dim i, N As Integer
        ResetMaterialIO()
        On Error GoTo ErrorHandler
        For i = 0 To 40
            If Job.JobModelMaterial(i) <> "" Then
                For N = 0 To 6
                    If Job.JobModelMaterial(i) = Part.PartNos(N) Then
                        Call SetMaterialIO(N)
                    End If
                Next
            End If
        Next
        If LoadWOfrRFID.JobProductMaterial = "Zamak" Then
            SetIO(2, 3, 0)
            SetIO(2, 3, 1)
            SetIO(2, 3, 2)
            UnSetIO(2, 3, 3)
        Else
            SetIO(2, 3, 0)
            SetIO(2, 3, 1)
            SetIO(2, 3, 3)
            UnSetIO(2, 3, 2)
        End If
        Return True
ErrorHandler:
        Return False
    End Function
    Private Sub ResetMaterialIO()
        UnSetIO(2, 1, 3)
        UnSetIO(2, 2, 2)

        UnSetIO(2, 1, 2)
        UnSetIO(2, 2, 3)

        UnSetIO(2, 1, 1)
        UnSetIO(2, 2, 1)

        UnSetIO(2, 1, 0)
        UnSetIO(2, 2, 0)

        UnSetIO(2, 1, 5)
        UnSetIO(2, 2, 5)

        UnSetIO(2, 1, 4)
        UnSetIO(2, 2, 4)
    End Sub
    Private Function SetMaterialIO(PartNos As Integer)
        Select Case PartNos

            Case 0
                SetIO(2, 1, 3) 'LED
                SetIO(2, 2, 2) 'Door

            Case 1
                SetIO(2, 1, 2) 'LED
                SetIO(2, 2, 3) 'Door
            Case 2
                SetIO(2, 1, 1) 'LED
                SetIO(2, 2, 1) 'Door
            Case 3
                SetIO(2, 1, 0) 'LED
                SetIO(2, 2, 0) 'Door
            Case 4
                SetIO(2, 1, 5) 'LED
                SetIO(2, 2, 5) 'Door
            Case 5
                SetIO(2, 1, 4)
                SetIO(2, 2, 4) 'Door
        End Select
    End Function
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        frm_labelspec.Show()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Barcode_Comm.Close()
        frm_reprint.Show()
    End Sub
    Private Sub ChromaDataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles Chroma_Comm.DataReceived
        ChromaBuf = Chroma_Comm.ReadExisting()
        If InStr(1, ChromaBuf, vbCrLf) <> 0 Then
            ChromaBuf = Mid(ChromaBuf, 1, InStr(1, ChromaBuf, vbCrLf) - 1)
            Me.Invoke(Sub()
                          ChromaFBbuf = ChromaBuf
                          ChromaBuf = ""
                      End Sub)
        End If
    End Sub

    Private Sub BarcodeDataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles Barcode_Comm.DataReceived
        AssyBuf = Barcode_Comm.ReadExisting()
        If InStr(1, AssyBuf, vbCrLf) <> 0 Then
            Me.Invoke(Sub()
                          Txt_Msg.BackColor = Color.DarkGray
                          Txt_Msg.Text = ""
                          AssyBuf = Mid(AssyBuf, 1, InStr(1, AssyBuf, vbCr) - 1)
                          Image2.Image = Nothing

                          Dim textBoxes As List(Of TextBox) = New List(Of TextBox) From {Txt_Originalstate1, Txt_Originalstate2, Txt_Originalstate3, Txt_Originalstate4, Txt_Originalstate5, Txt_Originalstate6, Txt_KeyState1, Txt_KeyState2, Txt_KeyState3, Txt_KeyState4, Txt_KeyState5, Txt_KeyState6, Txt_KeyTensionState1, Txt_KeyTensionState2, Txt_KeyTensionState3, Txt_KeyTensionState4, Txt_KeyTensionState5, Txt_KeyTensionState6}

                          For i As Integer = 0 To 17
                              textBoxes(i).Text = ""
                          Next
                          Txt_Msg.Text = ""

                          If lbl_WOnos.Text <> "MASTER" Then
                              'Checking Current WO first b4 Change Series is allowed. If WO status is open, check Quantity
                              CheckWO = ValidiateWONos(lbl_WOnos.Text)
                              If CheckWO = "NOK" Then
                                  Txt_Msg.Text = "Invalid WO - WO is not registered in Server"
                                  lbl_msg.Text = "PLEASE ASK FOR TECHNICAL ASSISTANCE"
                                  AssyBuf = ""
                                  Exit Sub
                              End If
                              If CheckWO <> "DISTRUP" Then
                                  'If Command1.Text = "Eye Open" Then
                                  '    If Val(lbl_unitaryCount.Text) >= (lbl_wocounter.Text) Then
                                  '        Txt_Msg.Text = "WO Quantity Reached - WO Completed"
                                  '        lbl_msg.Text = "STOP PROCESS"
                                  '       AssyBuf = ""
                                  '        Exit Sub
                                  '    End If
                                  'End If
                              Else
                                  Txt_Msg.Text = "WO Distrupted"
                                  lbl_msg.Text = "PLEASE CHANGE SERIES"
                                  AssyBuf = ""
                                  Exit Sub
                              End If
                          Else 'If MASTER TAG in use
                              'If Val(lbl_unitaryCount.Text) >= Val(lbl_wocounter.Text) Then
                              '    Txt_Msg.Text = "Quantity reached. WO Completed"
                              '    lbl_msg.Text = "STOP PROCESS"
                              '    AssyBuf = ""
                              '    Exit Sub
                              'End If
                          End If

                          If Microsoft.VisualBasic.Strings.Left(AssyBuf, 6) <> LoadWOfrRFID.JobArticle Then
                              Txt_Msg.Text = "--> PSN - " & AssyBuf & " = wrong reference"
                              Image2.Image = My.Resources.ResourceManager.GetObject("Pass")
                              AssyBuf = ""
                              Exit Sub
                          Else
                              Txt_Msg.Text = "PSN - " & AssyBuf & vbCrLf
                              PSNFileInfo.PSN = AssyBuf
                              Txt_Msg.Text = Txt_Msg.Text & "PSN Verified" & vbCrLf
                          End If

                          Txt_Msg.Text = Txt_Msg.Text & "Loading" & AssyBuf & ".Txt..." & vbCrLf

                          If Dir(INIPSNFOLDERPATH & AssyBuf & ".Txt") = "" Then
                              Txt_Msg.Text = Txt_Msg.Text & "--> Unable to find" & AssyBuf & ".Txt" & vbCrLf
                              Image2.Image = My.Resources.ResourceManager.GetObject("Fail")
                              AssyBuf = ""
                              Exit Sub
                          End If

                          If Not LOADPSNFILE(AssyBuf) Then
                              Txt_Msg.Text = Txt_Msg.Text & "--> Unable to load" & AssyBuf & ".Txt" & vbCrLf
                              Image2.Image = My.Resources.ResourceManager.GetObject("Fail")
                              AssyBuf = ""
                              Exit Sub
                          End If

                          PSNFileInfo.ConnTestCheckIn = currentDate.Date & "," & Now.ToLongTimeString
                          Txt_Msg.Text = Txt_Msg.Text & "Verifying PSN..." & vbCrLf
                          If PSNFileInfo.VacuumStatus <> "PASS" Then
                              Image2.Image = My.Resources.ResourceManager.GetObject("Fail")
                              Txt_Msg.Text = "--> PSN skip process" & vbCrLf
                              lbl_msg.Text = "PLEASE GO BACK TO PREVIOUS STATION" & vbCrLf
                              AssyBuf = ""
                              Exit Sub
                          Else
skip:
                              ScanPSnFlag = True
                              AssyBuf = ""
                              Exit Sub
                              'If PSNFileInfo.Stn5Status <> "PASS" Then 'Never Print before
                              If LabelVar.UnitLabelTemplate <> "" Then 'If no define template, skip print unitary
                                  If Not OpenDocument(INITEMPLATEPATH & LabelVar.UnitLabelTemplate) Then
                                      Txt_Msg.Text = "--> Unable to open label template" & vbCrLf
                                      Image2.Image = My.Resources.ResourceManager.GetObject("Fail")
                                      AssyBuf = ""
                                      Exit Sub
                                  End If
                                  Txt_Msg.Text = "Setting Printer..."
                                  If Not SetPrinter("Zebra 110XiIII Plus (600 dpi)", "USB001") Then
                                      'If Not PrinterSelect("Zebra 110XiIII Plus (600 dpi)", "LPT1:") Then
                                      Txt_Msg.Text = "--> Unable to select Printer" & vbCrLf
                                      Image2.Image = My.Resources.ResourceManager.GetObject("Fail")
                                      AssyBuf = ""
                                      Exit Sub
                                  End If


                                  If PSNFileInfo.ConnTestStatus <> "PASS" Then
                                      Txt_Msg.Text = "Printing label..."
                                      PrintLabel()
                                      lbl_unitaryCount.Text = Val(lbl_unitaryCount.Text) + 1
                                      LoadWOfrRFID.JobUnitaryCount = Val(lbl_unitaryCount.Text)
                                      UpdateStnStatus()
                                  Else
                                      Txt_Msg.Text = "Repeat label Print - No Printing"
                                  End If
                                  CloseDocument()
                                  Txt_Msg.Text = ""
                                  PSNFileInfo.Stn5Status = "PASS"
                                  PSNFileInfo.Stn5CheckOut = currentDate.Date & "," & Now.ToLongTimeString
                                  If Not WRITEPSNFILE(PSNFileInfo.PSN) Then
                                      Txt_Msg.Text = Txt_Msg.Text & "--> Unable to write " & PSNFileInfo.PSN & ".Txt in the server" & vbCrLf
                                      Image2.Image = My.Resources.ResourceManager.GetObject("Fail")
                                      AssyBuf = ""
                                      Exit Sub
                                  End If
                                  Image2.Image = My.Resources.ResourceManager.GetObject("Pass")

                                  If Val(lbl_unitaryCount.Text) >= Val(lbl_wocounter.Text) Then
                                      Txt_Msg.Text = "WO Quantity reached - WO Completed"
                                      lbl_msg.Text = "STOP PROCESS"
                                  End If
                              End If
                              'Else 'Printed Before
                              '    Txt_Msg.Text = "PSN - Already printed before"
                              '    Image2.Picture = LoadPicture(App.Path & "\Icons\PASS.Emf")
                              'End If
                          End If
                          AssyBuf = ""
                      End Sub)
        End If
    End Sub
    Private Function ValidiateWONos(strName As String) As String
        Dim temp As String
        On Error GoTo ErrorHandler
        Dim query = "Select * From CSUNIT Where WONOS = '" & strName & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)

        temp = dt.Rows(0).Item("STATUS")
        Return temp
ErrorHandler:
    End Function
    Public Function PrintLabel()
        ActiveDoc.Variables.FormVariables.Item("Var14").Value = INIPHOTOPATH & LabelVar.UnitImage
        If PrintLab(1) = False Then
            MsgBox("Error. Can't to print...", MsgBoxStyle.Critical)
        End If
    End Function

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        GetLastConfig()

        If Not LoadParameter(LoadWOfrRFID.JobModelName) Then
            MsgBox("Unable to load parameters from Database")
            Exit Sub
        End If

        If Not UpdateParameter2Tester() Then
            MsgBox("Unable to update parameters to tester")
            Exit Sub
        End If

        'Load label parameter
        LabelVar = LoadLabelParameter(LoadWOfrRFID.JobModelName)

        'Load Rack Material List
        If Not LoadRackMaterial() Then
            MsgBox("Unable to load Rack Materials")
            End
        End If

        'Load Model Material
        If Not LoadModelMaterial(LoadWOfrRFID.JobModelName) Then
            MsgBox("Unable to load Model Material")
            End
        End If

        'Update Rack indicator --> MotionNet.dll is missing
        If Not ActivateRackLED() Then
            MsgBox("Unable to communicate with PLC")
            End
        End If

        lbl_WOnos.Text = LoadWOfrRFID.JobNos
        lbl_currentref.Text = LoadWOfrRFID.JobModelName
        lbl_wocounter.Text = LoadWOfrRFID.JobQTy
        lbl_unitaryCount.Text = LoadWOfrRFID.JobUnitaryCount
        lbl_tagnos.Text = LoadWOfrRFID.JobRFIDTag
        lbl_ArticleNos.Text = LoadWOfrRFID.JobArticle

        Dim textBoxes As List(Of TextBox) = New List(Of TextBox) From {Txt_Originalstate1, Txt_Originalstate2, Txt_Originalstate3, Txt_Originalstate4, Txt_Originalstate5, Txt_Originalstate6, Txt_KeyState1, Txt_KeyState2, Txt_KeyState3, Txt_KeyState4, Txt_KeyState5, Txt_KeyState6, Txt_KeyTensionState1, Txt_KeyTensionState2, Txt_KeyTensionState3, Txt_KeyTensionState4, Txt_KeyTensionState5, Txt_KeyTensionState6}

        For i As Integer = 0 To 17
            textBoxes(i).Text = ""
        Next

        'Load Parameter from database - Server
        If Not LoadParameter(LoadWOfrRFID.JobModelName) Then
            MsgBox("Fail to load parameters from Server")
            Exit Sub
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = 0 Then
            HoldDUTCyldn()
        Else
            HoldDUTCylup()
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim Tagref As String
        Dim Tagnos As String
        Dim TagQty As String
        Dim Tagid As String
        Dim Test As ReadBackContact
        ReDim Test.KeyState(6)
        ReDim Test.KeyTension(6)
        ReDim Test.OriginState(6)

        ReadTagFlag = True
        'Ethernet.FillColor = vbBlack
        'GoTo NoChange
        Tagnos = RD_MULTI_RFID("0000", 10)

        If Tagnos = "NOK" Then GoTo NoChange
        Tagid = RD_MULTI_RFID("0040", 3) 'Read Tag ID
        If Tagnos = "MASTER" Then
            If Tagid = lbl_tagnos.Text Then GoTo NoChange
            If lbl_WOnos.Text <> "MASTER" Then
                'update the Current WO into the database before changing
                If CheckWOExist(lbl_WOnos.Text) Then
                    Call UpdateWO(lbl_WOnos.Text, lbl_unitaryCount.Text)
                Else
                    Call AddWO(lbl_WOnos.Text)
                    Call UpdateWO(lbl_WOnos.Text, lbl_unitaryCount.Text)
                End If
                GoTo WOChange
            ElseIf lbl_WOnos.Text = "MASTER" Then
                GoTo WOChange
            End If
        ElseIf Tagnos <> LoadWOfrRFID.JobNos Then
Master:
            If lbl_WOnos.Text <> "MASTER" Then
                'Checking Current WO first b4 Change Series is allowed. If WO status is open, check Quantity
                Dim CheckWO As String
                CheckWO = ValidiateWONos(Tagnos)
                If CheckWO = "NOK" Then
                    Txt_Msg.Text = "Invalid WO - WO is not registered in Server"
                    GoTo NoChange
                ElseIf CheckWO = "OPEN" Then

                ElseIf CheckWO = "CLOSED" Then

                ElseIf CheckWO = "FORCED" Then

                ElseIf CheckWO = "DISTRUP" Then

                End If
                'update the Current WO into the database before changing
                If CheckWOExist(lbl_WOnos.Text) Then
                    Call UpdateWO(lbl_WOnos.Text, lbl_unitaryCount.Text)
                Else
                    Call AddWO(lbl_WOnos.Text)
                    Call UpdateWO(lbl_WOnos.Text, lbl_unitaryCount.Text)
                End If

            End If
WOChange:
            Txt_Msg.Text = "Changing Series..." & vbCrLf
            Txt_Msg.Text = Txt_Msg.Text & "Reading info from RFID Tag..." & vbCrLf
            LoadWOfrRFID.JobNos = Tagnos
            Tagref = RD_MULTI_RFID("0014", 10) 'Read WO Reference from Tag
            If Tagref = "NOK" Then
                Txt_Msg.Text = Txt_Msg.Text & "--> Unable to read from RFID Tag" & vbCrLf
                Txt_Msg.Text = Txt_Msg.Text & "--> Change Series fail" & vbCrLf
                ReadTagFlag = False
                Exit Sub
            End If
            TagQty = RD_MULTI_RFID("0028", 10) 'Read WO Qty from Tag
            If TagQty = "NOK" Then
                Txt_Msg.Text = Txt_Msg.Text & "--> Unable to read from RFID Tag" & vbCrLf
                Txt_Msg.Text = Txt_Msg.Text & "--> Change Series fail" & vbCrLf
                ReadTagFlag = False
                Exit Sub
            End If
            Tagid = RD_MULTI_RFID("0040", 3) 'Read Tag ID
            If Tagid = "NOK" Then
                Txt_Msg.Text = Txt_Msg.Text & "--> Unable to read from RFID Tag" & vbCrLf
                Txt_Msg.Text = Txt_Msg.Text & "--> Change Series fail" & vbCrLf
                ReadTagFlag = False
                Exit Sub
            End If
            'Check if reference is valid from the database
            If Not RefCheck(Tagref) Then
                Txt_Msg.Text = Txt_Msg.Text & "--> Invalid Reference" & vbCrLf
                Txt_Msg.Text = Txt_Msg.Text & "--> Change Series fail" & vbCrLf
                ReadTagFlag = False
                Exit Sub
            End If
            Txt_Msg.Text = Txt_Msg.Text & "loading parameters of new reference..." & vbCrLf
            If Not LoadModelMaterial(Tagref) Then
                Txt_Msg.Text = Txt_Msg.Text & "--> Unable to load Model parameter" & vbCrLf
                Txt_Msg.Text = Txt_Msg.Text & "--> Change Series fail" & vbCrLf
                ReadTagFlag = False
                Exit Sub
            End If
            If Not LoadParameter(Tagref) Then
                Txt_Msg.Text = Txt_Msg.Text & "--> Unable to Load parameter from Server" & vbCrLf
                Txt_Msg.Text = Txt_Msg.Text & "--> Change Series fail" & vbCrLf
                ReadTagFlag = False
                Exit Sub
            End If

            Txt_Msg.Text = Txt_Msg.Text & "Configuring tester and station..." & vbCrLf
            System.Threading.Thread.Sleep(50)
            If Not UpdateParameter2Tester() Then
                Txt_Msg.Text = Txt_Msg.Text & "--> Unable to communicate with IO..."
                Txt_Msg.Text = Txt_Msg.Text & "--> Change Series fail" & vbCrLf
                ReadTagFlag = False
                Exit Sub
            End If

            If Not ActivateRackLED() Then
                Txt_Msg.Text = Txt_Msg.Text & "--> Unable to communicate with IO" & vbCrLf
                Txt_Msg.Text = Txt_Msg.Text & "--> Change Series fail" & vbCrLf
                ReadTagFlag = False
                Exit Sub
            End If
            lbl_WOnos.Text = Tagnos
            LoadWOfrRFID.JobNos = Tagnos
            lbl_currentref.Text = Tagref
            LoadWOfrRFID.JobModelName = Tagref
            lbl_wocounter.Text = TagQty
            LoadWOfrRFID.JobQTy = TagQty
            lbl_tagnos.Text = Tagid
            LoadWOfrRFID.JobRFIDTag = Tagid
            lbl_ArticleNos.Text = LoadWOfrRFID.JobArticle
            LabelVar = LoadLabelParameter(LoadWOfrRFID.JobModelName)
            'Image1.Picture = LoadPicture(INIPHOTOPATH & LabelVar.UnitImage)

            If Tagnos <> "MASTER" Then
                If CheckWOExist(Tagnos) Then
                    LoadWOfrRFID.JobUnitaryCount = Val(RetrieveWOQty(Tagnos))
                    lbl_unitaryCount.Text = LoadWOfrRFID.JobUnitaryCount
                Else
                    Call AddWO(Tagnos)
                    LoadWOfrRFID.JobUnitaryCount = 0 'Reset output counter
                    lbl_unitaryCount.Text = "0"
                End If
            Else
                lbl_unitaryCount.Text = "0"
                LoadWOfrRFID.JobUnitaryCount = 0
            End If


            LoadWOfrRFID.JobUnitaryCount = Val(lbl_unitaryCount.Text)
            UpdateStnStatus()

            Dim textBoxes As List(Of TextBox) = New List(Of TextBox) From {Txt_Originalstate1, Txt_Originalstate2, Txt_Originalstate3, Txt_Originalstate4, Txt_Originalstate5, Txt_Originalstate6, Txt_KeyState1, Txt_KeyState2, Txt_KeyState3, Txt_KeyState4, Txt_KeyState5, Txt_KeyState6, Txt_KeyTensionState1, Txt_KeyTensionState2, Txt_KeyTensionState3, Txt_KeyTensionState4, Txt_KeyTensionState5, Txt_KeyTensionState6}

            For i As Integer = 0 To 17
                textBoxes(i).Text = ""
            Next
            Txt_Msg.Text = Txt_Msg.Text & "Change Series completed" & vbCrLf
        End If
        ReadTagFlag = False


NoChange:
        TextBox1.Text = ReadKeyCyl()
        Interruptread = ReadIO(4, 0, 7)
        If Interruptread = 0 Then
            TestAction = 9999
            DisConnectContact()
            DisconnectHipot()
            DeEnergizeMagnet()
            TurnOffX1()
            TurnOffX2()

            frmMsg.Show
            frmMsg.Label1.Text = "EMERGENCY STOP"
            frmMsg.Button1.Visible = False
            TestAction = 9999
            InterruptFlag = True
        End If
        If Interruptread = 1 And InterruptFlag = True Then
            TestAction = 9999
            frmMsg.Label1.Text = "EMERGENCY STOP RELEASED"
            frmMsg.Button1.Visible = True
        End If
        Label6.Text = "System Information - " & TestAction

        Select Case TestAction
            Case 0
                'Read Door Close
                'If ScanPSnFlag = False Then GoTo DirectAssyAction
                If ReadIO(4, 0, 6) = 1 Then 'Door Sensor
                    TeststatusFlag = False
                    Txt_Msg.Text = ""
                    ClearDisplay()
                    'If ScanPSnFlag = True Then
                    TestAction = 20
                    'Else
                    '    lbl_msg.Text = "Please scan PSN before test..." & vbCrLf
                    '    ScanPSnFlag = False
                    'End If
                End If

            Case 10
                'ScanPSnFlag = False
                'Check Connector ID
                If ReadConnectorID <> LoadWOfrRFID.JobConnectorID Then
                    Txt_Msg.Text = Txt_Msg.Text & "Wrong Connector cable used..." & vbCrLf
                    TestAction = 9000
                Else
                    TestAction = 20
                End If

            Case 20
                'Lock Door
                LockDoor()
                TestAction = 30

            Case 30
                Txt_Msg.Text = ""
                ClearDisplay
                TeststatusFlag = False
                'Clamp Product
                HoldDUTCyldn()
                'DELAY (1)
                TestAction = 40

            Case 40
                'Check holddutcyl status
                '   If ReadIO(4, 0, 3) = 1 Then
                System.Threading.Thread.Sleep(50)
                TestAction = 50
   ' End If
            Case 50
                'connect product
                ConnectContact()
                System.Threading.Thread.Sleep(50)
                TestAction = 60
            Case 60
                'check Original state
                Dim ContactState As String
                ContactState = ReadContact()
                Txt_Originalstate1.Text = Mid(ContactState, 1, 1)
                Txt_Originalstate2.Text = Mid(ContactState, 2, 1)
                Txt_Originalstate3.Text = Mid(ContactState, 3, 1)
                Txt_Originalstate4.Text = Mid(ContactState, 4, 1)
                Txt_Originalstate5.Text = Mid(ContactState, 5, 1)
                Txt_Originalstate6.Text = Mid(ContactState, 6, 1)

                If ContactState <> DUTParameter.UnitOriginalSate Then
                    Txt_Msg.Text = Txt_Msg.Text & "--> Original State - FAIL"
                    TestAction = 9000
                Else
                    TestAction = 62
                End If

            Case 62
                DisConnectContact()
                TurnOffX1()
                TurnOffX2()
                DeEnergizeMagnet()
                System.Threading.Thread.Sleep(100)
                If ScanPSnFlag = True Then
                    If LoadWOfrRFID.JobProductMaterial = "Plastic" Then
                        TestAction = 100
                    Else
                        TestAction = 70
                    End If
                Else
                    TestAction = 100
                End If
            Case 70
                'connect product for di-electric
                ConnectHipot()
                System.Threading.Thread.Sleep(100)
                TestAction = 80

            Case 80
                'Di-electric test
                Call Set_chroma("SOUR:SAFE:STAR", "")
                System.Threading.Thread.Sleep(150)
                TestAction = 85
            Case 85
                'Check status of Di-Electric Test
                If Not Poll_Chroma_Result Then
                    TestAction = 9000
                    DisconnectHipot()
                Else
                    TestAction = 90
                End If
            Case 90
                'Disconnect Di-electric
                DisconnectHipot()
                TestAction = 100

            Case 100
                TestAction = 200


            Case 200
                'Connector E2 to product
                SetIO(0, 0, 7)  'E2
                TestAction = 210

            Case 210
                'connect X1 to product
                TurnOnX1()
                System.Threading.Thread.Sleep(100)
                TestAction = 220

            Case 220
                'Check X1 and X2 Sensor
                If ReadIO(4, 3, 1) <> 1 Or ReadIO(4, 3, 0) <> 0 Then
                    UnSetIO(0, 0, 7)  'E2
                    TurnOffX1()
                    Txt_Msg.Text = Txt_Msg.Text & "--> X1 LED Fail" & vbCrLf
                    TestAction = 9000
                Else
                    TestAction = 230
                End If

            Case 230
                'disconnect X1 and connect X2
                TurnOffX1()
                TurnOnX2()
                TurnOnX1()
                TestAction = 240

            Case 240
                'Check X1 and X2
                If ReadIO(4, 3, 1) <> 0 Or ReadIO(4, 3, 0) <> 1 Then
                    UnSetIO(0, 0, 7)  'E2
                    TurnOffX2()
                    Txt_Msg.Text = Txt_Msg.Text & "--> X2 LED Fail" & vbCrLf
                    TestAction = 9000
                Else
                    TestAction = 250
                End If
            Case 250
                'Disconnect X2
                TurnOffX2()
                TestAction = 260

            Case 260
                'SetIO 0, 1, 0 'S1_24V
                'SetIO 0, 1, 1 'S1
                'SetIO 0, 1, 3 'S2
                'SetIO 0, 1, 5 'S3
                'SetIO 0, 1, 7 'S4
                'SetIO 0, 0, 1 'S5
                'SetIO 0, 0, 3 'S6
                ConnectContact()
                TestAction = 270
                System.Threading.Thread.Sleep(50)

            Case 270
                Test.OriginState(1) = ReadIO(4, 1, (1 - 1))
                Txt_Originalstate1.Text = Test.OriginState(1)
                Test.OriginState(2) = ReadIO(4, 1, (2 - 1))
                Txt_Originalstate2.Text = Test.OriginState(2)
                Test.OriginState(3) = ReadIO(4, 1, (3 - 1))
                Txt_Originalstate3.Text = Test.OriginState(3)
                Test.OriginState(4) = ReadIO(4, 1, (4 - 1))
                Txt_Originalstate4.Text = Test.OriginState(4)
                Test.OriginState(5) = ReadIO(4, 1, (5 - 1))
                Txt_Originalstate5.Text = Test.OriginState(5)
                Test.OriginState(6) = ReadIO(4, 1, (6 - 1))
                Txt_Originalstate6.Text = Test.OriginState(6)

                TestAction = 271

            Case 271
                KeyCyldn()
                TestAction = 272

            Case 272
                'CheckKeyDn
                If ReadKeyCyl = 2 Then
                    TestAction = 273
                End If

            Case 273
                Test.KeyState(1) = ReadIO(4, 1, (1 - 1))
                Txt_KeyState1.Text = Test.KeyState(1)
                Test.KeyState(2) = ReadIO(4, 1, (2 - 1))
                Txt_KeyState2.Text = Test.KeyState(2)
                Test.KeyState(3) = ReadIO(4, 1, (3 - 1))
                Txt_KeyState3.Text = Test.KeyState(3)
                Test.KeyState(4) = ReadIO(4, 1, (4 - 1))
                Txt_KeyState4.Text = Test.KeyState(4)
                Test.KeyState(5) = ReadIO(4, 1, (5 - 1))
                Txt_KeyState5.Text = Test.KeyState(5)
                Test.KeyState(6) = ReadIO(4, 1, (6 - 1))
                Txt_KeyState6.Text = Test.KeyState(6)
                TestAction = 275

            Case 274
                If DUTParameter.UnitFunctionType = "Energisation" Then
                    TestAction = 275
                Else
                    TestAction = 275
                End If

            Case 275
                EnergizeMagnet()
                TestAction = 276

            Case 276
                Test.KeyTension(1) = ReadIO(4, 1, (1 - 1))
                Txt_KeyTensionState1.Text = Test.KeyTension(1)
                Test.KeyTension(2) = ReadIO(4, 1, (2 - 1))
                Txt_KeyTensionState2.Text = Test.KeyTension(2)
                Test.KeyTension(3) = ReadIO(4, 1, (3 - 1))
                Txt_KeyTensionState3.Text = Test.KeyTension(3)
                Test.KeyTension(4) = ReadIO(4, 1, (4 - 1))
                Txt_KeyTensionState4.Text = Test.KeyTension(4)
                Test.KeyTension(5) = ReadIO(4, 1, (5 - 1))
                Txt_KeyTensionState5.Text = Test.KeyTension(5)
                Test.KeyTension(6) = ReadIO(4, 1, (6 - 1))
                Txt_KeyTensionState6.Text = Test.KeyTension(6)
                TestAction = 277

            Case 277
                If DUTParameter.UnitFunctionType <> "Energisation" Then
                    DeEnergizeMagnet()
                    System.Threading.Thread.Sleep(100)
                End If
                TestAction = 278

            Case 278
                KeyCylup()
                TimeoutCount = 0
                TestAction = 279

            Case 279
                If ReadKeyCyl = 1 Then
                    TestAction = 280
                Else
                    TimeoutCount = TimeoutCount + 1
                    If TimeoutCount > 20 Then

                        Txt_Msg.Text = Txt_Msg.Text & "Unable to remove key from product"
                        TestAction = 9000
                    End If
                End If
    'Checkcylup

            Case 280
                DeEnergizeMagnet()
                TestAction = 281

            Case 281
                'checkcylup
                DisConnectContact()
                TestAction = 290

            Case 282

                If ReadIO(4, 0, 4) = 1 Then
                    HoldDUTCylup()
                    TestAction = 0
                End If

            Case 290
                If TestRead.OriginState(1) <> DUTParameter.UnitContactOriginaltemp(1) Or TestRead.KeyState(1) <> DUTParameter.UnitcontactWkeytemp(1) Or TestRead.KeyTension(1) <> DUTParameter.UnitContactWkeyTensiontemp(1) Then
                    Txt_ContactPN1.BackColor = Color.Red
                    Txt_ContactPN2.BackColor = Color.Red
                    TeststatusFlag = True
                End If
                If TestRead.OriginState(2) <> DUTParameter.UnitContactOriginaltemp(2) Or TestRead.KeyState(2) <> DUTParameter.UnitcontactWkeytemp(2) Or TestRead.KeyTension(2) <> DUTParameter.UnitContactWkeyTensiontemp(2) Then
                    Txt_ContactPN3.BackColor = Color.Red
                    Txt_ContactPN4.BackColor = Color.Red
                    TeststatusFlag = True
                End If
                If TestRead.OriginState(3) <> DUTParameter.UnitContactOriginaltemp(3) Or TestRead.KeyState(3) <> DUTParameter.UnitcontactWkeytemp(3) Or TestRead.KeyTension(3) <> DUTParameter.UnitContactWkeyTensiontemp(3) Then
                    Txt_ContactPN5.BackColor = Color.Red
                    Txt_ContactPN6.BackColor = Color.Red
                    TeststatusFlag = True
                End If
                If TestRead.OriginState(4) <> DUTParameter.UnitContactOriginaltemp(4) Or TestRead.KeyState(4) <> DUTParameter.UnitcontactWkeytemp(4) Or TestRead.KeyTension(4) <> DUTParameter.UnitContactWkeyTensiontemp(4) Then
                    Txt_ContactPN7.BackColor = Color.Red
                    Txt_ContactPN8.BackColor = Color.Red
                    TeststatusFlag = True
                End If
                If TestRead.OriginState(5) <> DUTParameter.UnitContactOriginaltemp(5) Or TestRead.KeyState(5) <> DUTParameter.UnitcontactWkeytemp(5) Or TestRead.KeyTension(5) <> DUTParameter.UnitContactWkeyTensiontemp(5) Then
                    Txt_ContactPN9.BackColor = Color.Red
                    Txt_ContactPN10.BackColor = Color.Red
                    TeststatusFlag = True
                End If
                If TestRead.OriginState(6) <> DUTParameter.UnitContactOriginaltemp(6) Or TestRead.KeyState(6) <> DUTParameter.UnitcontactWkeytemp(6) Or TestRead.KeyTension(6) <> DUTParameter.UnitContactWkeyTensiontemp(6) Then
                    Txt_ContactPN11.BackColor = Color.Red
                    Txt_ContactPN12.BackColor = Color.Red
                    TeststatusFlag = True
                End If

                If Txt_ContactPN1.BackColor = Color.Red And Txt_ContactPN3.BackColor = Color.Red And Txt_ContactPN5.BackColor = Color.Red And Txt_ContactPN7.BackColor = Color.Red And Txt_ContactPN9.BackColor = Color.Red And Txt_ContactPN11.BackColor = Color.Red Then
                    Txt_ContactPN15.BackColor = Color.Red
                End If

                If TeststatusFlag = True Then
                    TestAction = 9000
                Else
                    TestAction = 8000
                End If

            Case 8000
                DisConnectContact()
                DeEnergizeMagnet()
                TurnOffX1()
                TurnOffX2()
                DisconnectHipot()
                If ScanPSnFlag = True Then
                    If LabelVar.UnitLabelTemplate <> "" Then 'If no define template, skip print unitary
                        'If Not OpenLab(INITEMPLATEPATH & LabelVar.UnitLabelTemplate) Then
                        '    Txt_Msg.Text = "--> Unable to open label template" & vbCrLf
                        '    Image2.Picture = LoadPicture(App.Path & "\Icons\FAIL.Emf")
                        '    TestAction = 9030
                        '    GoTo DirectAssyAction
                        'End If
                        'If Not PrinterSelect("Zebra 110Xi4 (600 dpi)", "USB001") Then
                        'If Not PrinterSelect("Zebra 110XiIII Plus (600 dpi)", "LPT1:") Then
                        '    Txt_Msg.Text = "--> Unable to select Printer" & vbCrLf
                        '    Image2.Picture = LoadPicture(App.Path & "\Icons\FAIL.Emf")
                        '    TestAction = 9030
                        '    GoTo DirectAssyAction
                        'End If
                        If PSNFileInfo.ConnTestStatus <> "PASS" Then
                            'Txt_Msg.Text = "Printing label..."
                            'PrintLabel
                            lbl_unitaryCount.Text = Val(lbl_unitaryCount.Text) + 1
                            LoadWOfrRFID.JobUnitaryCount = Val(lbl_unitaryCount.Text)
                            UpdateStnStatus()
                            'lbl_msg.Text = "Please paste label on product"
                        Else
                            Txt_Msg.Text = "Repeat label Print - No Printing"
                        End If
                        'CloseLab
                    End If

                End If
                TestAction = 8010

            Case 8010
                If ScanPSnFlag = True Then
                    PSNFileInfo.ConnTestStatus = "PASS"
                    PSNFileInfo.ConnTestCheckOut = currentDate.Date & "," & Now.ToLongTimeString
                    If Not WRITEPSNFILE(PSNFileInfo.PSN) Then
                        Txt_Msg.Text = "--> Unable to write " & PSNFileInfo.PSN & ".Txt in the server" & vbCrLf
                        Image2.Image = My.Resources.ResourceManager.GetObject("Fail")
                    End If
                End If
                Image2.Image = My.Resources.ResourceManager.GetObject("Pass")
                UnlockDoor()
                PushOpenDoor()
                TestAction = 8020

            Case 8020
                If ReadIO(4, 0, 4) = 1 Then
                    HoldDUTCylup()
                    TestAction = 0
                    lbl_msg.Text = "Please scan the PSN barcode..."
                    ScanPSnFlag = False
                End If


            Case 9000 'FAIL SEQ

                'DisConnectContact
                'DeEnergizeMagnet
                'TurnOffX1
                'TurnOffX2
                'DisconnectHipot
                If DUTParameter.UnitFunctionType = "Energisation" Then
                    EnergizeMagnet()
                Else
                    DeEnergizeMagnet()
                End If
                KeyCylup()
                TestAction = 9010
                TimeoutCount = 0

            Case 9010
                If ReadKeyCyl = 1 Then
                    TestAction = 9020
                Else
                    TimeoutCount = TimeoutCount + 1
                    If TimeoutCount > 60 Then
                        KeyCylFree()
                        TestAction = 9020
                    End If
                End If

            Case 9020
                DisConnectContact()
                DeEnergizeMagnet()
                TurnOffX1()
                TurnOffX2()
                DisconnectHipot()
                HoldDUTCylup()
                TestAction = 9025

            Case 9025
                If ScanPSnFlag = True Then
                    PSNFileInfo.ConnTestStatus = "FAIL"
                    PSNFileInfo.ConnTestCheckOut = currentDate.Date & "," & Now.ToLongTimeString
                    Txt_Msg.Text = Txt_Msg.Text & PSNFileInfo.PSN & vbCrLf
                    If Not WRITEPSNFILE(PSNFileInfo.PSN) Then
                        Txt_Msg.Text = "--> Unable to update " & PSNFileInfo.PSN & ".Txt in the server" & vbCrLf
                        Image2.Image = My.Resources.ResourceManager.GetObject("Fail")
                    End If
                End If
                Image2.Image = My.Resources.ResourceManager.GetObject("Fail")
                ScanPSnFlag = False
                TestAction = 9030

            Case 9030
                UnlockDoor()
                PushOpenDoor()
                lbl_msg.Text = "Please scan the PSN barcode..."
                TestAction = 0

        End Select

DirectAssyAction:

        Select Case AssyAction
            Case 0

            Case 1


        End Select
    End Sub
    Private Function RefCheck(strName As String) As Boolean
        Dim temp As String
        On Error GoTo ErrorHandler

        Dim query = "Select * From Label Where ModelName = '" & strName & "'"
        Dim dt = koneksiDB.bacaData(query).Tables(0)

        temp = dt.Rows(0).Item("ModelName")
        If temp <> "" Then
            Return True
        End If
ErrorHandler:
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PushOpenDoor()
    End Sub
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.CheckState = 0 Then
            KeyCyldn()
        Else
            KeyCylup()
        End If
    End Sub
    Private Sub ClearDisplay()
        Dim textBoxes As List(Of TextBox) = New List(Of TextBox) From {Txt_Originalstate1, Txt_Originalstate2, Txt_Originalstate3, Txt_Originalstate4, Txt_Originalstate5, Txt_Originalstate6, Txt_KeyState1, Txt_KeyState2, Txt_KeyState3, Txt_KeyState4, Txt_KeyState5, Txt_KeyState6, Txt_KeyTensionState1, Txt_KeyTensionState2, Txt_KeyTensionState3, Txt_KeyTensionState4, Txt_KeyTensionState5, Txt_KeyTensionState6, Txt_ContactPN1, Txt_ContactPN2, Txt_ContactPN3, Txt_ContactPN4, Txt_ContactPN5, Txt_ContactPN6, Txt_ContactPN7, Txt_ContactPN8, Txt_ContactPN9, Txt_ContactPN10, Txt_ContactPN11, Txt_ContactPN12, Txt_ContactPN13, Txt_ContactPN14, Txt_ContactPN15, Txt_ContactPN16}

        For i As Integer = 0 To 33
            If i < 18 Then
                textBoxes(i).Text = ""
            End If
            textBoxes(i).BackColor = Color.White
        Next
        'Shape2.FillColor = vbBlack
        'Shape3.FillColor = vbBlack
        Image2.Image = Nothing
    End Sub
    Public Function Poll_Chroma_Result() As Boolean
Retry:
        Chroma_Comm.DiscardInBuffer()
        ChromaFBbuf = ""
        Thread.Sleep(20)
        Chroma_Comm.Write("SOUR:SAFE:RES:ALL?" & vbCrLf)
        Do While ChromaFBbuf = ""
            System.Windows.Forms.Application.DoEvents()
        Loop
        If ChromaFBbuf = "116" Then
            Return True
            Exit Function
        ElseIf ChromaFBbuf = "17" Then
            Return False
            Exit Function
        ElseIf ChromaFBbuf = "18" Then
            Return False
            Exit Function
        ElseIf ChromaFBbuf = "115" Then
            GoTo Retry
        ElseIf ChromaFBbuf = "114" Then
            Return False
            Exit Function
        ElseIf ChromaFBbuf = "113" Then
            Return False
            Exit Function
        ElseIf ChromaFBbuf = "112" Then
            Return False
            Exit Function
        End If
    End Function
End Class
