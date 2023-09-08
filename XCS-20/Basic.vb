Module Basic
    Dim UserHandle As Long
    Dim Attrib As SECURITY_ATTRIBUTES
    Dim giNumberOfCard As Integer
    Dim UserEvent As Long
    Public Function INIT_PCI112() As Boolean
        UserHandle = CreateEvent(Attrib, True, False, "UserEvent")

        If B_pll12_initial(giNumberOfCard) = 0 Then
            'MsgBox "PCIL112 Found!!"
        End If
        If giNumberOfCard = 0 Then
            MsgBox("No PCIL112 Found!!")
            End
        End If
        B_mnet_reset_ring(0)
        B_mnet_start_ring(0)
        System.Threading.Thread.Sleep(200)
        B_mnet_enable_soft_watchdog(0, UserEvent)
        B_mnet_set_ring_quality_param(0, 100, 800)
    End Function
    Public Function UnSetIO(ByVal ModNo As Integer, ByVal PortNo As Integer, ByVal BitNo As Integer)
        Dim input1 As Long
        Dim BinBit As Long
        Dim setval2 As Long

        BinBit = 2 ^ BitNo
        input1 = B_mnet_io_input(0, ModNo, PortNo)
        setval2 = input1 And Not (BinBit)

        B_mnet_io_output(0, ModNo, PortNo, setval2)

    End Function
    Public Function SetIO(ModNo As Integer, PortNo As Integer, BitNo As Integer)
        Dim input1 As Integer
        Dim BinBit As Integer
        Dim setval As Integer

        BinBit = 2 ^ BitNo
        input1 = B_mnet_io_input(0, ModNo, PortNo)
        setval = BinBit Or input1

        B_mnet_io_output(0, ModNo, PortNo, setval)
    End Function
    Public Sub HoldDUTCylup()
        'SetIO 2, 3, 2
        UnSetIO(0, 2, 2)
    End Sub

    Public Sub HoldDUTCyldn()
        SetIO(0, 2, 2)
        'UnSetIO 2, 3, 2
    End Sub
    Public Sub KeyCylup()
        SetIO(0, 2, 1)
        UnSetIO(0, 2, 0)
    End Sub
    Public Sub KeyCyldn()
        SetIO(0, 2, 0)
        UnSetIO(0, 2, 1)
    End Sub
    Public Sub PushOpenDoor()
        SetIO(0, 2, 4)
        System.Threading.Thread.Sleep(300)
        UnSetIO(0, 2, 4)
    End Sub
    Public Sub ConnectContact()
        SetIO(0, 1, 0) 'S1_24V
        SetIO(0, 1, 1) 'S1
        SetIO(0, 1, 2) 'S2_24V
        SetIO(0, 1, 3) 'S2
        SetIO(0, 1, 4) 'S3_24V
        SetIO(0, 1, 5) 'S3
        SetIO(0, 1, 6) 'S4_24V
        SetIO(0, 1, 7) 'S4
        SetIO(0, 0, 0) 'S5_24V
        SetIO(0, 0, 1) 'S5
        SetIO(0, 0, 2) 'S6_24V
        SetIO(0, 0, 3) 'S6
    End Sub
    Public Sub DisConnectContact()
        UnSetIO(0, 1, 0) 'S1_24V
        UnSetIO(0, 1, 1) 'S1
        UnSetIO(0, 1, 2) 'S2_24V
        UnSetIO(0, 1, 3) 'S2
        UnSetIO(0, 1, 4) 'S3_24V
        UnSetIO(0, 1, 5) 'S3
        UnSetIO(0, 1, 6) 'S4_24V
        UnSetIO(0, 1, 7) 'S4
        UnSetIO(0, 0, 0) 'S5_24V
        UnSetIO(0, 0, 1) 'S5
        UnSetIO(0, 0, 2) 'S6_24V
        UnSetIO(0, 0, 3) 'S6
    End Sub
    Public Function ReadKeyCyl() As Double
        Dim Reed(2) As String
        Reed(1) = ReadIO(4, 0, 0)
        Reed(2) = ReadIO(4, 0, 1)

        ReadKeyCyl = Bin162Dec("00000000000000" & Reed(1) & Reed(2))
    End Function
    Public Function ReadIO(SlaveNos As Integer, PortNo As Integer, BitNo As Integer) As Integer
        Dim input1 As Integer

        input1 = B_mnet_io_input(0, SlaveNos, PortNo)

        If input1 And 2 ^ BitNo Then
            Return 1
        Else
            Return 0
        End If
    End Function
    Public Sub ConnectHipot()
        SetIO(0, 3, 0)
        SetIO(0, 3, 1)
    End Sub
    Public Sub DisconnectHipot()
        UnSetIO(0, 3, 0)
        UnSetIO(0, 3, 1)
    End Sub
    Public Sub DeEnergizeMagnet()
        UnSetIO(0, 0, 6) 'E1
        UnSetIO(0, 0, 7) 'E2
    End Sub
    Public Sub TurnOnX1()
        SetIO(0, 0, 4)
    End Sub

    Public Sub TurnOffX1()
        UnSetIO(0, 0, 4)
    End Sub

    Public Sub TurnOnX2()
        SetIO(0, 0, 5)
    End Sub

    Public Sub TurnOffX2()
        UnSetIO(0, 0, 5)
    End Sub
    Public Sub KeyCylFree()
        UnSetIO(0, 2, 1)
        UnSetIO(0, 2, 0)
    End Sub
    Public Sub UnlockDoor()
        'SetIO 2, 3, 4
        UnSetIO(0, 2, 5)
    End Sub
    Public Function ReadConnectorID() As String
        Dim ID(3) As String
        ID(1) = ReadIO(4, 3, 5)
        ID(2) = ReadIO(4, 3, 6)
        ID(3) = ReadIO(4, 3, 7)

        ReadConnectorID = BIN2Dec(ID(3) & ID(2) & ID(1))
    End Function
    Public Sub LockDoor()
        SetIO(0, 2, 5)
        'UnSetIO 2, 3, 4
    End Sub
    Public Function ReadContact() As String
        Dim Bit(6) As String

        Bit(1) = ReadIO(4, 1, 0)
        Bit(2) = ReadIO(4, 1, 1)
        Bit(3) = ReadIO(4, 1, 2)
        Bit(4) = ReadIO(4, 1, 3)
        Bit(5) = ReadIO(4, 1, 4)
        Bit(6) = ReadIO(4, 1, 5)

        ReadContact = Bit(1) & Bit(2) & Bit(3) & Bit(4) & Bit(5) & Bit(6)
    End Function
    Public Sub EnergizeMagnet()
        SetIO(0, 0, 6) 'E1
        SetIO(0, 0, 7) 'E2
    End Sub

End Module
