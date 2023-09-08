Option Explicit On
Imports System.Threading
Module RFID
	Public Function ReadTagChkUpStream() As Boolean
		Dim RdInfo As String

		On Error GoTo Comm_Error

		RdInfo = RD_MULTI_RFID("0000", 5)
		If RdInfo = "NOK" Then GoTo Comm_Error
		'Check Station status
		If Mid(RdInfo, 3, 1) = "P" Then
			Return True
		Else
			Return False
		End If
		Exit Function

Comm_Error:
		Return False
	End Function

	Public Function ReadTagCounter() As String
		Dim RdInfo As String

		On Error GoTo Comm_Error

		RdInfo = RD_MULTI_RFID("007B", 3)
		If RdInfo = "NOK" Then GoTo Comm_Error
		Return RdInfo
		Exit Function

Comm_Error:

	End Function

	Public Function ReadTagRef() As Boolean
		Dim RdInfo As String

		On Error GoTo Comm_Error

		RdInfo = RD_MULTI_RFID("0000", 10)
		If RdInfo = "NOK" Then GoTo Comm_Error
		Parameter.UnitModel = RdInfo
		Return True
		Exit Function

Comm_Error:
		Return False
	End Function

	Public Function ReadTagSerial() As Boolean
		Dim RdInfo As String

		On Error GoTo Comm_Error

		RdInfo = RD_MULTI_RFID("0014", 10)
		If RdInfo = "NOK" Then GoTo Comm_Error
		Parameter.UnitTagNos = RdInfo
		Return True
		Exit Function

Comm_Error:
		Return False
	End Function

	Public Function WrTagCounter(CountVal As String) As Boolean
		Dim i, j As Integer
		Dim WrVal As String = ""
		Dim Schar As String
		Dim Ext As String = ""
		Dim Strlen As Integer

		Strlen = Len(CountVal)

		CountVal = Trim(CountVal)
		For i = 1 To Strlen
			Schar = Mid(CountVal, i, 1)
			WrVal = WrVal & Hex(Asc(Schar))
		Next
		For j = 1 To (6 - Strlen)
			Ext = Ext & "00"
		Next
		WrVal = Ext & WrVal
		If Not WR_RFID("007B", Mid(WrVal, 1, 4)) Then
			Return False
		End If
		If Not WR_RFID("007C", Mid(WrVal, 5, 4)) Then
			Return False
		End If
		If Not WR_RFID("007D", Mid(WrVal, 9, 4)) Then
			Return False
		End If
		'If Not WR_RFID("0056", Mid(WrVal, 13, 4)) Then
		'    Return False
		'End If
		Return True
	End Function

	Public Function CRC(DATABUF As String()) As String
		Dim CRCSetting As String
		Dim j, i As Integer
		Dim b() As String

		b = DATABUF
		CRCSetting = "&HFFFF"
		For i = 1 To 6
			CRCSetting = CRCSetting Xor b(i - 1)
			For j = 1 To 8
				If CRCSetting Mod 2 = 0 Then
					CRCSetting = "&H" & Hex(CRCSetting \ 2)
				Else
					CRCSetting = ("&H" & Hex(CRCSetting \ 2)) Xor "&HA001"
					CRCSetting = "&H" & Hex(CRCSetting)
				End If
			Next
		Next
		CRC = CRCSetting
	End Function

	Public Function BIN2Dec(Bin As String) As Double
		Dim i As Integer
		Dim TempDec As Double
		Dim Extract As String
		Dim n As Integer
		Dim BinLen As Integer

		BinLen = Len(Bin)
		For i = BinLen To 1 Step -1
			Extract = Mid(Bin, i, 1)
			If Extract = "1" Then
				TempDec = TempDec + 2 ^ n
			End If
			n = n + 1
		Next
		Return TempDec
	End Function

	Public Function BIN2HEX(Bin As String) As String
		Dim i As Integer
		Dim Data(4) As String
		Dim TempBinary(4) As String

		TempBinary(1) = Mid(Bin, 1, 4)
		TempBinary(2) = Mid(Bin, 5, 4)
		TempBinary(3) = Mid(Bin, 9, 4)
		TempBinary(4) = Mid(Bin, 13, 4)
		For i = 1 To 4
			Select Case TempBinary(i)
				Case "0000"
					Data(i) = "0"
				Case "0001"
					Data(i) = "1"
				Case "0010"
					Data(i) = "2"
				Case "0011"
					Data(i) = "3"
				Case "0100"
					Data(i) = "4"
				Case "0101"
					Data(i) = "5"
				Case "0110"
					Data(i) = "6"
				Case "0111"
					Data(i) = "7"
				Case "1000"
					Data(i) = "8"
				Case "1001"
					Data(i) = "9"
				Case "1010"
					Data(i) = "A"
				Case "1011"
					Data(i) = "B"
				Case "1100"
					Data(i) = "C"
				Case "1101"
					Data(i) = "D"
				Case "1110"
					Data(i) = "E"
				Case "1111"
					Data(i) = "F"
			End Select
		Next
		Return Data(1) & Data(2) & Data(3) & Data(4)
	End Function

	Public Function Dec2Bin(Data As Double) As String
		Dim i As Integer
		Dim bit16(16) As String
		Dim Mw As String

		For i = 15 To 0 Step -1
			If Data >= 2 ^ i Then
				bit16(i) = "1"
				Data = Data - 2 ^ i
			Else
				bit16(i) = "0"
			End If
		Next
		For i = 15 To 0 Step -1
			Mw = Mw & bit16(i)
		Next
		Return Mw
	End Function

	Public Function Hex2Dec(ByRef Data As String) As String
		Dim i As Integer
		Dim Dec As Integer
		Dim CalDec As Integer
		Dim Char_Renamed As String
		Dim n As Integer
		Dim LenData As Integer
		LenData = Len(Data)
		For i = (LenData - 1) To 0 Step -1
			Char_Renamed = Mid(Data, i + 1, 1)
			Select Case Char_Renamed
				Case "F"
					Dec = 15
				Case "E"
					Dec = 14
				Case "D"
					Dec = 13
				Case "C"
					Dec = 12
				Case "B"
					Dec = 11
				Case "A"
					Dec = 10
				Case Else
					Dec = CInt(Char_Renamed)
			End Select
			CalDec = CalDec + Dec * (16 ^ n)
			n = n + 1
		Next
		Return Str(CalDec)
	End Function

	Public Function Hex2Bin(Hexdata As String) As String
		Dim i As Integer
		Dim DataLen As Integer
		Dim TempHex(4) As String
		Dim TempBin(4) As String

		DataLen = Len(Hexdata)
		If DataLen = 1 Then
			TempHex(1) = Hexdata
		ElseIf DataLen = 2 Then
			TempHex(1) = Mid(Hexdata, 2, 1)
			TempHex(2) = Mid(Hexdata, 1, 1)
		ElseIf DataLen = 3 Then
			TempHex(1) = Mid(Hexdata, 3, 1)
			TempHex(2) = Mid(Hexdata, 2, 1)
			TempHex(3) = Mid(Hexdata, 1, 1)
		ElseIf DataLen = 4 Then
			TempHex(1) = Mid(Hexdata, 4, 1)
			TempHex(2) = Mid(Hexdata, 3, 1)
			TempHex(3) = Mid(Hexdata, 2, 1)
			TempHex(4) = Mid(Hexdata, 1, 1)
		End If

		TempBin(1) = "0000"
		TempBin(2) = "0000"
		TempBin(3) = "0000"
		TempBin(4) = "0000"

		For i = 1 To DataLen
			Select Case TempHex(i)
				Case "0"
					TempBin(i) = "0000"
				Case "1"
					TempBin(i) = "0001"
				Case "2"
					TempBin(i) = "0010"
				Case "3"
					TempBin(i) = "0011"
				Case "4"
					TempBin(i) = "0100"
				Case "5"
					TempBin(i) = "0101"
				Case "6"
					TempBin(i) = "0110"
				Case "7"
					TempBin(i) = "0111"
				Case "8"
					TempBin(i) = "1000"
				Case "9"
					TempBin(i) = "1001"
				Case "A"
					TempBin(i) = "1010"
				Case "B"
					TempBin(i) = "1011"
				Case "C"
					TempBin(i) = "1100"
				Case "D"
					TempBin(i) = "1101"
				Case "E"
					TempBin(i) = "1110"
				Case "F"
					TempBin(i) = "1111"
			End Select
		Next
		Return TempBin(4) & TempBin(3) & TempBin(2) & TempBin(1)
	End Function

	Public Function RD_RFID(AddrNos As String) As String
		Dim temp1, temp2 As String
		Dim n As Integer
		Dim ii() As String
		Dim SendString(8) As Byte
		Dim ReadBuf As String
		Dim ValueH, ValueL As String
		Dim CheckSum As String
		Dim Str1, Str2 As String
		Dim Bcount As Integer
		Dim t1, t2, t3, t4, t5, t6 As Integer
		Dim ErrCode As String
		Dim Retry As Integer

		On Error GoTo READ_RFID_TIMEOUT
READRETRY:
		ii = {"&HF8", "&H03", "&H", "&H", "&H00", "&H01", "&H", "&H"}
		ii(2) = ii(2) + Left(AddrNos, 2)
		ii(3) = ii(3) + Right(AddrNos, 2)
		CheckSum = CRC(ii) 'Calculate the Checksum
		Str1 = Right(CheckSum, Len(CheckSum) - 2)
		Select Case Len(Str1)
			Case 1
				ii(7) = ii(7) + "00"
				ii(6) = ii(6) + "0" + Str1
			Case 2
				ii(7) = ii(7) + "00"
				ii(6) = ii(6) + Str1
			Case 3
				ii(6) = ii(6) + Right(Str1, 2)
				ii(7) = ii(7) + "0" + Left(Str1, Len(Str1) - 2)
			Case 4
				ii(6) = ii(6) + Right(Str1, 2)
				ii(7) = ii(7) + Left(Str1, Len(Str1) - 2)
		End Select

		For i As Integer = 0 To 7
			SendString(i) = CByte(ii(i))
		Next

		frmMain.RFID_Comm.Write(SendString, 0, SendString.Length)
		Thread.Sleep(20)
		GoTo READINBYTE

READINBYTE:
		Str1 = ""
		If frmMain.RFID_Comm.BytesToRead = 0 Then GoTo READ_RFID_TIMEOUT
		Thread.Sleep(10)
		Bcount = frmMain.RFID_Comm.BytesToRead
		Do
			Str2 = frmMain.RFID_Comm.ReadExisting()
			n = n + 1
			Str2 = Asc(Str2)
			Str1 = Str1 + Trim(Str(Str2)) + "," 'Build the string received frm OsiTrack
		Loop Until n = Bcount
		frmMain.TextBox3.Text = Str1 'Display the string received
		t1 = InStr(1, Str1, ",")
		t2 = InStr(t1 + 1, Str1, ",")
		t3 = InStr(t2 + 1, Str1, ",")
		t4 = InStr(t3 + 1, Str1, ",")
		t5 = InStr(t4 + 1, Str1, ",")
		temp1 = Mid(Str1, t1 + 1, t2 - t1 - 1)
		temp2 = Mid(Str1, t1 + 1, t2 - t1 - 1)
		If Mid(Str1, t1 + 1, t2 - t1 - 1) = "131" Or Mid(Str1, t1 + 1, t2 - t1 - 1) = "132" Then '&H83 or &H84
			ErrCode = Mid(Str1, t2 + 1, t3 - t2 - 1)
			GoTo Comm_Error
		End If
		ValueH = Mid(Str1, t3 + 1, t4 - t3 - 1)
		ValueH = Hex(Val(ValueH))
		ValueL = Mid(Str1, t4 + 1, t5 - t4 - 1)
		ValueL = Hex(Val(ValueL))
		If Len(ValueL) = 1 Then
			ValueL = "0" + ValueL
		End If
		ReadBuf = ValueH + ValueL 'Read the LSB only(1 Byte)
		ReadBuf = Hex2Bin(ReadBuf)
		ReadBuf = Str(BIN2Dec(ReadBuf))
		Return Trim(ReadBuf)

		Exit Function

Comm_Error:
		Return "NOK"
		Exit Function

READ_RFID_TIMEOUT:
		If Retry < 3 Then
			Retry = Retry + 1
			Thread.Sleep(10)
			GoTo READRETRY
		End If
		Return "NOK"
		Exit Function
	End Function

	Public Function Wr_Tag(Data As String, StartAddr As String) As Boolean
		Dim i As Integer
		Dim Stringlen As Integer
		Dim n As Integer
		Dim CutdataH, CutdataL As String
		Dim WrdataH, WrdataL As String
		Dim WrAddr As String
		n = Len(Data) Mod 2
		Stringlen = (Len(Data) / 2)
		Stringlen = Stringlen + n
		For i = 0 To Stringlen - 1
			CutdataH = Mid(Data, 2 * i + 1, 1)
			CutdataL = Mid(Data, 2 * i + 2, 1)
			If CutdataH = "" Then
				WrdataH = "00"
			Else
				WrdataH = Hex(Asc(CutdataH))
			End If
			If CutdataL = "" Then
				WrdataL = "00"
			Else
				WrdataL = Hex(Asc(CutdataL))
			End If
			WrAddr = CStr(CDbl(Hex2Dec(StartAddr)) + 1 * i)
			WrAddr = Dec2Bin(CDbl(WrAddr))
			WrAddr = BIN2HEX(WrAddr)
			If Not WR_RFID(WrAddr, WrdataH & WrdataL) Then
				Return False
				Exit Function
			End If
		Next
		Return True
	End Function

	Public Function Clear_Tag(StartAddr As String, DataLen As Integer) As Boolean
		Dim i As Integer
		Dim WrdataH, WrdataL As String
		Dim WrAddr As String

		For i = 0 To DataLen - 1
			WrdataH = "00"
			WrdataL = "00"
			WrAddr = CStr(CDbl(Hex2Dec(StartAddr)) + 1 * i)
			WrAddr = Dec2Bin(CDbl(WrAddr))
			WrAddr = BIN2HEX(WrAddr)
			If Not WR_RFID(WrAddr, WrdataH & WrdataL) Then
				Return False
				Exit Function
			End If
		Next
		Return True
	End Function

	Public Function WR_RFID(AddrNos As String, SendData As String) As Boolean
		Dim n As Integer
		Dim ii() As String
		Dim SendString(8) As Byte
		Dim ValueH, ValueL As String
		Dim CheckSum As String
		Dim Str1, Str2 As String
		Dim Bcount As Integer
		Dim t1, t2, t3, t4, t5, t6 As Integer
		Dim ErrCode As String
		Dim Retry As Integer

		On Error GoTo WRITE_COMM_TIMEOUT
WRITERETRY:
		ii = {"&HF8", "&H06", "&H", "&H", "&H", "&H", "&H", "&H"}
		ii(2) = ii(2) + Left(AddrNos, 2)
		ii(3) = ii(3) + Right(AddrNos, 2)
		ii(4) = ii(4) + Left(SendData, 2)
		ii(5) = ii(5) + Right(SendData, 2)

		CheckSum = CRC(ii) 'Calculate Checksum
		Str1 = Right(CheckSum, Len(CheckSum) - 2)
		Select Case Len(Str1)
			Case 1
				ii(7) = ii(7) + "00"
				ii(6) = ii(6) + "0" + Str1
			Case 2
				ii(7) = ii(7) + "00"
				ii(6) = ii(6) + Str1
			Case 3
				ii(6) = ii(6) + Right(Str1, 2)
				ii(7) = ii(7) + "0" + Left(Str1, Len(Str1) - 2)
			Case 4
				ii(6) = ii(6) + Right(Str1, 2)
				ii(7) = ii(7) + Left(Str1, Len(Str1) - 2)
		End Select

		For i As Integer = 0 To 7
			SendString(i) = CByte(ii(i))
		Next

		frmMain.RFID_Comm.Write(SendString, 0, SendString.Length)
		Thread.Sleep(20)
		GoTo CHECKINBYTE

CHECKINBYTE:
		If frmMain.RFID_Comm.BytesToRead = 0 Then GoTo WRITE_COMM_TIMEOUT
		Thread.Sleep(10)
		Bcount = frmMain.RFID_Comm.BytesToRead
		Str1 = ""
		Do
			Str2 = frmMain.RFID_Comm.ReadExisting()
			n = n + 1
			Str2 = Asc(Str2)
			Str1 = Str1 + Trim(Str(Str2)) + ","
		Loop Until n = Bcount
		t1 = InStr(1, Str1, ",")
		t2 = InStr(t1 + 1, Str1, ",")
		t3 = InStr(t2 + 1, Str1, ",")
		t4 = InStr(t3 + 1, Str1, ",")
		t5 = InStr(t4 + 1, Str1, ",")
		t6 = InStr(t5 + 1, Str1, ",")
		If Mid(Str1, t1 + 1, t2 - t1 - 1) = "134" Then '&H83 or &H84 or &H86
			ErrCode = Mid(Str1, t2 + 1, t3 - t2 - 1)
			GoTo Comm_Error
		End If
		ValueH = Hex(CInt(Mid(Str1, t4 + 1, t5 - t4 - 1)))
		If Len(ValueH) = 1 Then ValueH = "0" & ValueH
		ValueL = Hex(CInt(Mid(Str1, t5 + 1, t6 - t5 - 1)))
		If Len(ValueL) = 1 Then ValueL = "0" & ValueL
		If ValueH & ValueL = SendData Then
			Return True
		Else
			Return False
		End If
		Exit Function

Comm_Error:
		Return False
		Exit Function

WRITE_COMM_TIMEOUT:
		If Retry < 3 Then
			Retry = Retry + 1
			Thread.Sleep(10)
			GoTo WRITERETRY
		End If
		Return False
		Exit Function
	End Function

	Public Function WR_MULTI_RFID(AddrNos As String, SendData As String) As Boolean
		Dim n As Integer
		Dim ii() As String
		Dim SendString(19) As Byte
		Dim intLenH, intLenL As Integer
		Dim ValueH, ValueL As String
		Dim CheckSum As String
		Dim Str1, Str2 As String
		Dim Bcount As Integer
		Dim t1, t2, t3, t4, t5, t6 As Integer
		Dim ErrCode As String
		Dim Retry As Integer

		On Error GoTo WRITE_COMM_TIMEOUT

WRITERETRY:
		ii = {"&HF8", "&H10", "&H", "&H", "&H00", "&0AH", "&14H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H", "&H"}
		ii(2) = ii(2) + Left(AddrNos, 2)
		ii(3) = ii(3) + Right(AddrNos, 2)
		For i As Integer = 0 To 19
			ii(i + 7) = Mid(SendData, (i * 2) + 1, 2)
		Next
		'ii(7) = ii(7) + Left(SendData, 2)
		'ii(8) = ii(8) + Right(SendData, 2)

		CheckSum = CRC(ii) 'Calculate Checksum
		Str1 = Right(CheckSum, Len(CheckSum) - 2)
		Select Case Len(Str1)
			Case 1
				ii(28) = ii(28) + "00"
				ii(27) = ii(27) + "0" + Str1
			Case 2
				ii(28) = ii(28) + "00"
				ii(27) = ii(27) + Str1
			Case 3
				ii(27) = ii(27) + Right(Str1, 2)
				ii(28) = ii(28) + "0" + Left(Str1, Len(Str1) - 2)
			Case 4
				ii(27) = ii(27) + Right(Str1, 2)
				ii(28) = ii(28) + Left(Str1, Len(Str1) - 2)
		End Select

		For i As Integer = 0 To 19
			SendString(i) = CByte(ii(i))
		Next

		frmMain.RFID_Comm.Write(SendString, 0, SendString.Length)
		Thread.Sleep(20)
		GoTo CHECKINBYTE

CHECKINBYTE:
		If frmMain.RFID_Comm.BytesToRead = 0 Then GoTo WRITE_COMM_TIMEOUT
		Thread.Sleep(50)
		Bcount = frmMain.RFID_Comm.BytesToRead
		Str1 = ""
		Do
			Str2 = frmMain.RFID_Comm.ReadExisting()
			n = n + 1
			Str2 = Asc(Str2)
			Str1 = Str1 + Trim(Str(Str2)) + ","
		Loop Until n = Bcount
		t1 = InStr(1, Str1, ",")
		t2 = InStr(t1 + 1, Str1, ",")
		t3 = InStr(t2 + 1, Str1, ",")
		t4 = InStr(t3 + 1, Str1, ",")
		t5 = InStr(t4 + 1, Str1, ",")
		t6 = InStr(t5 + 1, Str1, ",")
		If Mid(Str1, t1 + 1, t2 - t1 - 1) = "134" Then '&H83 or &H84 or &H86
			ErrCode = Mid(Str1, t2 + 1, t3 - t2 - 1)
			GoTo Comm_Error
		End If

		Return True
		Exit Function

Comm_Error:
		Return False
		Exit Function

WRITE_COMM_TIMEOUT:
		If Retry < 3 Then
			Retry = Retry + 1
			Thread.Sleep(10)
			GoTo WRITERETRY
		End If
		Return False
		Exit Function
	End Function

	Public Function RD_MULTI_RFID(AddrNos As String, WordLength As Integer) As String
		Dim temp1, temp2 As String
		Dim n As Integer
		Dim ii() As String
		Dim SendString(8) As Byte
		Dim intLenH, intLenL As Integer
		Dim ValueH, ValueL As String
		Dim LenH, LenL As String
		Dim CheckSum As String
		Dim Str1, Str2 As String
		Dim Bcount As Integer
		Dim t(25) As Integer
		Dim ErrCode As String
		Dim Retry As Integer

		On Error GoTo READ_RFID_TIMEOUT
READRETRY:
		ii = {"&HF8", "&H03", "&H", "&H", "&H", "&H", "&H", "&H"}
		ii(2) = ii(2) + Left(AddrNos, 2)
		ii(3) = ii(3) + Right(AddrNos, 2)
		intLenH = WordLength / 256
		LenH = Hex(intLenH)
		intLenL = WordLength Mod 256
		LenL = Hex(intLenL)
		ii(4) = ii(4) + LenH
		ii(5) = ii(5) + LenL
		CheckSum = CRC(ii) 'Calculate the Checksum
		Str1 = Right(CheckSum, Len(CheckSum) - 2)
		Select Case Len(Str1)
			Case 1
				ii(7) = ii(7) + "00"
				ii(6) = ii(6) + "0" + Str1
			Case 2
				ii(7) = ii(7) + "00"
				ii(6) = ii(6) + Str1
			Case 3
				ii(6) = ii(6) + Right(Str1, 2)
				ii(7) = ii(7) + "0" + Left(Str1, Len(Str1) - 2)
			Case 4
				ii(6) = ii(6) + Right(Str1, 2)
				ii(7) = ii(7) + Left(Str1, Len(Str1) - 2)
		End Select

		For i As Integer = 0 To 7
			SendString(i) = CByte(ii(i))
		Next

		frmMain.RFID_Comm.Write(SendString, 0, SendString.Length)
		Thread.Sleep(20)
		GoTo READ_RFID_TIMEOUT

READINBYTE:
		Str1 = ""
		If frmMain.RFID_Comm.BytesToRead = 0 Then GoTo READ_RFID_TIMEOUT
		Thread.Sleep(10)
		Bcount = frmMain.RFID_Comm.BytesToRead
		Do
			Str2 = frmMain.RFID_Comm.ReadExisting()
			'If Str2 = "" Then GoTo skip
			n = n + 1
			Str2 = Asc(Str2)
			Str1 = Str1 + Trim(Str(Str2)) + "," 'Build the string received frm OsiTrack
		Loop Until n = Bcount

skip:
		t(0) = InStr(1, Str1, ",")
		For i As Integer = 1 To (WordLength * 2 + 4)
			t(i) = InStr(t(i - 1) + 1, Str1, ",")
		Next
		temp1 = Mid(Str1, t(0) + 1, t(1) - t(0) - 1)
		temp2 = Mid(Str1, t(0) + 1, t(1) - t(0) - 1)
		If Mid(Str1, t(0) + 1, t(1) - t(0) - 1) = "131" Or Mid(Str1, t(0) + 1, t(1) - t(0) - 1) = "132" Then '&H83 or &H84
			ErrCode = Mid(Str1, t(1) + 1, t(2) - t(1) - 1)
			GoTo Comm_Error
		End If
		For i As Integer = 1 To (WordLength * 2)
			ValueL = Mid(Str1, t(i + 1) + 1, t(i + 2) - t(i + 1) - 1)
			If ValueL <> 0 Then
				ValueH = ValueH & Chr(CInt(ValueL))
			End If
		Next
		'ValueH = Mid(Str1, t3 + 1, t4 - t3 - 1)
		'ValueH = Hex(Val(ValueH))
		'ValueL = Mid(Str1, t4 + 1, t5 - t4 - 1)
		'ValueL = Hex(Val(ValueL))
		'If Len(ValueL) = 1 Then
		'    ValueL = "0" + ValueL
		'End If
		'ReadBuf = ValueH + ValueL 'Read the LSB only(1 Byte)
		'ReadBuf = Hex2Bin(ReadBuf)
		'ReadBuf = Str(Bin2Dec(ReadBuf))
		'RD_RFID = Trim(ReadBuf)
		Return ValueH
		Exit Function

Comm_Error:
		Return "NOK"
		Exit Function

READ_RFID_TIMEOUT:
		If Retry < 3 Then
			Retry = Retry + 1
			GoTo READRETRY
		End If
		Return "NOK"
		Exit Function
	End Function
End Module
