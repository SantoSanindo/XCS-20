Option Explicit On
Module modReadWrite
    Public INISTATUSPATH As String 'PATH TO SERVER\FRIDGE
    Public INITEMPLATEPATH As String 'PATH TO LABEL Template
    Public INIPHOTOPATH As String 'PATH TO THE PHOTO FOR THE LABEL TEMPLATE
    Public INIPSNFOLDERPATH As String 'Path to the PSN.Txt
    Public INIMATERIALPATH As String

    Public INIACHIEVEPATH As String 'Path to achieve folder

    Dim Fnum As Integer
    Dim LineStr As String
    Public Sub ReadINI(Filename As String)
        Dim ItemStr As String
        Dim SectionHeading As String
        Dim pos As Integer

        Fnum = FreeFile()

        FileOpen(Fnum, Filename, OpenMode.Input)

        Do While Not EOF(Fnum)
            LineStr = LineInput(Fnum)
            If Left(LineStr, 1) = "[" Then
                SectionHeading = Mid(LineStr, 2, Len(LineStr) - 2)
            Else
                If InStr(LineStr, "=") > 0 Then
                    pos = InStr(LineStr, "=")
                    ItemStr = Left$(LineStr, pos - 1)

                    Select Case UCase(SectionHeading)
                        Case "LABEL PHOTO PATH" 'Shared FILE
                            Select Case UCase(ItemStr)
                                Case "PATH" : INIPHOTOPATH = Mid$(LineStr, pos + 1)
                            End Select

                        Case "LABEL TEMPLATE PATH" 'Shared FILE
                            Select Case UCase(ItemStr)
                                Case "PATH" : INITEMPLATEPATH = Mid$(LineStr, pos + 1)
                            End Select

                        Case "PSN FOLDER PATH" 'Share FILE
                            Select Case UCase(ItemStr)
                                Case "PATH" : INIPSNFOLDERPATH = Mid$(LineStr, pos + 1)
                            End Select

                        Case "MATERIAL PATH" 'RACK MATERIAL LIST
                            Select Case UCase(ItemStr)
                                Case "PATH" : INIMATERIALPATH = Mid$(LineStr, pos + 1)
                            End Select

                        Case "ACHIEVE PATH"
                            Select Case UCase(ItemStr)
                                Case "PATH" : INIACHIEVEPATH = Mid$(LineStr, pos + 1)
                            End Select

                        Case "STATUS PATH"
                            Select Case UCase(ItemStr)
                                Case "PATH" : INISTATUSPATH = Mid$(LineStr, pos + 1)
                            End Select
                    End Select
                End If
            End If
        Loop
        FileClose(Fnum)
    End Sub
    Public Function LOADPSNFILE(ProductPSN As String) As Boolean
        Dim ItemStr As String
        Dim SectionHeading As String
        Dim pos1, pos2, pos3 As Integer

        Fnum = FreeFile()

        If Dir(INIPSNFOLDERPATH & ProductPSN & ".Txt") = "" Then
            'SetDefaultINIValues
            'WriteINI
            Exit Function
        End If

        FileOpen(Fnum, INIPSNFOLDERPATH & ProductPSN & ".Txt", OpenMode.Input)
        Do While Not EOF(Fnum)
            LineStr = LineInput(Fnum)
            'Check for Section heading
            If InStr(LineStr, "[") > 0 And InStr(LineStr, "]") > 0 Then
                pos1 = InStr(LineStr, "[")
                pos2 = InStr(LineStr, "]")
                pos3 = InStr(LineStr, ":")

                SectionHeading = Mid(LineStr, pos1 + 1, pos2 - pos1 - 1)

                Select Case UCase(SectionHeading)
                    Case "MODEL"
                        PSNFileInfo.MODELNAME = Trim(Mid(LineStr, pos3 + 1))

                    Case "DATE CREATED"
                        PSNFileInfo.DateCreated = Trim(Mid(LineStr, pos3 + 1))

                    Case "DATE COMPLETED"
                        PSNFileInfo.DateCompleted = Trim(Mid(LineStr, pos3 + 1))

                    Case "OPERATOR ID"
                        PSNFileInfo.OperatorID = Trim(Mid(LineStr, pos3 + 1))

                    Case "WORK ORDER NO"
                        PSNFileInfo.WONos = Trim(Mid(LineStr, pos3 + 1))

                    Case "MAIN PCBA S/N"
                        PSNFileInfo.MainPCBA = Trim(Mid(LineStr, pos3 + 1))

                    Case "SECONDARY PCBA S/N"
                        PSNFileInfo.SecondaryPCBA = Trim(Mid(LineStr, pos3 + 1))

                    Case "ELECTROMAGNET S/N"
                        PSNFileInfo.ElectroMagnet = Trim(Mid(LineStr, pos3 + 1))

                    Case "PSN"
                        PSNFileInfo.PSN = Trim(Mid(LineStr, pos3 + 1))

                    Case "BODY ASSY STATION CHECK IN DATE"
                        PSNFileInfo.BodyAssyCheckIn = Trim(Mid(LineStr, pos3 + 1))

                    Case "BODY ASSY STATION CHECK OUT DATE"
                        PSNFileInfo.BodyAssyCheckOut = Trim(Mid(LineStr, pos3 + 1))

                    Case "BODY ASSY STATION STATUS"
                        PSNFileInfo.BodyAssyStatus = Trim(Mid(LineStr, pos3 + 1))

                    Case "SCREWING STATION CHECK IN DATE"
                        PSNFileInfo.ScrewStnCheckIn = Trim(Mid(LineStr, pos3 + 1))

                    Case "SCREWING STATION CHECK OUT DATE"
                        PSNFileInfo.ScrewStnCheckOut = Trim(Mid(LineStr, pos3 + 1))

                    Case "SCREWING STATION STATUS"
                        PSNFileInfo.ScrewStnStatus = Trim(Mid(LineStr, pos3 + 1))

                    Case "FINAL TEST CHECK IN DATE"
                        PSNFileInfo.FTCheckIn = Trim(Mid(LineStr, pos3 + 1))

                    Case "FINAL TEST CHECK OUT DATE"
                        PSNFileInfo.FTCheckOut = Trim(Mid(LineStr, pos3 + 1))

                    Case "FINAL TEST SYNCH MEASUREMENT"
                        PSNFileInfo.FTSycMeas = Trim(Mid(LineStr, pos3 + 1))

                    Case "FINAL TEST STATUS"
                        PSNFileInfo.FTStatus = Trim(Mid(LineStr, pos3 + 1))

                    Case "STATION 5 CHECK IN DATE"
                        PSNFileInfo.Stn5CheckIn = Trim(Mid(LineStr, pos3 + 1))

                    Case "STATION 5 CHECK OUT DATE"
                        PSNFileInfo.Stn5CheckOut = Trim(Mid(LineStr, pos3 + 1))

                    Case "STATION 5 STATUS"
                        PSNFileInfo.Stn5Status = Trim(Mid(LineStr, pos3 + 1))

                    Case "VACUUM CHECK IN DATE"
                        PSNFileInfo.VacuumCheckIn = Trim(Mid(LineStr, pos3 + 1))

                    Case "VACUUM CHECK OUT DATE"
                        PSNFileInfo.VacummCheckOut = Trim(Mid(LineStr, pos3 + 1))

                    Case "VACUUM STATUS"
                        PSNFileInfo.VacuumStatus = Trim(Mid(LineStr, pos3 + 1))

                    Case "CONNECTOR TEST CHECK IN DATE"
                        PSNFileInfo.ConnTestCheckIn = Trim(Mid(LineStr, pos3 + 1))

                    Case "CONNECTOR TEST CHECK OUT DATE"
                        PSNFileInfo.ConnTestCheckOut = Trim(Mid(LineStr, pos3 + 1))

                    Case "CONNECTOR TEST STATUS"
                        PSNFileInfo.ConnTestStatus = Trim(Mid(LineStr, pos3 + 1))

                    Case "VACUUM #2 CHECK IN DATE"
                        PSNFileInfo.Vacuum2CheckIn = Trim(Mid(LineStr, pos3 + 1))

                    Case "VACUUM #2 CHECK OUT DATE"
                        PSNFileInfo.Vacumm2CheckOut = Trim(Mid(LineStr, pos3 + 1))

                    Case "VACUUM #2 STATUS"
                        PSNFileInfo.Vacuum2Status = Trim(Mid(LineStr, pos3 + 1))

                    Case "PACKAGING CHECK IN DATE"
                        PSNFileInfo.PackagingCheckIn = Trim(Mid(LineStr, pos3 + 1))

                    Case "PACKAGING CHECK OUT DATE"
                        PSNFileInfo.PackagingCheckOut = Trim(Mid(LineStr, pos3 + 1))

                    Case "PACKAGING STATUS"
                        PSNFileInfo.PackagingStatus = Trim(Mid(LineStr, pos3 + 1))

                    Case "DEBUG STATION #10 STATUS"
                        PSNFileInfo.DebugStatus = Trim(Mid(LineStr, pos3 + 1))

                    Case "DEBUG COMMENTS"
                        PSNFileInfo.DebugComment = Trim(Mid(LineStr, pos3 + 1))

                    Case "DEBUG TECHNICIANS ID"
                        PSNFileInfo.DebugTechnican = Trim(Mid(LineStr, pos3 + 1))

                    Case "DEBUG DATE REPAIRED"
                        PSNFileInfo.RepairDate = Trim(Mid(LineStr, pos3 + 1))

                End Select
            End If
        Loop
        FileClose(Fnum)
        Return True
    End Function
    'Write data to PSN file
    Public Function WRITEPSNFILE(ProductPSN As String) As Boolean
        On Error GoTo ErrorHandler
        Fnum = FreeFile()
        FileOpen(Fnum, INIPSNFOLDERPATH & ProductPSN & ".txt", OpenMode.Output)

        PrintLine(Fnum)
        PrintLine(Fnum, "[MODEL] : " & PSNFileInfo.MODELNAME)
        PrintLine(Fnum)
        PrintLine(Fnum, "[DATE CREATED] : " & PSNFileInfo.DateCreated)
        PrintLine(Fnum)
        PrintLine(Fnum, "[DATE COMPLETED] : " & PSNFileInfo.DateCompleted)
        PrintLine(Fnum)
        PrintLine(Fnum, "[OPERATOR ID] : " & PSNFileInfo.OperatorID)
        PrintLine(Fnum)
        PrintLine(Fnum, "[WORK ORDER NO] : " & PSNFileInfo.WONos)
        PrintLine(Fnum)
        PrintLine(Fnum, "[MAIN PCBA S/N] : " & PSNFileInfo.MainPCBA)
        PrintLine(Fnum)
        PrintLine(Fnum, "[SECONDARY PCBA S/N] : " & PSNFileInfo.SecondaryPCBA)
        PrintLine(Fnum)
        PrintLine(Fnum, "[ELECTROMAGNET S/N] : " & PSNFileInfo.ElectroMagnet)
        PrintLine(Fnum)
        PrintLine(Fnum, "[PSN] : " & PSNFileInfo.PSN)
        PrintLine(Fnum)
        PrintLine(Fnum, "[BODY ASSY STATION CHECK IN DATE] : " & PSNFileInfo.BodyAssyCheckIn)
        PrintLine(Fnum)
        PrintLine(Fnum, "[BODY ASSY STATION CHECK OUT DATE] : " & PSNFileInfo.BodyAssyCheckOut)
        PrintLine(Fnum)
        PrintLine(Fnum, "[BODY ASSY STATION STATUS] : " & PSNFileInfo.BodyAssyStatus)
        PrintLine(Fnum)
        PrintLine(Fnum, "[SCREWING STATION CHECK IN DATE] : " & PSNFileInfo.ScrewStnCheckIn)
        PrintLine(Fnum)
        PrintLine(Fnum, "[SCREWING STATION CHECK OUT DATE] : " & PSNFileInfo.ScrewStnCheckOut)
        PrintLine(Fnum)
        PrintLine(Fnum, "[SCREWING STATION STATUS] : " & PSNFileInfo.ScrewStnStatus)
        PrintLine(Fnum)
        PrintLine(Fnum, "[FINAL TEST CHECK IN DATE] : " & PSNFileInfo.FTCheckIn)
        PrintLine(Fnum)
        PrintLine(Fnum, "[FINAL TEST CHECK OUT DATE] : " & PSNFileInfo.FTCheckOut)
        PrintLine(Fnum)
        PrintLine(Fnum, "[FINAL TEST SYNCH MEASUREMENT] : " & PSNFileInfo.FTSycMeas)
        PrintLine(Fnum)
        PrintLine(Fnum, "[FINAL TEST STATUS] : " & PSNFileInfo.FTStatus)
        PrintLine(Fnum)
        PrintLine(Fnum, "[STATION 5 CHECK IN DATE] : " & PSNFileInfo.Stn5CheckIn)
        PrintLine(Fnum)
        PrintLine(Fnum, "[STATION 5 CHECK OUT DATE] : " & PSNFileInfo.Stn5CheckOut)
        PrintLine(Fnum)
        PrintLine(Fnum, "[STATION 5 STATUS] : " & PSNFileInfo.Stn5Status)
        PrintLine(Fnum)
        PrintLine(Fnum, "[VACUUM CHECK IN DATE] : " & PSNFileInfo.VacuumCheckIn)
        PrintLine(Fnum)
        PrintLine(Fnum, "[VACUUM CHECK OUT DATE] : " & PSNFileInfo.VacummCheckOut)
        PrintLine(Fnum)
        PrintLine(Fnum, "[VACUUM STATUS] : " & PSNFileInfo.VacuumStatus)
        PrintLine(Fnum)
        PrintLine(Fnum, "[CONNECTOR TEST CHECK IN DATE] : " & PSNFileInfo.ConnTestCheckIn)
        PrintLine(Fnum)
        PrintLine(Fnum, "[CONNECTOR TEST CHECK OUT DATE] : " & PSNFileInfo.ConnTestCheckOut)
        PrintLine(Fnum)
        PrintLine(Fnum, "[CONNECTOR TEST STATUS] : " & PSNFileInfo.ConnTestStatus)
        PrintLine(Fnum)
        PrintLine(Fnum, "[VACUUM #2 CHECK IN DATE] : " & PSNFileInfo.Vacuum2CheckIn)
        PrintLine(Fnum)
        PrintLine(Fnum, "[VACUUM #2 CHECK OUT DATE] : " & PSNFileInfo.Vacumm2CheckOut)
        PrintLine(Fnum)
        PrintLine(Fnum, "[VACUUM #2 STATUS] : " & PSNFileInfo.Vacuum2Status)
        PrintLine(Fnum)
        PrintLine(Fnum, "[PACKAGING CHECK IN DATE] : " & PSNFileInfo.PackagingCheckIn)
        PrintLine(Fnum)
        PrintLine(Fnum, "[PACKAGING CHECK OUT DATE] : " & PSNFileInfo.PackagingCheckOut)
        PrintLine(Fnum)
        PrintLine(Fnum, "[PACKAGING STATUS] : " & PSNFileInfo.PackagingStatus)
        PrintLine(Fnum)
        PrintLine(Fnum, "[DEBUG STATION #10 STATUS] : " & PSNFileInfo.DebugStatus)
        PrintLine(Fnum)
        PrintLine(Fnum, "[DEBUG COMMENTS] : " & PSNFileInfo.DebugComment)
        PrintLine(Fnum)
        PrintLine(Fnum, "[DEBUG TECHNICIANS ID] : " & PSNFileInfo.DebugTechnican)
        PrintLine(Fnum)
        PrintLine(Fnum, "[DEBUG DATE REPAIRED] : " & PSNFileInfo.RepairDate)
        Return True
        Exit Function
ErrorHandler:
        Return False
    End Function
End Module
