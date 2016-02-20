' This is the host application for evaluating the Arduino Pic Programmer
' 
' Gregor Schlechtriem
' Copyright (C) 2014 - 2016
' www.pikoder.com
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version. See <http://www.gnu.org/licenses/>
' for more details.

' You may incorporate all or parts of this program into your own project.
' Keep in mind that you got this code free and as such the author has no 
' responsibility to you, your customers, or anyone else regarding the use
' of this program code. No support or warranty of any kind can be provided. All
' liability falls upon the user of this code and thus you must judge the suitability
' of this code to your application and usage.
' 
' Documentation of changes:
' Last change: 01/26/2016 : migrated source to VB 2013 and fixed a bug in detecting a device
' Last change: 11/02/2014 : fixed bugs in COM detection
' Last change: 01/11/2014 : fixed bugs in Intel hex file export (checksum calculation, format for EEPROM data)
'

Option Strict Off
Option Explicit On
Friend Class frmUI_ArdPicProgHost
    Inherits System.Windows.Forms.Form

    Private mySerialLink As New SerialLink

    ' declaration of variables for program control
    Dim boolErrorFlag As Boolean ' global flag for errors in communication
    Dim boolDeviceFound As Boolean = False ' global flag indicating device found in socket
    Dim boolSourceFileImportedOK As Boolean = False ' global flag indicating source file availability

    'declaration of variables for memory management
    Dim strProgramMemory As String
    Dim strDataMemory As String
    Dim strConfigMemory As String
    Dim strWriteFileFormat As String = ""
    Dim intProgramStart As Integer = 0
    Dim intProgramEnd As Integer = &H7FF
    Dim intConfigStart As Integer = &H2000
    Dim intConfigEnd As Integer = &H2007
    Dim intDataStart As Integer = &H2100
    Dim intDataEnd As Integer = &H217F
    Dim idx As Integer

    ' declaration of subroutines
    ''' <summary>
    ''' This method is used to initialize the application. It will collect all
    ''' available COM ports in a list box. If no ports were found then the program
    ''' will stop.
    ''' </summary>
    ''' <remarks></remarks>

    Private Sub MyForm_Initialize()

        Led2.Color = LED.LEDColorSelection.LED_Green

        Dim myStringBuffer As String ' used for temporary storage of COM port designator
        ' define and initialize array for sorting COM ports 
        Dim myCOMPorts(99) As String ' assume we have a max of 100 COM Ports

        For Each s As String In myCOMPorts
            s = ""
        Next

        boolErrorFlag = False

        ' Connection setup - check for ports and display result
        ' Connection setup - check for ports and display result
        For Each sp As String In My.Computer.Ports.SerialPortNames
            myStringBuffer = ""
            idx = 0
            For intI = 0 To Len(sp) - 1
                If ((sp(intI) >= "A" And (sp(intI) <= "Z")) Or IsNumeric(sp(intI))) Then
                    myStringBuffer = myStringBuffer + sp(intI)
                    If IsNumeric(sp(intI)) Then idx = (idx * 10) + Val(sp(intI))
                End If
            Next
            myCOMPorts(idx) = myStringBuffer
        Next

        For Each s As String In myCOMPorts
            If s <> "" Then AvailableCOMPorts.Items.Add(s)
        Next

        ' Set active control 
        Me.ActiveControl = Me.AvailableCOMPorts

        ' Initialize LED
        Led1.Color = LED.LEDColorSelection.LED_Green
        Led1.Interval = 500
        
        If AvailableCOMPorts.Items.Count = 0 Then
            MsgBox("No COM Port found. Program will stop.", MsgBoxStyle.OkOnly, "Error Message")
            End
        End If
    End Sub
    Private Sub RetrieveArdPicProgParameters()

        Dim strChannelBuffer As String = ""

        If mySerialLink.SerialLinkConnected() Then
            If Not boolErrorFlag Then
                'retrieve channel information first to make sure that connection is available
                Call mySerialLink.SendDataToSerial("PROGRAM_PIC_VERSION" & vbCrLf)
                Call mySerialLink.GetResponse(strChannelBuffer, vbCrLf)
                If strChannelBuffer <> "TimeOut" Then
                    strArdPicProgFirmware.Text = Microsoft.VisualBasic.Mid(strChannelBuffer, 12, 3)
                Else ' error message
                    boolErrorFlag = True
                End If
            End If
        End If
    End Sub
    Private Sub ParseForRange(ByVal InputString As String, ByRef intLowerBoundary As Integer, ByRef intUpperBoundary As Integer)
        intLowerBoundary = Val("&H" + Microsoft.VisualBasic.Left(InputString, 4))
        intUpperBoundary = Val("&H" + Mid(InputString, 6, 4))
    End Sub
    Private Sub ParseForArgument(ByRef InputString As String, ByRef OutputString As String, ByVal ReferenceString As String)
        Dim i As Integer = 1
        OutputString = ""
        i = InStr(1, InputString, ReferenceString)
        If i > 0 Then
            i = i + Len(ReferenceString) + 1
            While (InputString(i) > Chr(31))
                OutputString = OutputString + Microsoft.VisualBasic.UCase(InputString(i))
                i = i + 1
            End While
        End If
    End Sub
    Private Sub RetrieveDeviceParameters()
        Dim strChannelBuffer As String = ""
        Dim i As Integer = 1
        If mySerialLink.SerialLinkConnected() Then
            If Not boolErrorFlag Then
                'retrieve channel information first to make sure that connection is available
                Call mySerialLink.SendDataToSerial("DEVICE" & vbCrLf)
                Call mySerialLink.GetResponse(strChannelBuffer, ".")
                If ((strChannelBuffer <> "TimeOut") And Not (InStr(1, strChannelBuffer, "ERROR", 1) > 0)) Then 'catch ERROR conditions
                    'parse string for information received
                    Call ParseForArgument(strChannelBuffer, lDeviceName.Text, "DeviceName")
                    Call ParseForArgument(strChannelBuffer, lProgramRange.Text, "ProgramRange")
                    Call ParseForRange(lProgramRange.Text, intProgramStart, intProgramEnd)
                    Call ParseForArgument(strChannelBuffer, lEEPROMRange.Text, "DataRange")
                    Call ParseForRange(lEEPROMRange.Text, intDataStart, intDataEnd)
                    Call ParseForArgument(strChannelBuffer, lConfigWord.Text, "ConfigWord")
                    Call ParseForArgument(strChannelBuffer, lConfigurationRange.Text, "ConfigRange")
                    Call ParseForRange(lConfigurationRange.Text, intConfigStart, intConfigEnd)
                    If lDeviceName.Text = "PIC12F675" Then
                        Call mySerialLink.SendDataToSerial("READ" & " " & "03FF" & vbCrLf)
                        Call mySerialLink.GetResponse(strProgramMemory, "." & vbCrLf)
                        If strProgramMemory <> "TimeOut" And strProgramMemory.Length = 13 Then
                            lOsccal.Text = strProgramMemory.Substring(4, 4)
                        End If
                    End If
                    boolDeviceFound = True
                    Status.Text = "Connected"
                    Led2.Color = LED.LEDColorSelection.LED_Green
                    Led2.State = True
                Else ' error message - delete all fields and information
                    lDeviceName.Text = ""
                    lProgramRange.Text = ""
                    lEEPROMRange.Text = ""
                    lConfigWord.Text = ""
                    lConfigurationRange.Text = ""
                    Status.Text = ""
                    Led2.Color = LED.LEDColorSelection.LED_Green
                    Led2.State = False
                    boolDeviceFound = False
                End If
            End If
        End If
    End Sub

    Private Sub AvailableCOMPorts_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AvailableCOMPorts.SelectedIndexChanged

        Dim strChannelBuffer As String = ""

        If (AvailableCOMPorts.SelectedItem = Nothing) Then Exit Sub

        Led1.Blink = True

        If (mySerialLink.EstablishSerialLink(AvailableCOMPorts.SelectedItem)) Then
            Led1.Blink = False ' indicate that the connection established
            Led1.State = True
            RetrieveArdPicProgParameters() 'retrieve parameters
            RetrieveDeviceParameters() 'check for device
        Else ' error message
            Led1.Blink = False
        End If
    End Sub

    Private Sub ImportButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportButton.Click
        Dim strSourceFileIntelHex As String
        Dim strDataAddress As String
        Dim intFilePointer As Integer = 0
        Dim intByteCount As Integer
        Dim intDataAddress As Integer
        With dlgOpenFile
            .Filter = "Hexfiles (*.hex)|*.hex|All files (*.*)|*.*"
            .FilterIndex = 1
            .FileName = ""

            Dim dlgResult As DialogResult = .ShowDialog()
            If DialogResult = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If
        End With

        Try
            strSourceFileIntelHex = My.Computer.FileSystem.ReadAllText(dlgOpenFile.FileName)
            HexFileName.Text = dlgOpenFile.FileName
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cannot open file.", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        If strSourceFileIntelHex(0) <> ":" Then
            MsgBox("Format error in Source File!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        While (strSourceFileIntelHex.Chars(intFilePointer + 8) <> "1")
            intFilePointer += 1
            intByteCount = Val("&H" + strSourceFileIntelHex.Substring(intFilePointer, 2))
            intFilePointer += 2
            intDataAddress = Val("&H" + strSourceFileIntelHex.Substring(intFilePointer, 4)) >> 1
            strDataAddress = Decimal2Hex(intDataAddress)
            intFilePointer += 5 'move filepointer to record type
            If strSourceFileIntelHex.Chars(intFilePointer) = "0" Then
                'process data
                intFilePointer += 1
                strWriteFileFormat = strWriteFileFormat & ":" & strDataAddress
                For i = 1 To intByteCount >> 1
                    strWriteFileFormat = strWriteFileFormat & " " & strSourceFileIntelHex.Substring(intFilePointer + 2, 2) & strSourceFileIntelHex.Substring(intFilePointer, 2)
                    intFilePointer += 4
                Next
                strWriteFileFormat = strWriteFileFormat & vbCrLf 'close line
            ElseIf strSourceFileIntelHex.Chars(intFilePointer) = "4" Then
                'process extended segment address -> not required
            End If
            ' read until start of next line
            While (strSourceFileIntelHex.Chars(intFilePointer) <> ":")
                intFilePointer += 1
            End While
        End While
        boolSourceFileImportedOK = True
    End Sub

    Private Sub ConnectButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConnectButton.Click
        Call RetrieveDeviceParameters()
    End Sub
    Private Function Decimal2Hex(ByVal Dec As Integer) As String
        Return Dec.ToString("X4")
    End Function
    Private Sub UpdateMemoryWindows(ByVal sMemMessage As String, ByVal sEEPMessage As String)
        Dim boolLedStatus = True
        Dim intUpdateWindow As Integer = 3
        Dim intStartOfLine As Integer = 0
        Dim intCharsPerLine As Integer = 40
        Dim intMemAdress As Integer = 0
        Dim intRefreshCounter = 0
        HexFileDump.Text = ""
        EEPROMDump.Text = ""
        Me.Refresh()
        intMemAdress = intProgramStart
        Led2.Color = LED.LEDColorSelection.LED_Red
        While (intMemAdress < intProgramEnd)
            Call mySerialLink.SendDataToSerial("READ" & " " & Decimal2Hex(intMemAdress) & "-" & Decimal2Hex(intMemAdress + 15) & vbCrLf)
            Call mySerialLink.GetResponse(strProgramMemory, "." & vbCrLf)
            If strProgramMemory <> "TimeOut" And strProgramMemory.Length = 89 Then
                HexFileDump.Text = HexFileDump.Text & Decimal2Hex(intMemAdress) & ": " & strProgramMemory.Substring(4, intCharsPerLine) & " "
                HexFileDump.Text = HexFileDump.Text & strProgramMemory.Substring(45, intCharsPerLine) & vbCrLf
                intMemAdress += 16 'address for next line
                intRefreshCounter += 1
                If intRefreshCounter = 4 Then
                    If intUpdateWindow > 0 Then
                        Me.HexFileDump.Refresh()
                        intUpdateWindow -= 1
                    End If
                    intRefreshCounter = 0
                    Led2.State = boolLedStatus
                    boolLedStatus = Not boolLedStatus
                    Status.Text = sMemMessage & Decimal2Hex(intMemAdress)
                    Me.Led2.Refresh()
                    Me.Status.Refresh()
                End If
            End If
        End While
        Status.Text = sEEPMessage
        Me.Status.Refresh()
        Call mySerialLink.SendDataToSerial("READ" & " " & lEEPROMRange.Text & vbCrLf)
        Call mySerialLink.GetResponse(strDataMemory, ".")
        intMemAdress = intDataStart
        intStartOfLine = 1
        If strDataMemory <> "TimeOut" Then
            strDataMemory = Mid(strDataMemory, 5, Len(strDataMemory) - 9) 'remove "OK" at the beginning and CRLF, ".", and CRLF at the end 
            While (intMemAdress < intDataEnd)
                EEPROMDump.Text = EEPROMDump.Text & Decimal2Hex(intMemAdress) & ": "
                For i = intStartOfLine + 2 To intStartOfLine + 33 Step 5
                    EEPROMDump.Text = EEPROMDump.Text & Mid(strDataMemory, i, 3)
                Next
                EEPROMDump.Text = EEPROMDump.Text & Mid(strDataMemory, intStartOfLine + 37, 2) & " "
                intStartOfLine = intStartOfLine + intCharsPerLine + 1 'move pointer and skip vbCrLf
                For i = intStartOfLine + 2 To intStartOfLine + 33 Step 5
                    EEPROMDump.Text = EEPROMDump.Text & Mid(strDataMemory, i, 3)
                Next
                EEPROMDump.Text = EEPROMDump.Text & Mid(strDataMemory, intStartOfLine + 37, 2) & vbCrLf
                intStartOfLine = intStartOfLine + intCharsPerLine + 1 'move pointer and skip vbCrLf
                intMemAdress += 16 'address for next line
            End While
        End If
        Led2.Color = LED.LEDColorSelection.LED_Green
        Led2.State = True 'make sure that LED is on
    End Sub
        Private Sub ReadButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReadButton.Click
        Dim strChannelBuffer As String = ""
        If boolDeviceFound Then
            Status.Text = "Reading"
            Me.Status.Refresh()
            UpdateMemoryWindows("Reading address ", "Reading EEPROM")
            Call mySerialLink.SendDataToSerial("DEVICE" & vbCrLf)
            Call mySerialLink.GetResponse(strChannelBuffer, ".")
            If ((strChannelBuffer <> "TimeOut") And Not (InStr(1, strChannelBuffer, "ERROR"))) Then 'catch ERROR conditions
                'parse string for information received
                Call ParseForArgument(strChannelBuffer, lConfigWord.Text, "ConfigWord")
            End If
            Status.Text = "Connected"
            Me.Status.Refresh()
        End If
    End Sub

    Private Sub EraseButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EraseButton.Click
        Dim strChannelBuffer As String = ""
        If boolDeviceFound Then
            Led2.Color = LED.LEDColorSelection.LED_Red
            Led2.Interval = 500
            Led2.Blink = True
            Status.Text = "Erasing"
            Me.Refresh()
            Call mySerialLink.SendDataToSerial("ERASE" & vbCrLf)
            Call mySerialLink.GetResponse(strProgramMemory, "OK")
            If strChannelBuffer = "TimeOut" Then 'Error message 
                MsgBox("Time out error erasing device - Please retry!", MessageBoxButtons.OK, "ArdPigProgHost: Communication time out")
            Else
                UpdateMemoryWindows("Erasing address ", "Erasing EEPROM")
                Call mySerialLink.SendDataToSerial("DEVICE" & vbCrLf)
                Call mySerialLink.GetResponse(strChannelBuffer, ".")
                If ((strChannelBuffer <> "TimeOut") And Not (InStr(1, strChannelBuffer, "ERROR"))) Then 'catch ERROR conditions
                    'parse string for information received
                    Call ParseForArgument(strChannelBuffer, lConfigWord.Text, "ConfigWord")
                End If
            End If
            Led2.Blink = False
            Led2.Color = LED.LEDColorSelection.LED_Green
            Status.Text = "Connected"
            Me.Refresh()
            Me.Led2.Refresh()
        End If
    End Sub

    Private Sub WriteButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WriteButton.Click
        Dim boolLedStatus = True
        Dim intRefreshCounter = 0
        Dim intFilePointer As Integer = 0
        Dim intLineEndPointer As Integer = 0
        Dim strCommandString As String = ""
        Dim strChannelBuffer As String = ""
        If boolSourceFileImportedOK And boolDeviceFound Then
            Led2.Color = LED.LEDColorSelection.LED_Red
            Status.Text = "Writing"
            Me.Status.Refresh()
            For Each bByte As Char In strWriteFileFormat
                If bByte = ":" Then
                    If strCommandString <> "" Then
                        Call mySerialLink.SendDataToSerial(strCommandString)
                        Call mySerialLink.GetResponse(strChannelBuffer, "OK")
                        If (InStr(1, strChannelBuffer, "ERROR")) Then
                            MsgBox("Write error to device - Please erase device first!", MessageBoxButtons.OK, "ArdPigProgHost: Write Error")
                            Led2.Blink = False
                            Led2.Color = LED.LEDColorSelection.LED_Green
                            Status.Text = "Connected"
                            Exit Sub
                        End If
                        intRefreshCounter += 1
                        If intRefreshCounter = 4 Then
                            Status.Text = "Writing address " & Mid(strCommandString, 7, 4)
                            Led2.State = boolLedStatus
                            boolLedStatus = Not boolLedStatus
                            Me.Status.Refresh()
                            Me.Led2.Refresh()
                            intRefreshCounter = 0
                        End If
                    End If
                    strCommandString = "WRITE "
                Else
                    strCommandString = strCommandString & bByte
                End If
            Next
            Call mySerialLink.SendDataToSerial(strCommandString) 'flush out last string
            Call mySerialLink.GetResponse(strChannelBuffer, "OK")
            'update window after programming 
            Call UpdateMemoryWindows("Writing address ", "Writing EEPROM")
            Call mySerialLink.SendDataToSerial("DEVICE" & vbCrLf)
            Call mySerialLink.GetResponse(strChannelBuffer, ".")
            If ((strChannelBuffer <> "TimeOut") And Not (InStr(1, strChannelBuffer, "ERROR"))) Then 'catch ERROR conditions
                'parse string for information received
                Call ParseForArgument(strChannelBuffer, lConfigWord.Text, "ConfigWord")
            End If
            Led2.Blink = False
            Led2.Color = LED.LEDColorSelection.LED_Green
            Status.Text = "Connected"
        Else
            MsgBox("No device present or no Source File!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Private Function strChecksum(ByRef intChecksum As Integer) As String
        Dim tmpString As String
        Dim bByte As Byte
        bByte = ((Not intChecksum) + 1) And &HFF
        tmpString = bByte.ToString("X2")
        If tmpString.Length > 2 Then
            tmpString = tmpString.Substring(tmpString.Length - 2, 2)
        End If
        Return tmpString
    End Function
    Private Sub ComputeChecksum(ByRef strByte As String, ByVal bByte As Char, ByRef intChecksum As Integer)
        strByte = strByte & bByte
        If strByte.Length = 2 Then
            intChecksum = intChecksum + Val("&H" & strByte)
            strByte = ""
        End If
    End Sub
    Private Sub ExportButton_Click(sender As System.Object, e As System.EventArgs) Handles ExportButton.Click
        Dim strExportIntelHex As String = ":"
        Dim intLineByteCounter As Integer = 0
        Dim intLineAddressCounter As Integer = 0
        Dim intChecksum As Integer = 0
        Dim intNibbleCount As Integer = 0
        Dim strByte As String = ""
        Dim strBaseAddress = ""
        Dim intBaseAddress As Integer = 0
        Dim boolSecondline As Boolean = True
        Dim intInByteCount = 3
        Dim strInByte As String = ""
        Dim tmpbByte As Char
        Dim intAddressOffset = 0
        For Each bByte As Char In HexFileDump.Text
            If IsNumeric(bByte) Or (bByte >= "A" And bByte <= "F") Then
                If intLineAddressCounter = 0 Then
                    intAddressOffset = 0
                    strExportIntelHex = strExportIntelHex & "10" 'add byte count
                    intChecksum = 16 'set inital checksum
                    strBaseAddress = bByte
                    intLineAddressCounter += 1
                ElseIf intLineAddressCounter < 4 Then
                    intLineAddressCounter += 1
                    strBaseAddress = strBaseAddress & bByte
                ElseIf intLineByteCounter = 0 Then 'set record type
                    intBaseAddress = (Val("&H" & strBaseAddress) << 1) + intAddressOffset
                    intChecksum = intChecksum + (intBaseAddress >> 8) + (intBaseAddress And &HFF)
                    strExportIntelHex = strExportIntelHex + Decimal2Hex(intBaseAddress) 'write address
                    strExportIntelHex = strExportIntelHex & "00"
                    strByte = ""
                    ComputeChecksum(strByte, bByte, intChecksum)
                    strInByte = bByte
                    intInByteCount = 3
                    intLineByteCounter += 1
                ElseIf intLineByteCounter < 31 Then
                    If intInByteCount = 3 Then
                        strInByte = strInByte & bByte
                        intInByteCount = 0
                    ElseIf intInByteCount = 0 Then
                        tmpbByte = bByte
                        intInByteCount = 1
                    ElseIf intInByteCount = 1 Then
                        strExportIntelHex = strExportIntelHex & tmpbByte & bByte & strInByte
                        intInByteCount = 2
                    ElseIf intInByteCount = 2 Then
                        strInByte = bByte
                        intInByteCount = 3
                    End If
                    ComputeChecksum(strByte, bByte, intChecksum)
                    intLineByteCounter += 1
                ElseIf intLineByteCounter = 31 Then
                    strExportIntelHex = strExportIntelHex & tmpbByte & bByte & strInByte
                    ComputeChecksum(strByte, bByte, intChecksum)
                    strExportIntelHex = strExportIntelHex & strChecksum(intChecksum) & vbCrLf
                    strExportIntelHex = strExportIntelHex & ":"
                    If boolSecondline Then
                        intAddressOffset = 16
                        strExportIntelHex = strExportIntelHex & "10"
                        boolSecondline = False
                        intLineAddressCounter = 4
                        intChecksum = 16 'set inital checksum
                    Else
                        intLineAddressCounter = 0
                        boolSecondline = True
                    End If
                    intLineByteCounter = 0
                End If
            End If
        Next
        ' add CONFIG-Word
        strExportIntelHex = strExportIntelHex & "02400E00"
        intChecksum = 80
        strExportIntelHex = strExportIntelHex & lConfigWord.Text(2) & lConfigWord.Text(3) & lConfigWord.Text(0) & lConfigWord.Text(1)
        strByte = lConfigWord.Text(2)
        ComputeChecksum(strByte, lConfigWord.Text(3), intChecksum)
        strByte = lConfigWord.Text(0)
        ComputeChecksum(strByte, lConfigWord.Text(1), intChecksum)
        strExportIntelHex = strExportIntelHex & strChecksum(intChecksum) & vbCrLf & ":"
        ' add data range
        intLineAddressCounter = 0
        boolSecondline = True
        intLineByteCounter = 0
        For Each bByte As Char In EEPROMDump.Text
            If IsNumeric(bByte) Or (bByte >= "A" And bByte <= "F") Then
                If intLineAddressCounter = 0 Then
                    intAddressOffset = 0
                    strExportIntelHex = strExportIntelHex & "10" 'add byte count
                    intChecksum = 16 'set inital checksum
                    strBaseAddress = bByte
                    intLineAddressCounter += 1
                ElseIf intLineAddressCounter < 4 Then
                    intLineAddressCounter += 1
                    strBaseAddress = strBaseAddress & bByte
                ElseIf intLineByteCounter = 0 Then 'set record type
                    intBaseAddress = (Val("&H" & strBaseAddress) << 1) + intAddressOffset
                    intChecksum = intChecksum + (intBaseAddress >> 8) + (intBaseAddress And &HFF)
                    strExportIntelHex = strExportIntelHex + Decimal2Hex(intBaseAddress) 'write address
                    strExportIntelHex = strExportIntelHex & "00"
                    tmpbByte = bByte
                    intInByteCount = 3
                    strByte = ""
                    ComputeChecksum(strByte, bByte, intChecksum)
                    intLineByteCounter += 1
                ElseIf intLineByteCounter < 15 Then
                    If intInByteCount = 3 Then
                        strExportIntelHex = strExportIntelHex & tmpbByte & bByte & "00"
                        intInByteCount = 2
                    ElseIf intInByteCount = 2 Then
                        tmpbByte = bByte
                        intInByteCount = 3
                    End If
                    ComputeChecksum(strByte, bByte, intChecksum)
                    intLineByteCounter += 1
                ElseIf intLineByteCounter = 15 Then
                    strExportIntelHex = strExportIntelHex & tmpbByte & bByte & "00"
                    ComputeChecksum(strByte, bByte, intChecksum)
                    strExportIntelHex = strExportIntelHex & strChecksum(intChecksum) & vbCrLf
                    strExportIntelHex = strExportIntelHex & ":"
                    If boolSecondline Then
                        intAddressOffset = 16
                        strExportIntelHex = strExportIntelHex & "10"
                        boolSecondline = False
                        intLineAddressCounter = 4
                        intChecksum = 16 'set inital checksum
                    Else
                        intLineAddressCounter = 0
                        boolSecondline = True
                    End If
                    intLineByteCounter = 0
                End If
            End If
        Next
        strExportIntelHex = strExportIntelHex & "00000001FF"
        ' output to file
        With dlgSaveFile
            .Filter = "Hexfiles (*.hex)|*.hex|All files (*.*)|*.*"
            .FilterIndex = 1
            .FileName = ""

            Dim dlgResult As DialogResult = .ShowDialog()
            If DialogResult = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If
        End With

        Try
            HexFileName.Text = dlgSaveFile.FileName
            My.Computer.FileSystem.WriteAllText(dlgSaveFile.FileName, strExportIntelHex, False)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cannot create file.", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub
End Class

