Option Strict Off
Option Explicit On
Imports System.IO.Ports
Public Class SerialLink
    
    ' This class is designed to interface to a PiKoder/COM through an available serial
    ' face (COMn). COMn might be an actual RS232 connection - usually available on an 
    ' older PC or a virtual COM port provided by a USB adaptor.
    '
    Public Event SerialLinkEstablished()
    Public Event SerialLinkLost()
    '
    ' The class would be involved by EstablishSerialLink(). From this point onwards an 
    ' existing connection would be monitored and an event would be generated once the
    ' connection is lost (SerialLinkLost).
    ' 
    ' The class provides for specialized methods to read and write PiKoder/COM information
    ' such as GetPulseLength() and SentPulseLength(). Please refer to the definitons for more
    ' details.
    '
    ' The following property settings are supported (read/write at both design time and run time)
    ' The properties are:
    ' 
    ' Written by: Gregor Schlechtriem, www.pikoder.com
    '
    ' You may incorporate all or parts of this program into your own project.
    ' Keep in mind that you got this code free and as such Carousel Design Solutions
    ' has no responsibility to you, your customers, or anyone else regarding the use
    ' of this program code. No support or warrenty of any kind can be provided. All
    ' liability falls upon the user of this code and thus you must judge the suitability
    ' of this code to your application and usage.
    '
    Private mySerialPort As New SerialPort
    Private Connected As Boolean = False ' connection status

    Public Function SerialLinkConnected() As Boolean
        Return Connected
    End Function
    Public Function EstablishSerialLink(ByVal SelectedPort As String) As Boolean
        If (mySerialPort.PortName = SelectedPort) And Connected Then Return True ' port is already open
        If Connected Then mySerialPort.Close() ' another port has been selected
        Try
            mySerialPort.PortName = SelectedPort
            With mySerialPort
                .BaudRate = 9600
                .DataBits = 8
                .Parity = Parity.None
                .StopBits = StopBits.One
                .Handshake = Handshake.None
            End With
            mySerialPort.Open()
            Connected = True
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Connected = False
            Return False
        End Try
    End Function
    Public Sub GetResponse(ByRef SerialInputString As String, ByVal EndToken As String)
        Dim j As Integer = 0
        Dim myByte As Char = " " 'init for correct loop start
        SerialInputString = ""
        'read complete response into buffer
        Do
            If Connected And (mySerialPort.BytesToRead > 0) Then
                For i = 1 To mySerialPort.BytesToRead
                    SerialInputString = SerialInputString + Chr(mySerialPort.ReadByte)
                Next
                If (InStr(1, SerialInputString, EndToken)) Then
                    Exit Do
                ElseIf (InStr(1, SerialInputString, "ERROR")) Then
                    Exit Do
                Else : j = 0 'reset time out counter to keep reading as long as bytes are sent
                End If
            End If
            'check for TimeOut
            j = j + 1
            If j > 100 Then
                SerialInputString = "TimeOut" : Exit Do ' timeout exit
            End If
            System.Threading.Thread.Sleep(20) 'wait for next frame and allow for task switch
        Loop
    End Sub
    Public Sub SendDataToSerial(ByVal strWriteBuffer As String)
        Try
            mySerialPort.Write(strWriteBuffer, 0, Len(strWriteBuffer))
        Catch ex As SystemException
            Connected = False
            MsgBox("Lost Connection. Program will stop.", MsgBoxStyle.OkOnly, "Error Message")
            End
        End Try
    End Sub
    Public Sub MyForm_Dispose()
        Try
            mySerialPort.Close()
        Catch ex As Exception
        End Try
    End Sub

End Class
