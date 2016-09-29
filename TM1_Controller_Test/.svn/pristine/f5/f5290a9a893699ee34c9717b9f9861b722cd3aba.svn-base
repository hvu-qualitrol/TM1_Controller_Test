Imports System.Net
Imports System.Net.Sockets

Public Class NTPServer
    Public Shared Function GetNetworkTime() As Date

        ' default Windows time server
        Dim ntpServer As String = "time.windows.com"

        ' NTP message size - 16 bytes of the digest (RFC 2030)
        Dim ntpData(47) As Byte

        ' Setting the Leap Indicator, Version Number and Mode values
        ' LI = 0 (no warning), VN = 3 (IPv4 only), Mode = 3 (Client Mode)
        ntpData(0) = 27

        Dim addresses() As IPAddress = Dns.GetHostEntry(ntpServer).AddressList

        ' The UDP port number assigned to NTP is 123
        Dim ipEndPoint As IPEndPoint = New IPEndPoint(addresses(0), 123)

        ' NTP uses UDP
        Dim socket As Socket = New Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)

        socket.Connect(ipEndPoint)

        socket.Send(ntpData)
        socket.Receive(ntpData)
        socket.Close()

        ' Offset to get to the "Transmit Timestamp" field (time at which the reply 
        ' departed the server for the client, in 64-bit timestamp format."
        Dim serverReplyTime As Byte = 40

        ' Get the seconds part
        Dim intPart As ULong = BitConverter.ToUInt32(ntpData, serverReplyTime)

        ' Get the seconds fraction
        Dim fractPart As ULong = BitConverter.ToUInt32(ntpData, serverReplyTime + 4)

        'Convert From big-endian to little-endian
        intPart = SwapEndianness(intPart)
        fractPart = SwapEndianness(fractPart)

        Dim milliseconds As Int64 = (intPart * 1000) + ((fractPart * 1000) / 4294967296)

        ' **UTC** time
        Dim networkDateTime As DateTime = (New DateTime(1900, 1, 1)).AddMilliseconds(Convert.ToInt64(milliseconds))

        Return networkDateTime
    End Function


    ' stackoverflow.com/a/3294698/162671
    Public Shared Function SwapEndianness(ByVal x As ULong) As UInteger
        Dim result As UInteger = 0

        result += (x And 255) << 24
        result += (x And 65280) << 8
        result += (x And 16711680) >> 8
        result += (x And 4278190080) >> 24

        Return result
    End Function
End Class

