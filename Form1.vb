'This sample code is for illustration only, without warranty either expressed or implied, including, but not limited to,
'the implied warranties of merchantability and/or fitness for a particular purpose. This sample code assumes that you are familiar
'with the programming language being demonstrated and the tools used to create and debug procedures. 

Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading

Public Class Form1

    Private connection As Socket            ' Socket for accepting a connection      
    Private readThread As Thread            ' Thread for processing incoming TCP messages
    Private SocketStream As NetworkStream   ' Socketstream
    Private TCPListener As TcpListener      ' Listener
    Private BytesToRead() As Byte           ' Network Stream Incoming Bytes (Incoming XML)
    Private BytesToSend() As Byte           ' Network Stream Outgoing Bytes ( Reply XML)
    Private NumBytesRead As Integer         ' Number of Bytes Read
    Private port As String = "51108"
    Private totalClients As Integer

    ' Global Constants
    Const sSuccess As String = "200"
    Const sFailure As String = "300"

    ' Global variables
    Dim sMsg_RcvdAll As String = ""         ' Parsed XML String with VbCrLf Characters
    Dim iRcvdMsgLength As Integer           ' Length of String for Do Loop
    Dim sRcvdmsg As String = ""             ' Parsed XML String without VbCrLf Characters
    Dim LocalHostName As String

    '3.7.23 KaryWa:
    Dim TCPClientList As New List(Of TcpClient)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Start TCP Server on its own thread - this keeps the UI from being blocked
        Me.Show()
        readThread = New Thread(New ThreadStart(AddressOf RunTcpServer))
        readThread.Start()
    End Sub


    ' Allows updating the textbox in the GUI from a different thread
    Sub SendUIMessage(message As String)
        StatusTextBox.Invoke(Sub()
                                 StatusTextBox.Text += message + System.Environment.NewLine
                                 Me.Refresh()
                             End Sub)
    End Sub

    Sub RunTcpServer()

        ' wait for a client connection and display the text that the client sends
        Try

            'Create the listener for all IP addresses
            TCPListener = New TcpListener(IPAddress.Any, port)

            'TcpListener waits for connection request
            TCPListener.Start()

            'Establish connection upon client request
            While True

                ' Accept the incoming connection then create a new thread to handle this client
                ' We need 1 thread per connection/client - we wait at this line until a client connects
                Dim TCPClient As TcpClient = TCPListener.AcceptTcpClient()

                ' A client connection was just accepted, create a thread for it and sent to HandleClient

                ' Add it to the client list first. 3.7.23 KaryWa: this line is not implemented I was just evaluating ideas
                TCPClientList.Add(TCPClient) '

                totalClients += 1
                Debug.Print("Client connected!")
                SendUIMessage("Client Connected! Total clients: " + totalClients.ToString())
                Dim clientThread As New Threading.Thread(Sub() HandleClient(TCPClient))
                clientThread.Start()

            End While
        Catch ex As Exception
            Dim ExceptionMsg As String
            ExceptionMsg = ex.ToString()
            My.Computer.FileSystem.WriteAllText("TCPServerLog_ERR.txt", ExceptionMsg & vbCrLf, True)
        End Try
    End Sub


    Sub HandleClient(Client As TcpClient)

        'Read string data sent from client
        Do
            Try
                ' Create NetworkStream object associated with socket
                SocketStream = Client.GetStream
                Dim BytesToRead(Client.ReceiveBufferSize) As Byte

                ' Read the Socket Stream data buffer string sent to the server
                NumBytesRead = SocketStream.Read(BytesToRead, 0,
                   CInt(Client.ReceiveBufferSize))

                sMsg_RcvdAll = Encoding.ASCII.GetString(BytesToRead, 0, NumBytesRead)

                ' Return XML Reply  
                BytesToSend = Encoding.ASCII.GetBytes("ACK")
                SocketStream.Write(BytesToSend, 0, BytesToSend.Length)

            Catch ex As Exception
                ' Handle exception if error reading data
                Exit Do
            End Try

        Loop While Client.Connected

        ' Client has disconnected, decrement totalClients and update UI
        '3.7.23 karywa: Possibly remove client from the TCPClientList here?
        totalClients -= 1
        SendUIMessage("Client disconnected: " + totalClients.ToString() + " clients connected.")

        ' Close connection  
        SocketStream.Close()
        Client.Close()
    End Sub

    ' This replaces the previous GetHostName()/IPHostEntry Code.
    ' I refactored to grab the first IPV4 address because when I was using GetHostEntry early on, the first
    ' IP in the list was IPV6 which I couldn't use. So, I created this to grab the first IPV4
    ' I ended up just using IPAddress.Any in line 57 which negates the need for this function.
    ' I'm only leaving it in in the event it is needed later and can be refactored to be selective as to which IP is used.

    Function GetIPV4Address() As String

        Dim sHostName As String                               ' Local Host Name
        Dim sHostIPAddress As String = String.Empty           ' Local IP Address

        sHostName = Dns.GetHostName()
        Dim ipEnter As IPHostEntry = Dns.GetHostEntry(sHostName)
        Dim IpAdd() As IPAddress = ipEnter.AddressList

        ' Grab the first IPV4 address in the list
        ' Otherwise it may grab an IPV6 address unless that's what we want.
        For Each IP As IPAddress In IpAdd
            If IP.AddressFamily = AddressFamily.InterNetwork Then
                sHostIPAddress = IP.ToString
                Debug.Print("Using IP:{0} Port:{1} ", sHostIPAddress)
                'Exit the loop when we find the first IPV4 address
                Exit For
            End If
        Next

        Return sHostIPAddress
    End Function

End Class
