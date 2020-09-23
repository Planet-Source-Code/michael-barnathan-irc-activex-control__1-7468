VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.UserControl IRC 
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   MaskPicture     =   "irc.ctx":0000
   Palette         =   "irc.ctx":0442
   PropertyPages   =   "irc.ctx":1084
   ScaleHeight     =   3945
   ScaleWidth      =   4560
   ToolboxBitmap   =   "irc.ctx":1097
   Begin VB.TextBox RichTextBox1 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   4096
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   -1  'True
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   6667
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
End
Attribute VB_Name = "IRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
'Const m_def_LocalAddress = ""
'Const m_def_LocalName = ""
'Const m_def_State = 0
'Const m_def_About = 0
Const m_def_Channel = "#IRCChat"
Const m_def_Nick = "IRCUser"
Const m_def_Hostmask = "IRCClient"
'Const m_def_Blocking= 0
'Const m_def_Nick = "IRCUser"
'Const m_def_Hostmask = "IR CClient"
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 1
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Servername = "irc.id-net.fr"
Const m_def_AddressFamily = 2
Const m_def_Protocol = 0
Const m_def_SocketType = 1
'Const m_def_Status = 0
'Const m_def_Port = 6667
'Const m_def_Channel = "#IRCChat"
Const m_def_Timeout = 10
'Property Variables:
Dim m_LocalAddress As String
Dim m_LocalName As String
Dim m_State As Integer
'Dim m_About As Variant
Dim m_Channel As String
Dim m_Nick As String
Dim m_Hostmask As String
'Dim m_Blocking As Boolean
'Dim m_Nick As String
'Dim m_Hostmask As String
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'Dim m_Servername As String
'Dim m_Status As Variant
'Dim m_Port As Variant
'Dim m_Channel As String
Dim m_Timeout As Variant
'Event Declarations:
Event Connected()
Event Message(UserMessage As String, FromUser As String, Direct)
'Event Connected()
Event Error()
Event Disconnected()
Event RecievedData(DataRecieved As String)
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event UserJoined(UserName As String)
Attribute UserJoined.VB_Description = "Occurs when a user joins the channel"
Event UserPart(UserName As String)
Attribute UserPart.VB_Description = "Occurs when a user parts from the channel"
Event UserQuit(UserName As String)
Attribute UserQuit.VB_Description = "Occurs when a user quits the server"
'Event Message(UserMessage As String, FromUser As String, Direct As Boolean)
Event UserAction(ActionName As String)
'Event Connected()
'Event Disconnected()
'Event StatusChange(Status As Integer)
Event TimedOut()



Private Sub Text1_Change()

End Sub

Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    RichTextBox1.Enabled = New_Enabled
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub


Public Function Join() As Variant
Attribute Join.VB_Description = "Joins the channel specified in the Channel property"
Socket1.SendLen = Len("JOIN " & Me.Channel & vbCrLf)
Socket1.SendData = "JOIN " & Me.Channel & vbCrLf
DoEvents
Socket1.SendLen = Len("NAMES " & Me.Channel & vbCrLf)
Socket1.SendData = "NAMES " & Me.Channel & vbCrLf
End Function

Public Function Part() As Variant
Attribute Part.VB_Description = "Parts from a channel"
Socket1.SendLen = Len("PART " & Me.Channel & vbCrLf)
Socket1.SendData = "PART " & Me.Channel & vbCrLf
End Function

Public Function Quit() As Variant
Attribute Quit.VB_Description = "Sends a Quit command to the IRC server. Does not actually disconnect from the server"
Socket1.SendLen = Len("QUIT" & vbCrLf)
Socket1.SendData = "QUIT" & vbCrLf
End Function

Public Property Get Timeout() As Variant
Attribute Timeout.VB_Description = "Sets the timeout(In seconds)"
    Timeout = m_Timeout
End Property

Public Property Let Timeout(ByVal New_Timeout As Variant)
    m_Timeout = New_Timeout
    PropertyChanged "Timeout"
End Property

Public Function Action(ActionName As String) As String
Attribute Action.VB_Description = "Equivalent to the /me IRC command."
Socket1.SendLen = Len(ActionName & vbCrLf)
Socket1.SendData = ActionName & vbCrLf
End Function

Public Function Msg(UserToSend As String, MessageToSend As String) As Variant
Attribute Msg.VB_Description = "Sends a message to a user"
Socket1.SendLen = Len("PRIVMSG " & UserToSend & " :" & MessageToSend & vbCrLf)
Socket1.SendData = "PRIVMSG " & UserToSend & " :" & MessageToSend & vbCrLf
O = Len(RichTextBox1.Text)
RichTextBox1.Text = RichTextBox1.Text & Me.Nick & ": "
DoEvents
RichTextBox1.Text = RichTextBox1.Text & MessageToSend & vbCrLf
End Function

Private Sub Socket1_Connect()
RaiseEvent Connected
End Sub

Private Sub Socket1_Disconnect()
RaiseEvent Disconnected
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
Socket1.RecvLen = DataLength
A$ = Socket1.RecvData
RaiseEvent RecievedData(A$)

'If there's a message...

M% = InStr(1, A$, "PRIVMSG")
If M% <> 0 Then
H% = InStr(1, A$, "!")
J% = 1
Do Until J% = 0
J% = InStr(J% + 1, A$, ":")
If J% <> 0 Then K% = J%
Loop
If InStr(1, A$, Me.Nick) > 0 Then DirectMode = True Else: DirectMode = False
RaiseEvent Message(Mid$(A$, 2, H% - 2), Mid$(A$, K% + 1, Len(A$) - K% - 2), DirectMode)
O = Len(RichTextBox1.Text)
RichTextBox1.Text = RichTextBox1.Text + Mid$(A$, 2, H% - 2) & ": "
DoEvents
DoEvents
RichTextBox1.Text = RichTextBox1.Text + Mid$(A$, K% + 1, Len(A$) - K% - 2) & vbCrLf
DoEvents
Exit Sub 'Don't let anything else fire
'So this way they can't say JOIN in a message
End If
'For a join
If InStr(1, A$, "JOIN") > 0 Then
H% = InStr(1, A$, "!")
J% = 1
Do Until J% = 0
J% = InStr(J% + 1, A$, ":")
If J% <> 0 Then K% = J%
Loop
RaiseEvent UserJoined(Trim$(Mid$(A$, 2, H% - 2)))
End If
'For a part
If InStr(1, A$, "PART") > 0 Then
H% = InStr(1, A$, "!")
J% = 1
Do Until J% = 0
J% = InStr(J% + 1, A$, ":")
If J% <> 0 Then K% = J%
Loop
RaiseEvent UserPart(Trim$(Mid$(A$, 2, H% - 2)))
End If
'For a quit
If InStr(1, A$, "QUIT") > 0 Then
H% = InStr(1, A$, "!")
J% = 1
Do Until J% = 0
J% = InStr(J% + 1, A$, ":")
If J% <> 0 Then K% = J%
Loop
RaiseEvent UserQuit(Trim$(Mid$(A$, 2, H% - 2)))
End If

'Now parse the names list here

R% = InStr(1, LCase$(A$), "= " & LCase$(Me.Channel) & " :")
If R% <> 0 Then
Dim OO%, OP%
If InStr(1, A$, "@") <> 0 Then A$ = Left$(A$, InStr(1, A$, "@") - 1) & Mid$(A$, InStr(1, A$, "@") + 1, Len(A$) - InStr(1, A$, "@"))
OO% = R% + 13 - (9 - Len(Me.Channel))
'If Mid$(Text1.Tag, O%, 1) = "@" Then O% = O% + 1
OP% = OO%
For L = OP% To DataLength
If Mid$(A$, L, 1) = " " Then AL$ = Trim$(Mid$(A$, OO%, L - OO%)): OO% = L
If AL$ <> "" And AL$ <> Me.Nick Then RaiseEvent UserJoined(AL$)
If AL$ <> "" Then AL$ = ""
Next
End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Socket1.AddressFamily = AF_INET
    Socket1.Protocol = IPPROTO_IP
    Socket1.SocketType = SOCK_STREAM
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_Servername = "irc.id-net.fr"
    m_Port = m_def_Port
'   m_Channel = m_def_Channel
    m_Timeout = m_def_Timeout
'    m_Nick = m_def_Nick
'    m_Hostmask = m_def_Hostmask
    m_Blocking = m_def_Blocking
    m_Channel = m_def_Channel
    m_Nick = m_def_Nick
    m_Hostmask = m_def_Hostmask
    m_LocalAddress = m_def_LocalAddress
    m_LocalName = m_def_LocalName
    m_State = m_def_State
    End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Servername = PropBag.ReadProperty("Servername", m_def_Servername)
    m_Port = PropBag.ReadProperty("Port", m_def_Port)
    m_Timeout = PropBag.ReadProperty("Timeout", m_def_Timeout)
    Socket1.AddressFamily = PropBag.ReadProperty("AddressFamily", 2)
    m_Blocking = PropBag.ReadProperty("Blocking", m_def_Blocking)
    Socket1.SocketType = PropBag.ReadProperty("SocketType", 1)
    Socket1.Protocol = PropBag.ReadProperty("Protocol", 0)
'    Socket1.LocalAddress = PropBag.ReadProperty("LocalAddress", "")
'    Socket1.LocalName = PropBag.ReadProperty("LocalName", "")
'    Socket1.State = PropBag.ReadProperty("State", 0)
    Socket1.RemotePort = PropBag.ReadProperty("Port", 0)
    Socket1.HostName = PropBag.ReadProperty("Servername", "")
    m_Channel = PropBag.ReadProperty("Channel", m_def_Channel)
    m_Nick = PropBag.ReadProperty("Nick", m_def_Nick)
    m_Hostmask = PropBag.ReadProperty("Hostmask", m_def_Hostmask)
    Socket1.Blocking = PropBag.ReadProperty("Blocking", False)
    m_LocalAddress = PropBag.ReadProperty("LocalAddress", m_def_LocalAddress)
    m_LocalName = PropBag.ReadProperty("LocalName", m_def_LocalName)
    m_State = PropBag.ReadProperty("State", m_def_State)
End Sub

Private Sub UserControl_Resize()
RichTextBox1.Width = UserControl.Width - 225
RichTextBox1.Height = UserControl.Height - 225
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Port", m_Port, m_def_Port)
    Call PropBag.WriteProperty("Timeout", m_Timeout, m_def_Timeout)
    Call PropBag.WriteProperty("AddressFamily", Socket1.AddressFamily, 2)
    Call PropBag.WriteProperty("SocketType", Socket1.SocketType, 1)
    Call PropBag.WriteProperty("Protocol", Socket1.Protocol, 0)
    Call PropBag.WriteProperty("LocalAddress", Socket1.LocalAddress, "")
    Call PropBag.WriteProperty("LocalName", Socket1.LocalName, "")
    Call PropBag.WriteProperty("State", Socket1.State, 0)
    Call PropBag.WriteProperty("Servername", Socket1.HostName, "")
    Call PropBag.WriteProperty("Channel", m_Channel, m_def_Channel)
    Call PropBag.WriteProperty("Nick", m_Nick, m_def_Nick)
    Call PropBag.WriteProperty("Hostmask", m_Hostmask, m_def_Hostmask)
    Call PropBag.WriteProperty("Blocking", Socket1.Blocking, False)
    Call PropBag.WriteProperty("LocalAddress", m_LocalAddress, m_def_LocalAddress)
    Call PropBag.WriteProperty("LocalName", m_LocalName, m_def_LocalName)
    Call PropBag.WriteProperty("State", m_State, m_def_State)
End Sub

Public Property Let Hostmask(ByVal New_Hostmask As String)
    If Socket1.Connected = True Then MsgBox "Cannot be set once the control is connected", 16, "Error setting property": GoTo 2
    m_Hostmask = New_Hostmask
    PropertyChanged "Hostmask"
2     'Exits the sub
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Socket1,Socket1,-1,AddressFamily
Public Property Get AddressFamily() As Integer
    AddressFamily = Socket1.AddressFamily
End Property

Public Property Let AddressFamily(ByVal New_AddressFamily As Integer)
    Socket1.AddressFamily() = New_AddressFamily
    PropertyChanged "AddressFamily"
End Property
'
'Public Property Get Blocking() As Boolean
'    Blocking = m_Blocking
'End Property
'
'Public Property Let Blocking(ByVal New_Blocking As Boolean)
'    m_Blocking = New_Blocking
'    PropertyChanged "Blocking"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Socket1,Socket1,-1,SocketType
Public Property Get SocketType() As Integer
    SocketType = Socket1.SocketType
End Property

Public Property Let SocketType(ByVal New_SocketType As Integer)
    Socket1.SocketType() = New_SocketType
    PropertyChanged "SocketType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Socket1,Socket1,-1,Protocol
Public Property Get Protocol() As Integer
    Protocol = Socket1.Protocol
End Property

Public Property Let Protocol(ByVal New_Protocol As Integer)
    Socket1.Protocol() = New_Protocol
    PropertyChanged "Protocol"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Socket1,Socket1,-1,LocalAddress
'Public Property Get LocalAddress() As String
'    LocalAddress = Socket1.LocalAddress
'End Property
'
'Public Property Let LocalAddress(ByVal New_LocalAddress As String)
'    Socket1.LocalAddress() = New_LocalAddress
'    PropertyChanged "LocalAddress"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Socket1,Socket1,-1,LocalName
'Public Property Get LocalName() As String
'    LocalName = Socket1.LocalName
'End Property
'
'Public Property Let LocalName(ByVal New_LocalName As String)
'    Socket1.LocalName() = New_LocalName
'    PropertyChanged "LocalName"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Socket1,Socket1,-1,State
'Public Property Get State() As Integer
'    State = Socket1.State
'End Property
'
'Public Property Let State(ByVal New_State As Integer)
'    Socket1.State() = New_State
'    PropertyChanged "State"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Socket1,Socket1,-1,Abort
Public Function Abort() As Integer
Attribute Abort.VB_Description = "Terminates any socket operations and disconnects the socket"
    Abort = Socket1.Abort()
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Socket1,Socket1,-1,Disconnect
Public Function Disconnect() As Integer
Attribute Disconnect.VB_Description = "Disconnects from the server."
    Disconnect = Socket1.Disconnect()
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Socket1,Socket1,-1,RemotePort
Public Property Get Port() As Integer
    Port = Socket1.RemotePort
End Property

Public Property Let Port(ByVal New_Port As Integer)
    Socket1.RemotePort() = New_Port
    PropertyChanged "Port"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Socket1,Socket1,-1,HostName
Public Property Get Servername() As String
    Servername = Socket1.HostName
End Property

Public Property Let Servername(ByVal New_Servername As String)
    Socket1.HostName() = New_Servername
    PropertyChanged "Servername"
End Property

Public Property Get Channel() As String
Attribute Channel.VB_Description = "Returns or sets the channel to connect to"
    Channel = m_Channel
End Property

Public Property Let Channel(ByVal New_Channel As String)
    m_Channel = New_Channel
    PropertyChanged "Channel"
End Property

Public Property Get Nick() As String
Attribute Nick.VB_Description = "Sets or returns the nickname to use when logging on to the server. Default is IRCUSER"
    Nick = m_Nick
End Property

Public Property Let Nick(ByVal New_Nick As String)
    If Socket1.Connected = True Then Socket1.SendLen = Len("NICK " & New_Nick & vbCrLf): Socket1.SendData = "NICK " & New_Nick & vbCrLf
    m_Nick = New_Nick
    PropertyChanged "Nick"
End Property

Public Property Get Hostmask() As String
Attribute Hostmask.VB_Description = "Sets or returns the hostmask when logging in. Default is IRCClient"
    Hostmask = m_Hostmask
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Socket1,Socket1,-1,Blocking
Public Property Get Blocking() As Boolean
    Blocking = Socket1.Blocking
End Property

Public Property Let Blocking(ByVal New_Blocking As Boolean)
    Socket1.Blocking() = New_Blocking
    PropertyChanged "Blocking"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Socket1,Socket1,-1,Connect
'Public Function Connect() As Integer
'    Connect = Socket1.Connect()
'End Function

Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Displays the about screen"
frmAbout.Show 1
End Sub

Public Function Connect() As Integer
Attribute Connect.VB_Description = "Connects to the server specified in HostName"
Socket1.Connect
R& = Timer + 10
Do Until Timer > R& Or Socket1.Connected = True: DoEvents: Loop
If Timer > R& Then Socket1.Abort: MsgBox "Connection timed out", 16, "Error connecting": GoTo 1
Socket1.SendLen = Len("USER " & Me.Hostmask & " a a a a" & vbCrLf)
Socket1.SendData = "USER " & Me.Hostmask & " a a a a" & vbCrLf
Do Until Socket1.State = 1: DoEvents: Loop
Socket1.SendLen = Len("NICK " & Me.Nick & vbCrLf)
Socket1.SendData = "NICK " & Me.Nick & vbCrLf
1 'End of function
End Function

Public Property Get LocalAddress() As String
    LocalAddress = Socket1.LocalAddress
End Property

Public Property Let LocalAddress(ByVal New_LocalAddress As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_LocalAddress = New_LocalAddress
    PropertyChanged "LocalAddress"
End Property

Public Property Get LocalName() As String
    LocalName = Socket1.LocalName
End Property

Public Property Let LocalName(ByVal New_LocalName As String)
    If Ambient.UserMode = False Then Err.Raise 394
    If Ambient.UserMode Then Err.Raise 393
    m_LocalName = New_LocalName
    PropertyChanged "LocalName"
End Property

Public Property Get State() As Integer
Attribute State.VB_MemberFlags = "400"
    State = Socket1.State
End Property

Public Property Let State(ByVal New_State As Integer)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_State = New_State
    PropertyChanged "State"
End Property

Public Function AddText(TextToAdd As String, Color As Long) As Variant
Attribute AddText.VB_Description = "Adds text to the chat text box"

End Function

