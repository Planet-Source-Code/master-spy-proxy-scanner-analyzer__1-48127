VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Proxy Scanner 2K3 - FREE!"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   5200
      Left            =   6480
      Top             =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Compile Proxy info page"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan this one"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Scan List"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3240
      TabIndex        =   15
      Top             =   1320
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6240
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select a list of Proxies"
      Filter          =   "*.txt"
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   6240
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   5640
      Width           =   6735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6240
      Top             =   240
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5530
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Proxy"
         Object.Width           =   5027
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Port"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Response"
         Object.Width           =   1677
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Works"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Anonimity"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Text            =   "80"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "200.52.213.51"
      Top             =   720
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6360
      Top             =   0
   End
   Begin MSWinsockLib.Winsock ws1 
      Left            =   6360
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Text            =   "Workin on it..."
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Not found Yet!"
      Top             =   120
      Width           =   1695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5760
      TabIndex        =   21
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Proxy list:"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Not Connected"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      TabIndex        =   10
      Top             =   5280
      Width           =   6255
   End
   Begin VB.Label Label4 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Adress:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Your external IP:"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Your internal ip:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'*                       -                           *
'*                                                   *
'*                  PROXY SCANNER                    *
'*                by: Pedro Camanho                  *
'*      pedrocamanho@hotmail.com    ICQ:84851212     *
'*                                                   *
'*        Sorry, little comments! but feel           *
'*          free to use my prog for your             *
'*          stuff and to make it better!             *
'*                                                   *
'*                        -                          *
'*****************************************************
'This is not perfect or done! i just posted it here because this is VERY usefull
'and no one on PSC has done this, which is VERY NEEDED for hacking and all
'Don't email me with 'how do i hack', 'how do i use a proxy' or 'how do i spoof my ip'
'or anyother stupid question!

Option Explicit
Dim Litem As ListItem
Dim X As Integer
Dim Try As Integer
Dim ConType As String



Private Sub Command1_Click()
On Error GoTo errr
Command2.Enabled = False
Command5.Enabled = False
Command1.Enabled = False
ConType = "one"
X = 0
Try = 0
Text6 = ""
'here i connect to the proxy on the specific port!
ws1.Close
ws1.Connect Text3, Text4
Timer2.Enabled = True
'Here i create the header for which i will send to http://age.ne.jp/x/maxwell/cgi-bin/prxjdg.cgi
'or any other proxy check cgi, to retrieve the info
        strHeaders = "GET http://age.ne.jp/x/maxwell/cgi-bin/prxjdg.cgi" & " HTTP/1.0" & vbCrLf
        strHeaders = strHeaders & "Accept: */*" & vbCrLf
        strHeaders = strHeaders & "Accept-Language: en-us" & vbCrLf
        strHeaders = strHeaders & "Content-Encoding: gzip, deflate" & vbCrLf
        strHeaders = strHeaders & "Host: " & "www.uol.com.br" & vbCrLf
        strHeaders = strHeaders & "User-Agent: Mozilla 4.0 (Windows)" & vbCrLf
        strHeaders = strHeaders & "Proxy-Connection: Close" & vbCrLf & vbCrLf
Label5 = "Connecting..."
Exit Sub
errr:
MsgBox Err.Description, vbCritical, "error"

End Sub

Private Sub Command2_Click()
'Here i open a file, and save the retireved html code from the proxy cgi script
    Open App.Path & "\render.htm" For Output As #1
    Print #1, Text6.Text
    Close #1
Shell ("rundll32.exe url.dll,FileProtocolHandler " & App.Path & "\render.htm")
End Sub

Private Sub Command3_Click()
cd.ShowOpen
Text5 = cd.filename

End Sub

Private Sub Command4_Click()
'here i open um the list of proxies and put them in listbox1
File2ListBox Text5, List1
Command5.Enabled = True
End Sub

Private Sub Command5_Click()
On Error Resume Next
Command2.Enabled = False
Command5.Enabled = False
Command1.Enabled = False
Label7 = List1.ListCount
ConType = "list"
X = 0
Try = 0

List1.Selected(0) = True
If List1.Text = "" Then
Command2.Enabled = True
Command1.Enabled = True
Exit Sub
End If
Text6 = ""
ParseText List1.Text
List1.RemoveItem 0
ws1.Close
ws1.Connect Text3, Text4
Timer2.Enabled = True
'Here i create the header for which i will send to http://age.ne.jp/x/maxwell/cgi-bin/prxjdg.cgi
'or any other proxy check cgi, to retrieve the info
        strHeaders = "GET http://age.ne.jp/x/maxwell/cgi-bin/prxjdg.cgi" & " HTTP/1.0" & vbCrLf
        strHeaders = strHeaders & "Accept: */*" & vbCrLf
        strHeaders = strHeaders & "Accept-Language: en-us" & vbCrLf
        strHeaders = strHeaders & "Content-Encoding: gzip, deflate" & vbCrLf
        strHeaders = strHeaders & "Host: " & "www.uol.com.br" & vbCrLf
        strHeaders = strHeaders & "User-Agent: Mozilla 4.0 (Windows)" & vbCrLf
        strHeaders = strHeaders & "Proxy-Connection: Close" & vbCrLf & vbCrLf
Label5 = "Connecting..."
Exit Sub

End Sub

Private Sub Form_Load()

Text1 = ws1.LocalIP


End Sub

Public Sub File2ListBox(sFile As String, oList As ListBox)
'This is someones code for opening a file and puting it item by
'item into a listbox! very nice!
    Dim fnum As Integer
    Dim sTemp As String
    fnum = FreeFile()
    oList.Clear
    Open sFile For Input As fnum


    While Not EOF(fnum)
        Line Input #fnum, sTemp
        oList.AddItem sTemp
        Label7 = Label7 + 1
    Wend
    Close fnum
End Sub





Private Sub List1_Click()
If List1.Text <> "" Then
Command5.Enabled = True
Else
Command5.Enabled = False
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Here i sort by ping, i could have made this better and all but...
'i was lazy   =)   i or u still can though!
Module1.SortColumn ListView1, 3, sortNumeric, sortAscending
End Sub

Private Sub Text2_Change()
'if i recieve the 'Bandwidth exceeded error' then i tell the guy to try next month
If Left(Text2, 1) = "<" Then Text2 = "Try again next month"
End Sub

Private Sub Text2_DblClick()
'This is someguys code for retrieving your real ip
'It conects to a website and retireves your real ip!
'Very usefull for people who are in networks who have ips like 10.10.10.5
On Error Resume Next
    Dim MyIP As String
    Timer3.Enabled = True
    Text2 = "Workin on it..."
    MyIP = Inet1.OpenURL("http://pchelplive.com/ip.php")

    Text2 = MyIP

End Sub

Private Sub Text5_Change()
If Text5 = "" Then
Command4.Enabled = False
Else
Command4.Enabled = True
End If
End Sub

Private Sub Text6_Change()
Dim r As Integer
Dim z As Integer
If Text6 = "" Then Exit Sub
Try = Try + 1
If Try = 2 Then Exit Sub
Command2.Enabled = True
Timer2.Enabled = False


'If the webpage has loaded sucessfully, then continue to next step
If UCase(Left(Text6, 15)) = "HTTP/1.0 200 OK" Then
'Here i set the variables
'r and z are the character number in which the text i look for is located at
r = InStr(1, Text6.Text, "ProxyJudge V2.27", vbTextCompare)
z = InStr(1, Text6.Text, "AnonyLevel :", vbTextCompare)
Label5 = ws1.RemoteHost & ":" & ws1.RemotePort & "  WORKS, "
'Listview workings....  Add the ip number
Set Litem = ListView1.ListItems.Add(, , Text3)
'add the port
Litem.ListSubItems.Add , , Text4
'add the ping which is counted by a timer
Litem.ListSubItems.Add , , X
'reset the ping for the next one
X = 0
'-------------------- <- these just help me concentrate when im writing the code


'If r <>0 then it must be something, so if i find the words
''ProxyJudge V2.27' in the text, then the proxy must be working
'because it was able to navigate me to that page
'so here i say if the proxy is valid or not!
If r > 0 Then
Label5 = Label5 & "and IS valid!"
Litem.ListSubItems.Add , , "YES"
Else
Label5 = Label5 & "but is NOT valid!"
Litem.ListSubItems.Add , , "NO"
End If
'Here is the anon level
If Mid(Text6, z + 48, 1) < 5 Then
Litem.ListSubItems.Add , , Mid(Text6, z + 48, 1)
Else
Litem.ListSubItems.Add , , "Err. " & Mid(Text6, z + 48, 1)
End If
'--------------------
Else


'just header issues
'All over again but checking for http 1.1 now, not 1.0
'i wish i could have used the 'Or' function, but i cant ever get it to work!!
'also would be nice to do a http 0.9 one
If UCase(Left(Text6, 15)) = "HTTP/1.1 200 OK" Then
r = InStr(1, Text6.Text, "ProxyJudge V2.27", vbTextCompare)
z = InStr(1, Text6.Text, "AnonyLevel :", vbTextCompare)
Label5 = ws1.RemoteHost & ":" & ws1.RemotePort & "  WORKS, "
Set Litem = ListView1.ListItems.Add(, , Text3)
Litem.ListSubItems.Add , , Text4
Litem.ListSubItems.Add , , X
X = 0
'--------------------
If r > 0 Then
Label5 = Label5 & "and IS valid!"
Litem.ListSubItems.Add , , "YES"
Else
Label5 = Label5 & "but is NOT valid!"
Litem.ListSubItems.Add , , "NO"
End If
If Mid(Text6, z + 48, 1) < 5 Then
Litem.ListSubItems.Add , , Mid(Text6, z + 48, 1)
Else
Litem.ListSubItems.Add , , "Err. " & Mid(Text6, z + 48, 1)
End If
'--------------------
If Text6 = "" Then Label5 = "Not Connected"
Command5.Enabled = True
Command1.Enabled = True
If ConType = "list" Then Command5_Click
Exit Sub
End If
Label5 = ws1.RemoteHost & ":" & ws1.RemotePort & "  IS NOT a valid proxy server!"
Set Litem = ListView1.ListItems.Add(, , Text3)
Litem.ListSubItems.Add , , Text4
Litem.ListSubItems.Add , , X
Litem.ListSubItems.Add , , "NO"
Litem.ListSubItems.Add , , "-"
End If


If Text6 = "" Then Label5 = "Not Connected"
Command5.Enabled = True
Command1.Enabled = True
'if we are in the listmode, then do everything again to finish the list!
If ConType = "list" Then Command5_Click
End Sub

Private Sub Timer1_Timer()
'i just put this because if this loads in the load of the form
'the form will only show once this has finished
'and this takes a couple of seconds to finish!
    Dim MyIP As String
    MyIP = Inet1.OpenURL("http://pchelplive.com/ip.php")
    Text2 = MyIP
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
'this is the ping counter!

'This is the timeout value
'it can also timeout if winsock timesout!
If X = "10000" Then
Timer2.Enabled = False
Set Litem = ListView1.ListItems.Add(, , Text3)
Litem.ListSubItems.Add , , Text4
Litem.ListSubItems.Add , , "Timeout"
Litem.ListSubItems.Add , , "NO"
Litem.ListSubItems.Add , , "-"
List1.Selected(0) = True
If List1.Text = "" Then Exit Sub
ParseText List1.Text
List1.RemoveItem 0
If ConType = "list" Then Command5_Click
If ConType = "one" Then Command1_Click
End If
X = X + 10
Label8 = X
End Sub

Private Sub Timer3_Timer()
'Time out for the real ip function
If Text2 = "Workin on it..." Then
Text2 = "Timedout! DblClick to retry!"
Inet1.Cancel
End If
Timer3.Enabled = False
End Sub

Private Sub ws1_Connect()

Label5 = "Testing..."
ws1.SendData strHeaders
End Sub

Private Sub ws1_DataArrival(ByVal bytesTotal As Long)
Dim X As String
ws1.GetData X
Text6 = Text6 & X
End Sub

Private Sub ws1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Label5 = "Error! - " & Description
Set Litem = ListView1.ListItems.Add(, , Text3)
Litem.ListSubItems.Add , , Text4
Litem.ListSubItems.Add , , "Timeout"
Litem.ListSubItems.Add , , "NO"
Litem.ListSubItems.Add , , "-"
List1.Selected(0) = True
If List1.Text = "" Then Exit Sub
ParseText List1.Text
List1.RemoveItem 0
If ConType = "list" Then Command5_Click
If ConType = "one" Then Command1_Click
End Sub
