VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmmainwin 
   Caption         =   "Form2"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form2"
   ScaleHeight     =   5925
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "218.69.93.4"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtresponse 
      Height          =   3615
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmmainwin.frx":0000
      Top             =   1560
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Text            =   "http://www.uol.com.br"
      Top             =   240
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtstatus 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5640
      Width           =   5055
   End
   Begin MSWinsockLib.Winsock ws_http 
      Left            =   6600
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmmainwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
' Name: HTTP Client - pure WinSock
' Description:Allow retrieve HTML page s
'     ources anywhere from web, directly or vi
'     a proxy server, can access virtual domai
'     ns. Pure winsock, no any other component
'     s used! Wanna know web transactions in d
'     eep?
' By: Tair Abdurman
'
' Assumes:'based on HTTP 1.0 - RFC 1945
'see http://www.tair.freeservers.com for
'     more info, details and downloads!
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=5166&lngWId=1'for details.'**************************************

'based on HTTP 1.0 - RFC 1945
'see http://www.tair.freeservers.com for
'     more info, details and downloads!
Public JobURL As String
Public ResponseDocument As String
Public StepCount As Long
Public IsProxyUsed As Boolean
Public ServerHostIP As String
Public ServerPort As Long
'---------------------------------------
'     ---------------------
Dim LocalStepCounter As Long
Dim RequestHeader As String
Dim RequestTemplate As String
'---------------------------------------
'     ---------------------


Public Sub ActionStartup()
    


    If UCase(Left(JobURL, 7)) <> "HTTP://" Then
        MsgBox "Please enter url With http://", vbCritical + vbOK
        frmactionwait.Hide
        Unload frmactionwait
        Exit Sub
    End If
    
    LocalStepCounter = 0
    RequestHeader = ""
    RequestTemplate = "GET _$-$_$- HTTP/1.0" & Chr(13) & Chr(10) & _
    "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-comet, */*" & Chr(13) & Chr(10) & _
    "Accept-Language: en" & Chr(13) & Chr(10) & _
    "Accept-Encoding: gzip , deflate" & Chr(13) & Chr(10) & _
    "Cache-Control: no-cache" & Chr(13) & Chr(10) & _
    "Proxy-Connection: Keep-Alive" & Chr(13) & Chr(10) & _
    "User-Agent: SSM Agent 1.0" & Chr(13) & Chr(10) & _
    "Host: @$@@$@" & Chr(13) & Chr(10)
    pureURL = Right(JobURL, Len(JobURL) - 7)
    startPos = InStr(1, pureURL, "/")
    


    If startPos < 1 Then
        ServerAddress = pureURL


        documentURI = "/"
        Else
            ServerAddress = Left(pureURL, startPos - 1)


            documentURI = Right(pureURL, Len(pureURL) - startPos + 1)
            End If
            


            If ServerAddress = "" Or documentURI = "" Then
                MsgBox "Unable To detect target page!", vbCritical + vbOK
                frmactionwait.Hide
                Unload frmactionwait
                Exit Sub
            End If
            


            If IsProxyUsed Then
                


                If ServerHostIP = "" Then
                    MsgBox "Unable To detect proxy address!", vbCritical + vbOK
                    frmactionwait.Hide
                    Unload frmactionwait
                    Exit Sub
                End If
                
                RequestHeader = RequestTemplate
                RequestHeader = Replace(RequestHeader, "_$-$_$-", JobURL)
            Else
                ServerHostIP = ServerAddress
                ServerPort = 80
                RequestHeader = RequestTemplate
                RequestHeader = Replace(RequestHeader, "_$-$_$-", documentURI)
            End If
            
            Me.Show
            RequestHeader = Replace(RequestHeader, "@$@@$@", ServerAddress)
            RequestHeader = RequestHeader & Chr(13) & Chr(10)
            txtstatus.Text = "Connecting To server ..."
            txtstatus.Refresh
            
            ws_http.Connect ServerHostIP, ServerPort
        End Sub


Private Sub txt_status_Change()

End Sub

Private Sub Command1_Click()
JobURL = Text1
If Text2 <> "" Then
IsProxyUsed = True
ServerHostIP = Text2
Else
IsProxyUsed = False
End If
ActionStartup
Text2 = ServerHostIP
Text3 = serverhostadress
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub WS_HTTP_Close()
    ws_http.Close
    txtstatus.Text = "Transaction completed ..."
    txtstatus.Refresh

End Sub


Private Sub WS_HTTP_Connect()
    ws_http.SendData RequestHeader
    txtstatus.Text = "Connected, try To obtain page ..."
    txtstatus.Refresh
    frmmainwin.txtresponse.Text = ""
    frmmainwin.txtresponse.Refresh
End Sub


Private Sub WS_HTTP_DataArrival(ByVal bytesTotal As Long)
    Dim tmpString As String
    ws_http.GetData tmpString, vbString
    frmmainwin.txtresponse.Text = frmmainwin.txtresponse.Text & tmpString
    frmmainwin.txtresponse.Refresh
    txtstatus.Text = "Data from server, continue ..."
    txtstatus.Refresh
End Sub


Private Sub WS_HTTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ws_http.Close
    txtstatus.Text = "Errors occured ...  -  " & Description

    txtstatus.Refresh
End Sub

