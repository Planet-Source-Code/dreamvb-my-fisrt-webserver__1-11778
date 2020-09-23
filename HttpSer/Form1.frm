VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple http Server"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   3195
      Top             =   1020
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   345
      Index           =   2
      Left            =   2145
      TabIndex        =   6
      Top             =   1110
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Stop"
      Height          =   345
      Index           =   1
      Left            =   1155
      TabIndex        =   5
      Top             =   1110
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   345
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   1110
      Width           =   930
   End
   Begin VB.TextBox txtpage 
      Height          =   285
      Left            =   660
      TabIndex        =   3
      Text            =   "main.htm"
      Top             =   630
      Width           =   2850
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   645
      TabIndex        =   2
      Text            =   "\"
      Top             =   135
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Page"
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Users As Integer
Dim ResData As String

Sub CGI(Data, Index As Integer)

Dim StrBuff As String

Dim StrBuffer As String
Dim mStart, mEnd As Integer
mEnd = Len(Data)

    mStart = InStr(Data, "Name")
    If mStart Then
        StrBuffer = Mid(Data, mStart, mEnd)
        End If
        '
        StrBuffer = Replace(StrBuffer, "=", "")
        StrBuffer = Replace(StrBuffer, "Name", "")
        StrBuffer = Replace(StrBuffer, "Message", "")
        StrBuffer = Replace(StrBuffer, "Comments", "")
        StrBuffer = Replace(StrBuffer, "Send", "")
        StrBuffer = Replace(StrBuffer, "Submit", "")
        StrBuffer = Replace(StrBuffer, "+", " ")
        StrBuffer = Replace(StrBuffer, "&", vbNewLine)
        StrBuffer = Replace(StrBuffer, "%0D", "")
        StrBuffer = Replace(StrBuffer, "%0A", vbNewLine)
        
        StrBuffer = Trim(StrBuffer)
        If Len(StrBuffer) > 8 Then
            StrBuffer = Replace(StrBuffer, " ", "&nbsp;")
            StrBuffer = Replace(StrBuffer, Chr(13), "<p>")
        
            Open App.Path & "\Board.htm" For Append As #1
            Print #1, StrBuffer & "<br>"
            Print #1, Date & Space(5) & Time
            Print #1, "<hr>"
            Close #1
            
            Winsock1(Index).SendData "<p><b>Thank you for your feed back</b></p>"
            Winsock1(Index).SendData "<p>&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; Return back to <a href=" & Chr(34) & Chr(34) & ">" & Winsock1(0).LocalHostName & "</a> </p>"
        End If
        
            b1 = InStr(Data, "TSerach")
                If b1 Then
                
                StrBuff = Mid(Data, b1 + 8, Len(Data))
                StrBuff = Replace(StrBuff, "Submit", "")
                StrBuff = Replace(StrBuff, "&", "")
                StrBuff = Replace(StrBuff, "=", "")
                StrBuff = Replace(StrBuff, "+", " ")
                StrBuff = Replace(StrBuff, Chr(10), "")
                StrBuff = Replace(StrBuff, Chr(13), "")
                StrBuff = Trim(StrBuff)
                LoadSerachData StrBuff, Index
    End If

    
End Sub

Sub LoadSerachData(FindWhat As String, Index As Integer)
Dim Filenum As Integer
Dim A1, A2 As Integer
Dim K As Integer

Filenum = FreeFile

Img = Replace("<h3><font face=ÿCopperplate Gothic Boldÿ color=ÿ#FF0000ÿ>Hound Search</font></h3><p align=ÿcenterÿ><img border=ÿ0ÿ src=ÿbulldog.gifÿ width=ÿ456ÿ height=ÿ70ÿ></p><p align=ÿrightÿ><font color=ÿ#0000FFÿ>Your number one search place</font></p><hr>", Chr(255), Chr(34))



Winsock1(Index).SendData Img

Open App.Path & "\Serach.txt" For Input As #Filenum
  Do While Not EOF(Filenum)
   Input #Filenum, StrBuffer
   
A1 = InStr(StrBuffer, FindWhat)
A2 = InStr(StrBuffer, "|")

Href = Trim(Mid(StrBuffer, A2 + 1, Len(StrBuffer)))

If A1 Then
    K = K + 1
        
        Winsock1(Index).SendData StrConv(Href & "<hr>", vbProperCase)
        Else
End If

Loop
Close #Filenum

    Winsock1(Index).SendData StrConv("<b>Found " & K & " Results for " & FindWhat & "</b>", vbProperCase)
    Winsock1(Index).SendData StrConv("<p><i>Want to Try searching for something else Go back to </i><i><a href=" & Chr(34) & "Serach.htm" & Chr(34) & ">search", vbProperCase)
    
End Sub

Function FindFile(FileName As String) As Boolean
If Dir(FileName) = "" Then
    FindFile = False
    Else
    FindFile = True
End If

End Function
Sub SendData(page, Index)
Dim databyte() As Byte


On Error Resume Next

If page = " " Then page = txtpage.Text

If FileExists(txtPath.Text & page) Then

                Open txtPath.Text & page For Binary Shared As #1
                ReDim databyte(0 To LOF(1))
                Get #1, , databyte()
                Close #1
                    Winsock1(Index).SendData databyte()
                    
Else
   Module1.Http404
   Winsock1(Index).SendData Module1.Http_404_Error
   
End If

End Sub
Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        Winsock1(0).Listen
        Command1(0).Enabled = False
        Command1(1).Enabled = True
    Case 1
        Winsock1(0).Close
        Command1(0).Enabled = True
        Command1(1).Enabled = False
    Case 2
        End
End Select

End Sub


Private Sub Form_Load()

 If Right(App.Path, 1) = "\" Then
    txtPath = App.Path
    Else
    txtPath = App.Path & "\"
 End If
    Winsock1(0).LocalPort = 80
    Form1.Caption = "Simple http Server: " & Winsock1(0).LocalIP

End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index = 0 Then
    Users = Users + 1
    Load Winsock1(Users)
    Winsock1(Users).LocalPort = 0
    Winsock1(Users).Accept requestID
End If
    If Err Then Err.Clear
    
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim StrData As String
Dim page As String

Winsock1(Index).GetData StrData
CGI StrData, Index

If Mid(StrData, 1, 3) = "GET" Then
    StrGet = InStr(StrData, "GET ")
    spc2 = InStr(StrGet + 5, StrData, " ")
    page = Mid$(StrData, StrGet + 5, spc2 - (StrGet + 4))
    
    SendData page, Index
End If

End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
Winsock1(Index).Close

End Sub
