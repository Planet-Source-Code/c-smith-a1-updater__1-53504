VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "oVoy Online Updater"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "updater.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.ListBox lstStat 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This can EASILY be modified to
'be an updater for any program!

'To use, include this program in your program's
'folder upon installation. When someone checks
'for newer versions of your program, launch
'this one right before ending the main prog
'and it will update and re-open the new version
'automatically.

'Note:Replace every instance of 'program.exe' with
'the name of the executable you're updating.


Private m_GettingFileSize     As Boolean
Private m_DownloadingFile     As Boolean
Private m_DownloadingFileSize As Long
Private m_LocalSaveFile       As String
Private m_FileSize As String
Private FirstResponse As Boolean

Private Sub cmdCancel_Click()
On Error GoTo Error
Inet1.Cancel
'exit this program and open the main
'exe back up.
Shell App.Path & "\program.exe", vbNormalFocus
End

Exit Sub
Error:
'if it can't be found
MsgBox "Couldn't find program.exe", vbCritical, "Error"
End Sub




Private Sub Form_Load()


If App.PrevInstance Then
End
End If


Form1.Show

Dim RemoteFileToGet As String
'Name of the updated exe
RemoteFileToGet = "http://www.yoursite.com/program.exe"

FirstResponse = False
m_FileSize = GetHTTPFileSize(RemoteFileToGet)
lstStat.AddItem "Establishing file size & location..."
lblStatus.Caption = "0/" & (m_FileSize)
m_LocalSaveFile = App.Path & "\program.exe"
Inet1.Execute RemoteFileToGet, "GET " & Chr(34) & App.Path & "\program.exe" & Chr(34)



End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Error
Inet1.Cancel

Shell App.Path & "\program.exe", vbNormalFocus
End

Exit Sub
Error:
MsgBox "Couldn't find program.exe", vbCritical, "Error"
End Sub
Private Function GetHTTPFileSize(strHTTPFile As String) As Long
On Error GoTo ErrorHandler
    Dim GetValue As String
    Dim GetSize  As Long
    
    m_GettingFileSize = True
    
    Inet1.Execute strHTTPFile, "HEAD " & Chr(34) & strHTTPFile & Chr(34)

    Do Until Inet1.StillExecuting = False
        DoEvents
    Loop

    GetValue = Inet1.GetHeader("Content-length")
    
    Do Until Inet1.StillExecuting = False
        DoEvents
    Loop
    
    If IsNumeric(GetValue) = True Then
        GetSize = CLng(GetValue)
    Else
        GetSize = -1
    End If

    If GetSize <= 0 Then GetSize = -1

    m_GettingFileSize = False
    GetHTTPFileSize = GetSize
Exit Function

ErrorHandler:
    m_GettingFileSize = False
    GetHTTPFileSize = -1
End Function

Private Sub Inet1_StateChanged(ByVal State As Integer)


Dim vtData()  As Byte
Dim FreeNr    As Integer
Dim SizeDone  As Long
Dim bDone     As Boolean
Dim GetPerc   As Integer

Select Case State
    
     Case 1
        lstStat.AddItem "Trying to resolve host..."
    Case 2
        lstStat.AddItem "Host is resolved"
    Case 3
        lstStat.AddItem "Sending connection request..."
    Case 4
        lstStat.AddItem "Connected"
    Case 5
        lstStat.AddItem "Sending request..."
    Case 6
        lstStat.AddItem "Request sent"
    Case 7
    If FirstResponse = False Then
        lstStat.AddItem "Receiving response..."
        FirstResponse = True
        End If
        
    Case 8
    If FirstResponse = False Then
        lstStat.AddItem "Response received"
        FirstResponse = True
        End If
        
    Case 9
        lstStat.AddItem "Disconnecting..."
    Case 10
        lstStat.AddItem "Disconnected"
    
    Case 11
        lstStat.AddItem "Error downloading file"
        Call cmdCancel_Click

    Case 12
    
    If m_GettingFileSize = True Then
    Exit Sub
    End If
    
    FreeNr = FreeFile
    
    Open App.Path & "\program.exe" For Binary Access Write As FreeNr
                
    'this shows the status in real time
    'kinda fancy
    
                Do While Not bDone
                    vtData = Inet1.GetChunk(1024, icByteArray) ' Get next chunk.
                    
                    SizeDone = SizeDone + UBound(vtData)
                    
                    lblStatus.Caption = SizeDone & "/" & m_FileSize
                    
                    GetPerc = (SizeDone / m_FileSize) * 100
                    If GetPerc > 100 Then GetPerc = 100
                    If GetPerc < 0 Then GetPerc = 0
                    
                    Me.Caption = "Online Updater - " & GetPerc & "%"
                                        
                    Put #FreeNr, , vtData()           'chunk wegschrijven naar bestand
                    If UBound(vtData) = -1 Then
                        bDone = True  'Er zijn geen chunks meer, KLAAR DUS
                    Else
                        DoEvents      'Yield to other processes
                    End If
                Loop
                
                Close FreeNr
    
End Select
lstStat.ListIndex = lstStat.ListCount - 1

If GetPerc = 100 Then
Call cmdCancel_Click
End If


End Sub

