VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmAtualizacao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atualização"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   Icon            =   "frmAtualizacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "temvale.selfip.com"
      URL             =   "http://temvale.selfip.com/atualizaposto.txt"
      Document        =   "atualizaposto.txt"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
   Begin VB.ListBox lstStat 
      Height          =   2595
      Left            =   0
      TabIndex        =   1
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
Attribute VB_Name = "frmAtualizacao"
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
Private ArquivoAtual As String
Private Servidor As String

Private Sub cmdCancel_Click()
On Error GoTo Error
Inet1.Cancel
'exit this program and open the main
'exe back up.
'Shell App.Path & "\program.exe", vbNormalFocus

Exit Sub
Error:
'if it can't be found
MsgBox Err.Number & " - " & Err.Description
End Sub




Private Sub Form_Load()
Dim A As Integer

If App.PrevInstance Then
  Unload Me
End If


Me.Show

Dim RemoteFileToGet As String
'Name of the updated exe

TentaDeNovo:

lstStat.Clear

RemoteFileToGet = ReadINI("Atualiza", "CaminhoVersao", "/atualiza/AtualizaPosto.txt", App.Path & "\Posto.ini")
Servidor = ReadINI("Atualiza", "Servidor", "temvale.selfip.com", App.Path & "\Posto.ini")

FirstResponse = False

ArquivoAtual = "AtualizaPosto.txt"

lstStat.AddItem "Procurando atualização..."

If Dir(App.Path & "\" & ArquivoAtual) <> "" Then
  Kill App.Path & "\" & ArquivoAtual
End If

With Inet1
  If .StillExecuting = False Then
    .RemoteHost = Servidor
    .Document = RemoteFileToGet
    .RemotePort = 80
  End If
  Do While .StillExecuting = True
    If Err.Number <> 0 Then Exit Sub
    DoEvents
  Loop
End With

StrTemp = Inet1.OpenURL(Inet1.URL)
If StrTemp <> 0 Then
  A = FreeFile()
  Open App.Path & "\" & ArquivoAtual For Output As #A
  Print #A, StrTemp
  Close #A
End If

If Dir(App.Path & "\" & ArquivoAtual) = "" Then
  Resposta = MsgBox("Não conseguiu fazer download! Deseja tentar de novo?", vbYesNo)
  If Resposta = vbNo Then Exit Sub
  If Resposta = vbYes Then GoTo TentaDeNovo
End If
StrTemp = ReadINI("Versao", "Versao", "", App.Path & "\" & ArquivoAtual)
If StrTemp = "" Then
  Resposta = MsgBox("Não conseguiu fazer download! Deseja tentar de novo?", vbYesNo)
  If Resposta = vbNo Then Exit Sub
  If Resposta = vbYes Then GoTo TentaDeNovo
End If
If StrTemp <> "" Then
  A = InStr(1, StrTemp, ".")
  Versao = CDbl(Mid(StrTemp, 1, A - 1)) * 10000000
  B = A + 1
  A = InStr(B, StrTemp, ".")
  Versao = Versao + CDbl(Mid(StrTemp, B, A - B)) * 10000
  B = A + 1
  Versao = Versao + CDbl(Mid(StrTemp, B))
  Revisao = (App.Major * 10000000) + (App.Minor * 10000) + (App.Revision)
  If Versao > Revisao Then
    Resposta = MsgBox("Existe atualização! Deseja atualizar agora?", vbYesNo)
    If Resposta = vbNo Then
      Unload Me
      Exit Sub
    End If
    
    StrTemp = ReadINI("Versao", "Caminho", "", App.Path & "\" & ArquivoAtual)
    ArquivoAtual = App.Path & "\AtualizaPostoDeCombustivel.exe"
    If Dir(ArquivoAtual) <> "" Then
      Kill ArquivoAtual
    End If
    lstStat.AddItem "Fazendo download da atualização..."
    lstStat.AddItem "Aguarde. Isso pode levar alguns minutos..."
    With Inet1
      .RemoteHost = Servidor
      .Document = StrTemp
      '.Document = RemoteFileToGet
      Dim bt() As Byte
      Open ArquivoAtual For Binary Access Write As #A
      bt() = Inet1.OpenURL(.URL, icByteArray)
      Put #A, , bt()
      Close #A
    End With
    If Dir(ArquivoAtual) <> "" Then
      Shell ArquivoAtual, vbNormalFocus
      End
    End If
  End If
  Unload Me
  Exit Sub
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Error
Inet1.Cancel

Exit Sub
Error:

MsgBox Err.Number & " - " & Err.Description

End Sub
Private Function GetHTTPFileSize(strHTTPFile As String) As Long
On Error GoTo ErrorHandler
    Dim GetValue As String
    Dim GetSize  As Long
    
    m_GettingFileSize = True
    
    With Inet1
      .RemoteHost = Servidor
      .Document = strHTTPFile
      .Execute , "Head"
    End With
    
    
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
    MsgBox Err.Number & " - " & Err.Description
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
      lstStat.AddItem "Localizando o Computador remoto..."
  Case 2
      lstStat.AddItem "Localizado"
  Case 3
      lstStat.AddItem "Conectando ao computador remoto..."
  Case 4
      lstStat.AddItem "Conectado"
  Case 5
      lstStat.AddItem "Solicitando informações..."
  Case 6
      lstStat.AddItem "Solicitação enviada"
  Case 7
    If FirstResponse = False Then
      lstStat.AddItem "Recebendo resposta..."
      FirstResponse = True
    End If
  Case 8
    If FirstResponse = False Then
      lstStat.AddItem "Resposta recebida"
      FirstResponse = True
    End If
  Case 9
      lstStat.AddItem "Desconectando..."
  Case 10
      lstStat.AddItem "Desconectado"
  Case 11
      lstStat.AddItem "Erro baixando o arquivo"
      Call cmdCancel_Click
  Case 12
    If m_GettingFileSize = True Then
      Exit Sub
    End If
    FreeNr = FreeFile
  
    Open App.Path & "\" & ArquivoAtual For Binary Access Write As FreeNr
              
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
                            
        Put #FreeNr, , vtData()
        If UBound(vtData) = -1 Then
            bDone = True
        Else
            DoEvents
        End If
    Loop
    
    Close FreeNr
  
End Select
lstStat.ListIndex = lstStat.ListCount - 1

If GetPerc = 100 Then
  Call cmdCancel_Click
End If

End Sub


