VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatClientesChequesIncompletos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes de Cheque com cadastro incompleto"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Data dbClientes 
      Caption         =   "dbClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ChequesClientes where posicao=0 and consultado=0 and devolvidos=0 Order by nome"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatClientesChequesIncompletos.frx":0000
      Height          =   4695
      Left            =   120
      OleObjectBlob   =   "frmRelatClientesChequesIncompletos.frx":0019
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmRelatClientesChequesIncompletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdem As String

Private Sub CabecaSoma(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters

Printer.FontBold = False
Printer.FontName = "Arial"
Printer.FontSize = 14
StrTemp = "Clientes com cadastro incompleto"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp
StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
StrTemp = "Data: " & Format(Dia, "Long date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

End Sub

Private Sub ImprimeChequesSomados()
Dim Largura As Double, Dia As Date, StrTemp As String
Dim DiaAtual As Date, SubTotal As Currency, Total As Currency

With dbClientes
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo naoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.FontBold = False
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.DrawWidth = 2
  
  Largura = 180
  Dia = Now
  
  
  CabecaSoma Largura, Dia
  
  Printer.FontSize = 10
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 35 Then
      Printer.NewPage
      CabecaSoma Largura, Dia
    End If
    
    Y1 = Printer.CurrentY + 1
    
    Printer.FontSize = 8
    Printer.ForeColor = RGB(180, 180, 180)
    Printer.Line (0, Y1)-(190, Y1 + 33), , BF
  
    Printer.ForeColor = vbWhite
    Printer.Line (2, Y1 + 3)-(32, Y1 + 9), , BF
    Printer.Line (33, Y1 + 3)-(188, Y1 + 9), , BF
    Printer.Line (2, Y1 + 10)-(126, Y1 + 16), , BF
    Printer.Line (127, Y1 + 10)-(188, Y1 + 16), , BF
    Printer.Line (2, Y1 + 17)-(34, Y1 + 23), , BF
    Printer.Line (35, Y1 + 17)-(72, Y1 + 23), , BF
    Printer.Line (73, Y1 + 17)-(116, Y1 + 23), , BF
    Printer.Line (117, Y1 + 17)-(160, Y1 + 23), , BF
    Printer.Line (161, Y1 + 17)-(188, Y1 + 23), , BF
    Printer.Line (2, Y1 + 24)-(57, Y1 + 30), , BF
    Printer.Line (58, Y1 + 24)-(105, Y1 + 30), , BF
    Printer.Line (106, Y1 + 24)-(159, Y1 + 30), , BF
    Printer.Line (160, Y1 + 24)-(188, Y1 + 30), , BF
    
    Printer.FontName = "Arial"
    Printer.FontSize = 7
    Printer.ForeColor = vbBlack
    Printer.FillColor = vbBlack
    On Error Resume Next
    StrTemp = "Código"
    Printer.CurrentX = 3
    Printer.CurrentY = Y1 + 3
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 3
    StrTemp = .Recordset!codigochequecliente
    Printer.Print StrTemp
    
    StrTemp = "Nome"
    Printer.CurrentX = 34
    Printer.CurrentY = Y1 + 3
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 34
    StrTemp = .Recordset!Nome
    Printer.Print StrTemp
    
    StrTemp = "Endereço"
    Printer.CurrentX = 3
    Printer.CurrentY = Y1 + 10
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 3
    StrTemp = .Recordset!Endereco
    Printer.Print StrTemp
    
    StrTemp = "Bairro"
    Printer.CurrentX = 128
    Printer.CurrentY = Y1 + 10
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 128
    StrTemp = .Recordset!Codigo
    Printer.Print StrTemp
    
    StrTemp = "CEP"
    Printer.CurrentX = 3
    Printer.CurrentY = Y1 + 17
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 3
    StrTemp = .Recordset!CEP
    Printer.Print StrTemp
    
    StrTemp = "Telefone"
    Printer.CurrentX = 36
    Printer.CurrentY = Y1 + 17
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 36
    StrTemp = Format(.Recordset!telefone, "(###)####-####")
    Printer.Print StrTemp
    
    
    StrTemp = "CIC"
    Printer.CurrentX = 74
    Printer.CurrentY = Y1 + 17
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 74
    StrTemp = Format(.Recordset!cic, "##,###,###,###-##")
    Printer.Print StrTemp
    
    StrTemp = "RG"
    Printer.CurrentX = 118
    Printer.CurrentY = Y1 + 17
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 118
    StrTemp = Format(.Recordset!rg, "###,###,###,###-#")
    Printer.Print StrTemp
    
    StrTemp = "Emissão"
    Printer.CurrentX = 162
    Printer.CurrentY = Y1 + 17
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 162
    StrTemp = .Recordset!Origem & " - " & .Recordset!origem2
    Printer.Print StrTemp
    
    StrTemp = "CNPJ"
    Printer.CurrentX = 3
    Printer.CurrentY = Y1 + 24
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 3
    StrTemp = Format(.Recordset!cnpj, "##,###,###/####-##")
    Printer.Print StrTemp
    
    StrTemp = "I.E."
    Printer.CurrentX = 59
    Printer.CurrentY = Y1 + 24
    Printer.Print StrTemp
    StrTemp = ""
    Printer.CurrentX = 59
    StrTemp = Format(.Recordset!ie, "###,###,###,###")
    Printer.Print StrTemp
    
    StrTemp = "Carro"
    Printer.CurrentX = 106
    Printer.CurrentY = Y1 + 24
    Printer.Print StrTemp
    StrTemp = ""
    
    StrTemp = "Placa"
    Printer.CurrentX = 161
    Printer.CurrentY = Y1 + 24
    Printer.Print StrTemp
    StrTemp = ""
    
    Printer.CurrentY = Y1 + 33
    Printer.FontSize = 10
    
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  Printer.EndDoc
End With
naoImprime:

End Sub


Private Sub cmdImprimir_Click()
ImprimeChequesSomados
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If StrOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField Then
  StrOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc"
Else
  StrOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField
End If
With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from ChequesClientes where posicao=0 and consultado=0 and devolvidos=0" & StrOrdem
  .Refresh
End With
End Sub

Private Sub Form_Load()
StrOrdem = " order by nome"
With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
End Sub
