VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatKilometragem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kilometragem de Clientes"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11235
   Icon            =   "frmRelatKilometragem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "dbfs"
      Height          =   2295
      Left            =   4080
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Data qNotasTotal 
         Caption         =   "qNotasTotal"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   $"frmRelatKilometragem.frx":0442
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Data qNotas 
         Caption         =   "qNotas"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   $"frmRelatKilometragem.frx":0512
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Data dbClientes 
         Caption         =   "dbClientes"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from clientes order by nome"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Data dbCarros 
         Caption         =   "dbCarros"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from ClientesCarros where codigocliente=0 order by placa"
         Top             =   720
         Width           =   2655
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatKilometragem.frx":05E2
      Height          =   5175
      Left            =   120
      OleObjectBlob   =   "frmRelatKilometragem.frx":05F7
      TabIndex        =   10
      Top             =   840
      Width           =   10935
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   10200
      Picture         =   "frmRelatKilometragem.frx":1D36
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin MSDBCtls.DBCombo cboClientesNota 
      Bindings        =   "frmRelatKilometragem.frx":27B8
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   315
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37665
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   315
      Left            =   6480
      TabIndex        =   7
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37665
   End
   Begin MSDBCtls.DBCombo cboPlaca 
      Bindings        =   "frmRelatKilometragem.frx":27D1
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Placa"
      Text            =   ""
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Placa:"
      Height          =   195
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   6240
      TabIndex        =   6
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "frmRelatKilometragem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String, ColunaDeQuebra As Integer

Private Sub cboClientesNota_GotFocus()
With dbCarros
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from ClientesCarros where codigocliente=0 order by placa"
  .Refresh
End With
End Sub

Private Sub cboClientesNota_LostFocus()
With dbClientes
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboClientesNota.Text = "" Then Exit Sub
  .Recordset.FindFirst "nome='" & cboClientesNota.Text & "'"
  If .Recordset.NoMatch = False Then
    cboClientesNota.Text = .Recordset!Nome
    With dbCarros
      .Connect = Conectar
      .DatabaseName = Caminho
      .RecordSource = "select *from ClientesCarros where codigocliente=" & dbClientes.Recordset!CodigoCliente & " order by placa"
      .Refresh
    End With
  Else
    With dbCarros
      .Connect = Conectar
      .DatabaseName = Caminho
      .RecordSource = "select *from ClientesCarros where codigocliente=0 order by placa"
      .Refresh
    End With
  End If
End With
End Sub

Private Sub cboPlaca_LostFocus()
With dbCarros
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboPlaca.Text = "" Then Exit Sub
  .Recordset.FindFirst "placa='" & cboPlaca.Text & "'"
  If .Recordset.NoMatch = False Then
    cboPlaca.Text = .Recordset!Placa
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String
Dim StrTemp2 As String

If dbClientes.Recordset.EOF = True Then
  MsgBox "Tabela de clientes está vazia!"
  Exit Sub
End If

StrTemp = "select clientesnota2.*, clientesCarros.* from Clientesnota2, clientescarros where clientesnota2.codigocarro=clientescarros.codigocarro"
StrTemp2 = "select sum(valorprevisto) as total from Clientesnota2"

StrTemp = StrTemp & " and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
StrTemp2 = StrTemp2 & " where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"

If dbClientes.Recordset!Nome = cboClientesNota.Text Then
  StrTemp = StrTemp & " and clientescarros.codigocliente=" & dbClientes.Recordset!CodigoCliente
  StrTemp2 = StrTemp2 & " and codigocliente=" & dbClientes.Recordset!CodigoCliente
End If
If dbCarros.Recordset.RecordCount <> 0 Then
  If cboPlaca.Text = dbCarros.Recordset!Placa Then
    StrTemp = StrTemp & " and clientescarros.codigocarro=" & dbCarros.Recordset!codigocarro
    StrTemp2 = StrTemp2 & " and codigocarro=" & dbCarros.Recordset!codigocarro
  End If
End If

StrTemp = StrTemp & strOrdem

With qNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTemp
  .Refresh
End With
With qNotasTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTemp2
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format("0", "Currency")
  End If
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then GoTo NaoImprime
On Error GoTo 0

If dbClientes.Recordset.EOF = False Then
  If cboClientesNota.Text = dbClientes.Recordset!Nome Then
    StrTemp = Chr(vbKeyReturn) & "Cliente: " & cboClientesNota.Text
  End If
End If
If dbCarros.Recordset.EOF = False Then
  If cboPlaca.Text = dbCarros.Recordset!Placa Then
    StrTemp = StrTemp & Chr(vbKeyReturn) & "Veículo: " & dbCarros.Recordset!Veiculo & " - Placa: " & dbCarros.Recordset!Placa
  End If
End If
StrTemp = StrTemp & Chr(vbKeyReturn) & "Período: " & txtDataIni.Value & " a " & txtDataFim.Value

ImprimeGrid DBGrid1, Printer, qNotas, 4, , , ColunaDeQuebra, , , "Controle de Kilometragem", NomePosto, StrTemp

Printer.EndDoc

NaoImprime:

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If UCase(strOrdem) = UCase(" order by " & DBGrid1.Columns(ColIndex).DataField & ", clientescarros.codigocarro, km") Then
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc, clientescarros.codigocarro, km"
Else
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & ", clientescarros.codigocarro, km"
End If
If ColIndex = 8 Then
  ColunaDeQuebra = 8
Else
  ColunaDeQuebra = -1
End If
Call cmdExibir_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
strOrdem = " order by clientescarros.codigocliente, clientescarros.codigocarro, km"
ColunaDeQuebra = 8
With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCarros
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qNotasTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

End Sub

Private Sub txtDataFim_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataFim_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtDataIni_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataIni_LostFocus()
Me.KeyPreview = True
End Sub
