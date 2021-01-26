VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatFormaDePg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Forma de Pagamento Recebido"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   Icon            =   "frmRelatFormaDePg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optDataBanco 
      Caption         =   "Data Recebida no Banco"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.OptionButton optDataBordero 
      Caption         =   "Data do borderô"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.OptionButton optDataCaixa 
      Caption         =   "Data do Caixa"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.Data QPgRecebido2 
      Caption         =   "QPgRecebido2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmRelatFormaDePg.frx":0442
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data dbFormadePgTotalizado 
      Caption         =   "dbFormadePgTotalizado"
      Connect         =   "Access"
      DatabaseName    =   "C:\rede\dados\Atalai.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from FormaDePagamentototalizado order by descri"
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data dbCheques 
      Caption         =   "dbCheques"
      Connect         =   "Access"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cheques"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data dbFormadePg2 
      Caption         =   "dbFormadePg2"
      Connect         =   "Access"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from FormaDePagamento order by descri"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   7560
      Picture         =   "frmRelatFormaDePg.frx":0526
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "Imprimir"
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdExibe 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin MSDBCtls.DBCombo cboFormaPg 
      Bindings        =   "frmRelatFormaDePg.frx":0FA8
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Data dbFormadePg 
      Caption         =   "dbFormadePg"
      Connect         =   "Access"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from FormaDePagamento order by descri"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data QPgRecebidoTotal 
      Caption         =   "QPgRecebidoTotal"
      Connect         =   "Access"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valorbruto) as bruto, sum(valor) as Liquido from QFormaDePgRecebidoFechamento"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37767
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37767
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatFormaDePg.frx":0FC2
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "frmRelatFormaDePg.frx":0FE6
      TabIndex        =   20
      Top             =   1440
      Width           =   8295
   End
   Begin VB.Label lblTotalCusto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblCheques 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Cheques no período:"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblTotalLiquido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Total Liquido:"
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label lblTotalBruto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Total Bruto:"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Período:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo de Recebimento:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Total Custo:"
      Height          =   255
      Left            =   6720
      TabIndex        =   19
      Top             =   5280
      Width           =   975
   End
End
Attribute VB_Name = "frmRelatFormaDePg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Public strOrdem As String, Indice As Integer

Private Sub Cabeca(ByVal Dia As Date, Largura As Double)
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Relatório de Forma de Pagamento Recebido"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10

StrTemp = "Data: " & Format(Dia, "Short Date") & " - " & Format(Dia, "Short time")
Printer.CurrentX = 0
Printer.Print StrTemp
If optDataCaixa.Value = True Then
  StrTemp = "Período do caixa: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short date")
  Printer.CurrentX = 0
  Printer.Print StrTemp
Else
  StrTemp = "Período do Borderô: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short date")
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If

If cboFormaPg.Text <> "" Then
  StrTemp = "Forma de Pagamento: " & cboFormaPg.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If
Printer.CurrentY = Printer.CurrentY + 1

StrTemp = DataGrid1.Columns(0).Caption
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = DataGrid1.Columns(1).Caption
Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = DataGrid1.Columns(2).Caption
Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub Cabeca2(ByVal Dia As Date, Largura As Double)
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Relatório de Forma de Pagamento Recebido"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10

StrTemp = "Data: " & Format(Dia, "Short Date") & " - " & Format(Dia, "Short time")
Printer.CurrentX = 0
Printer.Print StrTemp

If optDataCaixa.Value = True Then
  StrTemp = "Período do caixa: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short date")
  Printer.CurrentX = 0
  Printer.Print StrTemp
Else
  StrTemp = "Período do Borderô: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short date")
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If

If cboFormaPg.Text <> "" Then
  StrTemp = "Forma de Pagamento: " & cboFormaPg.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If
Printer.CurrentY = Printer.CurrentY + 1

StrTemp = DBGrid1.Columns(0).Caption
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = DBGrid1.Columns(1).Caption
Printer.CurrentX = 100 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = DBGrid1.Columns(2).Caption
Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = DBGrid1.Columns(3).Caption
Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = DBGrid1.Columns(4).Caption
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub Filtrar()
Dim StrTemp As String, StrTemp2 As String, strTemp3 As String
Dim Bruto As Currency, Liquido As Currency, Total As Currency
Dim tempBruto As Currency, tempLiquido As Currency, tempTotal As Currency
Dim TempTaxa As Double


lblTotalBruto.Caption = ""
lblTotalLiquido.Caption = ""

If optDataCaixa.Value = True Then
  StrTemp2 = "select *from QFormaDePgRecebidoFechamento2 where fechado=-1 and datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
ElseIf optDataBordero.Value = True Then
  StrTemp2 = "select *from QFormaDePgRecebidoFechamento2 where fechado=-1 and data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
Else
  StrTemp2 = "select *from cartoes where confirmado=-1 and datarecebida between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
End If
If cboFormaPg.Text <> "" Then
  If dbFormaDePg.Recordset.EOF = False Then
    If dbFormaDePg.Recordset!Descri <> cboFormaPg.Text Then
      MsgBox "Forma de pagamento não encontrada!"
      Exit Sub
    Else
      If optDataBanco.Value = False Then
        StrTemp2 = StrTemp2 & " and codigopagamento=" & dbFormaDePg.Recordset!CodigoPagamento
      Else
        StrTemp2 = StrTemp2 & " and codigoformapg=" & dbFormaDePg.Recordset!CodigoPagamento
      End If
    End If
  End If
End If

With QPgRecebido2
  .RecordSource = StrTemp2
  .Refresh
End With

cmdExibe.Enabled = False

With dbFormadePgTotalizado
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Delete
      .Refresh
    Loop
  End If
End With
With dbFormaDePg
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    tempBruto = 0
    tempLiquido = 0
    tempTotal = 0
    TempTaxa = 0
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      With dbFormadePgTotalizado
        .Recordset.AddNew
        .Recordset!Descri = dbFormaDePg.Recordset!Descri
        .Recordset!Taxa = TempTaxa
        .Recordset!ValorBruto = tempBruto
        .Recordset!valorliquido = tempLiquido
        .Recordset!custo = tempTotal
        .Recordset.Update
      End With
      .Recordset.MoveNext
    Loop
  End If
End With
With QPgRecebido2
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      With dbFormadePgTotalizado
        If optDataBanco.Value = False Then
          .Recordset.FindFirst "descri='" & QPgRecebido2.Recordset("formadepagamento.descri") & "'"
        Else
          .Recordset.FindFirst "descri='" & QPgRecebido2.Recordset("descri") & "'"
        End If
        If .Recordset.NoMatch = False Then
          If optDataBanco.Value = False Then
            If QPgRecebido2.Recordset!ValorDesconto <> 0 And QPgRecebido2.Recordset!ValorBruto <> 0 Then
              TempTaxa = (QPgRecebido2.Recordset!ValorDesconto / QPgRecebido2.Recordset!ValorBruto) * 100
            Else
              TempTaxa = 0
            End If
            tempTotal = QPgRecebido2.Recordset!ValorBruto - QPgRecebido2.Recordset!Valor
            tempBruto = QPgRecebido2.Recordset!ValorBruto
            tempLiquido = QPgRecebido2.Recordset!Valor
          Else
            tempBruto = QPgRecebido2.Recordset!ValorBruto
            tempLiquido = QPgRecebido2.Recordset!ValorRecebido
            If tempLiquido <> 0 And tempBruto <> 0 Then
              TempTaxa = (1 - (tempLiquido / tempBruto)) * 100
            Else
              TempTaxa = 0
            End If
            tempTotal = tempBruto - tempLiquido
          End If
          .Recordset.Edit
          If .Recordset!Taxa < TempTaxa Then
            .Recordset!Taxa = TempTaxa
          End If
          .Recordset!ValorBruto = .Recordset!ValorBruto + tempBruto
          .Recordset!valorliquido = .Recordset!valorliquido + tempLiquido
          .Recordset!custo = .Recordset!custo + tempTotal
          .Recordset.Update
        End If
      End With
      DoEvents
      .Recordset.MoveNext
    Loop
  End If
End With
With QPgRecebidoTotal
  .RecordSource = "select sum(valorbruto) as bruto, sum(custo) as Total, sum(valorliquido) as Liquido from formadepagamentototalizado"
  .Refresh
  If IsNull(.Recordset!Bruto) = False Then
    Bruto = Bruto + .Recordset!Bruto
  End If
  If IsNull(.Recordset!Liquido) = False Then
    Liquido = Liquido + .Recordset!Liquido
  End If
  If IsNull(.Recordset!Total) = False Then
    Total = .Recordset!Total
  End If
End With
strTemp3 = ""
With dbCheques
  .RecordSource = "Select sum(valor) as total from qchequescx where datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblCheques.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblCheques.Caption = Format(0, "Currency")
  End If
End With
dbFormadePgTotalizado.Refresh
lblTotalBruto.Caption = Format(Bruto, "Currency")
lblTotalLiquido.Caption = Format(Liquido, "Currency")
lblTotalCusto.Caption = Format(Total, "Currency")
cmdExibe.Enabled = True
End Sub

Private Sub cboFormaPg_LostFocus()
With dbFormaDePg
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboFormaPg.Text = "" Then Exit Sub
  On Error Resume Next
  .Recordset.FindFirst "descri='" & cboFormaPg.Text & "'"
  If .Recordset.NoMatch = False Then
    cboFormaPg.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdExibe_Click()
Filtrar
End Sub

Private Sub cmdImprime_Click()
Dim TempValor As Currency, TempValor2 As Currency
Dim TempValor3 As Currency, TempValor4 As Currency
Dim Dia As Date, Largura As Double, ValorAntigo As String
Dim TotalCusto As Currency


With dbFormadePgTotalizado
  If .Recordset.RecordCount <> 0 Then
    On Error GoTo NaoImprime
    If ShowPrinter(Me) = 0 Then Exit Sub
    On Error GoTo 0
  
    Dia = Now
    Largura = 190
    
    Printer.ScaleMode = vbMillimeters
    Printer.FontName = "Arial"
    
    Cabeca2 Dia, Largura
    
    Printer.FontSize = 10
    
    .Recordset.MoveLast
    .Recordset.MoveFirst
    ValorAntigo = DBGrid1.Columns(Indice).Text
    Do While .Recordset.EOF = False
      If Printer.CurrentY > Printer.ScaleHeight - 35 Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        StrTemp = Format(TempValor, "Currency")
        Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = Format(TempValor2, "Currency")
        Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = Format(TotalCusto, "Currency")
        Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        Printer.NewPage
        Cabeca2 Dia, Largura
      End If
      
      StrTemp = DBGrid1.Columns(0).Text
      Printer.CurrentX = 0
      Printer.Print StrTemp;
      
      StrTemp = DBGrid1.Columns(1).Text
      Printer.CurrentX = 100 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      
      TempValor = TempValor + CCur(DBGrid1.Columns(2).Text)
      StrTemp = DBGrid1.Columns(2).Text
      Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      TempValor2 = TempValor2 + CCur(DBGrid1.Columns(3).Text)
      StrTemp = DBGrid1.Columns(3).Text
      Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      TotalCusto = TotalCusto + CCur(DBGrid1.Columns(4).Text)
      StrTemp = DBGrid1.Columns(4).Text
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      .Recordset.MoveNext
    Loop
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    StrTemp = Format(TempValor, "Currency")
    Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(TempValor2, "Currency")
    Printer.CurrentX = 160 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(TotalCusto, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    StrTemp = "Cheques no Período: " & lblCheques.Caption
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
  End If
End With
Printer.EndDoc
NaoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = DataGrid1.Columns(ColIndex).DataField & ", " & DataGrid1.Columns(0).DataField Then
  strOrdem = DataGrid1.Columns(ColIndex).DataField & ", " & DataGrid1.Columns(0).DataField & " desc"
Else
  strOrdem = DataGrid1.Columns(ColIndex).DataField & ", " & DataGrid1.Columns(0).DataField
End If
Indice = ColIndex
QPgRecebido.Recordset.Sort = strOrdem
QPgRecebido2.Recordset.Sort = strOrdem
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = DBGrid1.Columns(ColIndex).DataField & ", " & DBGrid1.Columns(0).DataField Then
  strOrdem = DBGrid1.Columns(ColIndex).DataField & ", " & DBGrid1.Columns(0).DataField & " desc"
Else
  strOrdem = DBGrid1.Columns(ColIndex).DataField & ", " & DBGrid1.Columns(0).DataField
End If
Indice = ColIndex
'QPgRecebido.Recordset.Sort = StrOrdem
'QPgRecebido2.Recordset.Sort = StrOrdem
With dbFormadePgTotalizado
  .RecordSource = "Select *from formadepagamentototalizado order by " & strOrdem
  .Refresh
End With
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
txtDataIni.Value = Date
txtDataFim.Value = Date
Indice = 0
With dbFormadePgTotalizado
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With QPgRecebidoTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valorbruto) as bruto, sum(custo) as Total, sum(valorliquido) as Liquido from formadepagamentototalizado"
  .Refresh
End With
With dbFormaDePg
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With QPgRecebido2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from QFormaDePgRecebidoFechamento2 where fechamentodecaixa.codigofechamento=0"
  .Refresh
End With
With dbFormadePg2
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCheques
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
