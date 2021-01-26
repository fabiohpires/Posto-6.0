VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatResumoDia 
   Caption         =   "Venda por dia"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10725
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   1095
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
      CurrentDate     =   37678
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
      CurrentDate     =   37678
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin CRVIEWERLibCtl.CRViewer CR 
      Height          =   6855
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   10695
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
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
End
Attribute VB_Name = "frmRelatResumoDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CRReport As New CRAXDRT.Report
Dim CRApp As New CRAXDRT.Application

Private Sub cmdExibir_Click()
Dim db As New ADODB.Connection
Dim dbVendas As New ADODB.Recordset
Dim dbEncerrantes As New ADODB.Recordset
Dim dbNotas As New ADODB.Recordset
Dim dbTemp As New ADODB.Recordset
Dim dbPDVs As New ADODB.Recordset

Dim Resposta As Integer
Dim DiaAtual As Date
Dim CodigoPdv As Double


If txtDataIni.Value > txtDataFim.Value Then
  MsgBox "A data inicial deve ser menor que a data final!"
  txtDataIni.SetFocus
  Exit Sub
End If

db.Open CaminhoADO

Screen.MousePointer = vbHourglass
cmdExibir.Enabled = False
ProgressBar1.Visible = True

db.Execute "delete *from resumodia where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"

dbTemp.CursorLocation = adUseServer
dbTemp.Open "select *from resumodia where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#", db, adOpenKeyset, adLockOptimistic



If dbTemp.RecordCount = 0 Then
  DiaAtual = txtDataIni.Value
  Do While DiaAtual <= txtDataFim.Value
    
    db.Execute "insert into resumodia (data,combustivel,produtos,total) values (#" & DataInglesa(DiaAtual) & "#,0,0,0)"
    
    DiaAtual = DateAdd("d", 1, DiaAtual)
  Loop
End If

dbPDVs.Open "select *from pdvs", db, adOpenKeyset, adLockOptimistic

If dbPDVs.RecordCount <> 0 Then
    CodigoPdv = dbPDVs!CodigoPdv
End If
dbVendas.Open "select  data, sum(valor) as total from qvendadiaprodutos2 where combustivel=0 and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# group by data", db, adOpenKeyset, adLockOptimistic
dbEncerrantes.Open "select datacaixa, sum(valortotal) as total from qbicoencerrantes where codigopdv=" & CodigoPdv & " and datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# group by datacaixa", db, adOpenKeyset, adLockOptimistic
dbNotas.Open "select clientesnota2.data, sum(lucrodif) as total from clientesnota2, produtos where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# and produtos.codigo=clientesnota2.codigoproduto group by data", db, adOpenKeyset, adLockOptimistic

dbTemp.Close
dbTemp.Open

If dbTemp.RecordCount <> 0 Then
  ProgressBar1.Max = dbTemp.RecordCount
  ProgressBar1.Visible = True
  
  Do While dbTemp.EOF = False
    If dbVendas.RecordCount <> 0 Then
      dbVendas.MoveFirst
      dbVendas.Find "data=#" & dbTemp!Data & "#"
      If dbVendas.EOF = False Then
        dbTemp!Produtos = dbVendas!Total
      End If
    End If
    If dbEncerrantes.RecordCount <> 0 Then
      dbEncerrantes.MoveFirst
      dbEncerrantes.Find "datacaixa=#" & dbTemp!Data & "#"
      If dbEncerrantes.EOF = False Then
        dbTemp!Combustivel = dbEncerrantes!Total
      End If
    End If
    If dbNotas.RecordCount <> 0 Then
      If dbNotas.RecordCount <> 0 Then
      dbNotas.MoveFirst
      dbNotas.Find "data=#" & dbTemp!Data & "#"
      If dbNotas.EOF = False Then
        dbTemp!Combustivel = dbTemp!Combustivel + dbNotas!Total
      End If
    End If
    End If
    dbTemp!Total = dbTemp!Combustivel + dbTemp!Produtos
    dbTemp.Update
    dbTemp.MoveNext
    If dbTemp.EOF = False Then ProgressBar1.Value = dbTemp.AbsolutePosition
  Loop
End If
ProgressBar1.Visible = False

dbVendas.Close
dbEncerrantes.Close
dbNotas.Close

db.Close



Screen.MousePointer = vbDefault

cmdExibir.Enabled = True

Exibe

End Sub

Private Sub Exibe()
Set CRReport = CRApp.OpenReport("Vendas por Dia.rpt")
'On Error GoTo ExitLabel
    With CRReport
        For i = 1 To .Database.Tables.Count
            .Database.Tables(i).Location = Caminho
        Next i
        .ParameterFields.GetItemByName("DataIni").AddCurrentValue txtDataIni.Value
        .ParameterFields.GetItemByName("DataFim").AddCurrentValue txtDataFim.Value
        .ParameterFields.GetItemByName("NomePosto").AddCurrentValue NomePosto
        
    End With
    With CR
        
        .ReportSource = CRReport
        .EnablePopupMenu = True
        .ViewReport
    End With
    CRApp.CanClose
    Exit Sub
'ExitLabel:
'    MsgBox "DungTran:" & Err.Description

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

End Sub

Private Sub Form_Paint()
CR.Width = Me.Width
CR.Height = Me.Height - 1400
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
