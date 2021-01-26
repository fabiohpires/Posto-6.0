VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatCartoesPorBordero 
   Caption         =   "Cartões por data de Borderô"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmRelatCartoesPorBordero.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo cboFormaDePg 
      Bindings        =   "frmRelatCartoesPorBordero.frx":0442
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "descri"
      Text            =   ""
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc dbFormaDePG 
      Height          =   375
      Left            =   4680
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select codigopagamento, descri from formadepagamento order by descri"
      Caption         =   "dbFormaDePG"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin CRVIEWERLibCtl.CRViewer CR 
      Height          =   4695
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   9615
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
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   39331
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   5400
      TabIndex        =   5
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   39331
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Período a ser impresso:"
      Height          =   195
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label Label3 
      Caption         =   "a"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Forma de Pagamento:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmRelatCartoesPorBordero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CRReport As New CRAXDRT.Report
Dim CRApp As New CRAXDRT.Application

Private Sub cmdExibir_Click()
If Dir("Forma de Pg Recebido.rpt") = "" Then
  MsgBox "Não foi encontrado o arquivo 'Forma de Pg Recebido.rpt'"
  Exit Sub
End If
With dbFormaDePg
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboFormaDePg.Text = "" Then
    MsgBox "Selecione uma forma de pagamento!"
    cboFormaDePg.SetFocus
    Exit Sub
  End If
  .Recordset.Find "descri='" & cboFormaDePg.Text & "'"
  If .Recordset.EOF = True Then
    MsgBox "Forma de pagamento não localizada!"
    cboFormaDePg.SetFocus
    Exit Sub
  End If
End With

Set CRReport = CRApp.OpenReport("Forma de Pg Recebido.rpt")
'On Error GoTo ExitLabel
    With CRReport
        For i = 1 To .Database.Tables.Count
            .Database.Tables(i).Location = Caminho
        Next i
        .ParameterFields.GetItemByName("Codigoformadepg").AddCurrentValue CInt(dbFormaDePg.Recordset!CodigoPagamento)
        .ParameterFields.GetItemByName("Dataini").AddCurrentValue txtDataIni.Value
        .ParameterFields.GetItemByName("Datafim").AddCurrentValue txtDataFim.Value
        
                
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

With dbFormaDePg
  .ConnectionString = CaminhoADO
  .Refresh
End With

txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
End Sub

Private Sub Form_Paint()
CR.Width = Me.Width
CR.Height = Me.Height - 1000
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
