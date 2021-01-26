VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCartoesTransfere 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Trasnferência de Cartões"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "frmCartoesTransfere.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   120
      Picture         =   "frmCartoesTransfere.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Imprimir"
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Data dbCartoesTransfere 
      Caption         =   "dbCartoesTransfere"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Administrador\Meus documentos\Projeto For Windows\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CartoesTransfereHistorico"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2940
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmCartoesTransfere.frx":0EC4
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "frmCartoesTransfere.frx":0EE5
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmCartoesTransfere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdem As String

Private Sub Filtra()
Dim StrTemp As String
StrTemp = "select *from cartoestransferehistorico where codigoformadepg=" & frmCartoes.qPendentes.Recordset!CodigoFormaPg & StrOrdem
With dbCartoesTransfere
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTemp
  .Refresh
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String


On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

StrTemp = frmCartoes.qCartoes.Recordset!Descri & "  - Impresso em:  " & Format(Date, "Short Date")

ImprimeGrid DBGrid1, Printer, dbCartoesTransfere, , , , , , , "Transferência de Cartôes", NomePosto, StrTemp

Printer.EndDoc

NaoImprime:

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
Filtra
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
StrOrdem = " order by codigohistorico"
With dbCartoesTransfere
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
Filtra
End Sub
