VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliaNovaCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cheques"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   Icon            =   "frmConciliaNovaCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   1680
      Top             =   2040
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      RecordSource    =   "Select *from bloqueiafechamento"
      Caption         =   "dbBloqueiaFechamento"
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
   Begin VB.Data qConcilia 
      Caption         =   "qConcilia"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ConciliaNova where codigoconta=0 order by compensado, data, datalanc"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdCompensar 
      Caption         =   "Compensar"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.Data dbConcilia 
      Caption         =   "dbConcilia"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ConciliaNova where codigoconta=0 order by compensado, data, datalanc"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConciliaNovaCheques.frx":0442
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "frmConciliaNovaCheques.frx":045B
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
   Begin MSComCtl2.DTPicker txtDataCompensa 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37257
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Compensado:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
End
Attribute VB_Name = "frmConciliaNovaCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoConta As Double

Public Sub AbreDados()
With dbConcilia
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from concilianova where codigoconta=" & CodigoConta & " and compensado=0 order by compensado, data, datalanc"
  .Refresh
End With
With qConcilia
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from concilianova where codigoconta=" & CodigoConta & " and compensado=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With

End Sub

Private Sub cmdCompensar_Click()

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtDataCompensa.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

With dbConcilia
  If .Recordset.EOF = True Then
    MsgBox "Selecione um registro para ser compensado!"
    Exit Sub
  End If
  If .Recordset!compensado = True Then
    MsgBox "Registro já compensado! Selecione outro!"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja compensar o registro atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  
  .Recordset.Edit
  .Recordset!compensado = True
  .Recordset!Data = txtDataCompensa.Value
  .Recordset.Update
End With
With dbConcilia
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from concilianova where codigoconta=" & CodigoConta & " and compensado=0 order by compensado, data, datalanc"
  .Refresh
End With
With qConcilia
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from concilianova where codigoconta=" & CodigoConta & " and compensado=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbConcilia_Reposition()
On Error Resume Next
txtDataCompensa.Value = dbConcilia.Recordset!DataLanc
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
AbreDados
End Sub

Private Sub txtDataCompensa_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataCompensa_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataCompensa_LostFocus()
Me.KeyPreview = True
End Sub
