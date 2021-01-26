VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatFormaPg2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumo de forma de pagamento recebido"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExibe 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "frmRelatFormaPg2.frx":0000
      TabIndex        =   0
      Top             =   840
      Width           =   7335
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   37767
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   37767
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "Período:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmRelatFormaPg2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExibe_Click()
Dim Ws As Workspace, Db As Database
Dim PgRecebido As Recordset, dbTemp As Recordset
Dim Coluna As Integer, Dias As Double, DataAtual As Date

'Set Ws = DBEngine.Workspaces(0)
'Set Db = Ws.OpenDatabase(Caminho, , , Conectar)
'Set PgRecebido = Db.OpenRecordset("select *from formadepagamento order by descri")
'If PgRecebido.RecordCount <> 0 Then
'  PgRecebido.MoveLast
'  PgRecebido.MoveFirst
'  Dias = DateDiff("d", txtDataIni.Value, txtDataFim.Value)
'  If Dias < 0 Then Dias = Dias * -1
'  DataAtual = txtDataIni.Value
'  DBGrid1.RowBuffer.RowCount = Dias - 1
'  For i = 0 To Dias - 1
'    DBGrid1.RowBuffer.Value(i, 0) = DataAtual
'    DataAtual = DateAdd("d", 1, DataAtual)
'  Next i
'  Do While PgRecebido.EOF = False
'    With DBGrid1
'      Coluna = PgRecebido.AbsolutePosition + 1
'      .Columns.Add (Coluna)
'      .Columns(Coluna).Caption = PgRecebido!Descri
'      .Columns(Coluna).Visible = True
'      .Columns(Coluna).NumberFormat = "Currency"
'
'      Set dbTemp = Db.OpenRecordset("select codigoformadepg, data, sum(valorBruto) as bruto, sum(valor) as liquido from formadepagamentorecebido where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# and fechamentodiario=-1 and codigoformadepg=" & PgRecebido!codigopagamento & " group by data, codigoformadepg order by data")
'
'    End With
'    PgRecebido.MoveNext
'  Loop
'End If

End Sub
