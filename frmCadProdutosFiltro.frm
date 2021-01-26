VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCadProdutosFiltro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtra Produtos"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2670
   Icon            =   "frmCadProdutosFiltro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtCodigoFim 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtCodigoIni 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin MSAdodcLib.Adodc dbCategoria 
      Height          =   330
      Left            =   1680
      Top             =   720
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select Categoria from Produtos group by Categoria order by Categoria"
      Caption         =   "dbCategoria"
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
   Begin MSAdodcLib.Adodc dbSubCategoria 
      Height          =   330
      Left            =   1680
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from ProdutosSubCategoria order by descri"
      Caption         =   "dbSubCategoria"
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
   Begin MSDataListLib.DataCombo cboSubCategoria 
      Bindings        =   "frmCadProdutosFiltro.frx":0442
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboCategoria 
      Bindings        =   "frmCadProdutosFiltro.frx":045F
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Categoria"
      Text            =   ""
   End
   Begin VB.Label Label4 
      Caption         =   "Sub-Categoria:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Categoria:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Até:"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Do Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCadProdutosFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Filtro As String

Private Sub cboCategoria_LostFocus()
With dbCategoria
  .Refresh
  If cboCategoria.Text = "" Then Exit Sub
  .Recordset.Find "Categoria='" & cboCategoria.Text & "'"
  If .Recordset.EOF = False Then
    cboCategoria.Text = .Recordset!Categoria
  End If
End With
End Sub

Private Sub cboSubCategoria_LostFocus()
With dbSubCategoria
  .Refresh
  If cboSubCategoria.Text <> "" Then
    .Recordset.Find "descri='" & cboSubCategoria.Text & "'"
    If .Recordset.EOF = False Then
      cboSubCategoria.Text = .Recordset!Descri
    End If
  End If
End With
End Sub

Private Sub cmdOk_Click()
If IsNumeric(txtCodigoIni.Text) = True Then
  If IsNumeric(txtCodigoFim.Text) = True Then
    If Filtro = "" Then
      Filtro = " where codigo between " & txtCodigoIni.Text & " and " & txtCodigoFim.Text
    Else
      Filtro = Filtro & " and codigo between " & txtCodigoIni.Text & " and " & txtCodigoFim.Text
    End If
  End If
End If
If cboCategoria.Text <> "" Then
  If Filtro = "" Then
    Filtro = " where categoria='" & cboCategoria.Text & "'"
  Else
    Filtro = Filtro & " and categoria='" & cboCategoria.Text & "'"
  End If
End If
If cboSubCategoria.Text <> "" Then
  If dbSubCategoria.Recordset.EOF = False Then
    If dbSubCategoria.Recordset!Descri = cboSubCategoria.Text Then
      If Filtro = "" Then
        Filtro = " where subcategoria=" & dbSubCategoria.Recordset!codigosubcategoria
      Else
        Filtro = Filtro & " and subcategoria=" & dbSubCategoria.Recordset!codigosubcategoria
      End If
    End If
  End If
End If
Me.Hide
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
With dbCategoria
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbSubCategoria
  .ConnectionString = CaminhoADO
  .Refresh
End With

End Sub

Private Sub txtCodigoFim_GotFocus()
With txtCodigoFim
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigoIni_Change()
txtCodigoFim.Text = txtCodigoIni.Text
End Sub

Private Sub txtCodigoIni_GotFocus()
With txtCodigoIni
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub
