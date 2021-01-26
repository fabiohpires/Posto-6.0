VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadChequesClienteLocalizar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localizar Cheque"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmCadChequesClienteLocalizar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc qCheques 
      Height          =   330
      Left            =   3720
      Top             =   3600
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from cheques"
      Caption         =   "qCheques"
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
   Begin VB.ComboBox cboStatusCheque 
      Height          =   315
      ItemData        =   "frmCadChequesClienteLocalizar.frx":0442
      Left            =   8400
      List            =   "frmCadChequesClienteLocalizar.frx":0458
      TabIndex        =   15
      Text            =   "Todos"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdProcurar 
      Caption         =   "Procurar"
      Height          =   375
      Left            =   10200
      TabIndex        =   17
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtCod 
      Height          =   285
      Left            =   7560
      TabIndex        =   13
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtCMC7 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc dbCheques 
      Height          =   330
      Left            =   3720
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from cheques"
      Caption         =   "dbCheques"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCadChequesClienteLocalizar.frx":04A0
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   19
      BeginProperty Column00 
         DataField       =   "CodigoCliente"
         Caption         =   "Cod."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Compensado"
         Caption         =   "Compensado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Comp"
         Caption         =   "Comp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Banco"
         Caption         =   "Banco"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Agencia"
         Caption         =   "Agencia"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Conta"
         Caption         =   "Conta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "ChequeNr"
         Caption         =   "ChequeNr"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "DataCheque"
         Caption         =   "DataCheque"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Valor"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Devolvido"
         Caption         =   "Devolvido"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "DataDevolucao"
         Caption         =   "DataDevolucao"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Cobrando"
         Caption         =   "Cobrando"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "DataCobrando"
         Caption         =   "DataCobrando"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "DataPgto"
         Caption         =   "DataPgto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "ValorPgto"
         Caption         =   "ValorPgto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Protesto"
         Caption         =   "Protesto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "DataProtesto"
         Caption         =   "DataProtesto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "EmpresaDeCobranca"
         Caption         =   "Emp.Cob."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   1
            TrueValue       =   "S"
            FalseValue      =   "N"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "DataEmpresaDeCobranca"
         Caption         =   "DataEmpresaDeCobranca"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   569,764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1409,953
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1319,811
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   360
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   2
      Left            =   5040
      TabIndex        =   4
      Top             =   360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   3
      Left            =   5760
      TabIndex        =   5
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   8
      Mask            =   "999999-9"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   4
      Left            =   6720
      TabIndex        =   6
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   6
      Mask            =   "999999"
      PromptChar      =   " "
   End
   Begin VB.Label lblCheques 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8040
      TabIndex        =   21
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Cheques:"
      Height          =   255
      Left            =   7320
      TabIndex        =   20
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9600
      TabIndex        =   19
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Total:"
      Height          =   255
      Left            =   9120
      TabIndex        =   18
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Posição do Cheque:"
      Height          =   255
      Left            =   8400
      TabIndex        =   14
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Código:"
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      Caption         =   "Comp:"
      Height          =   195
      Left            =   3840
      TabIndex        =   12
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   195
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      Caption         =   "Agência:"
      Height          =   195
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "Cheque:"
      Height          =   195
      Left            =   6720
      TabIndex        =   8
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label8 
      Caption         =   "Leitor de Código de barras:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmCadChequesClienteLocalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProcuraCheque()
  Dim Indice As String, StrTemp As String
  
  With dbCheques
    Indice = .Recordset.Sort
    If IsNumeric(txtCod.Text) = True Then
      If StrTemp = "" Then
        StrTemp = "codigocliente=" & txtCod.Text
      Else
        StrTemp = StrTemp & " and codigocliente=" & txtCod.Text
      End If
    End If
    If MaskEdBox1(0).Text <> "   " Then
      If StrTemp = "" Then
        StrTemp = "comp like '" & MaskEdBox1(0).Text & "*'"
      Else
        StrTemp = StrTemp & " and comp like '" & MaskEdBox1(0).Text & "*'"
      End If
    End If
    If MaskEdBox1(1).Text <> "   " Then
      If StrTemp = "" Then
        StrTemp = "banco like '" & MaskEdBox1(1).Text & "*'"
      Else
        StrTemp = StrTemp & " and banco like '" & MaskEdBox1(1).Text & "*'"
      End If
    End If
    If MaskEdBox1(2).Text <> "    " Then
      If StrTemp = "" Then
        StrTemp = "agencia like '" & MaskEdBox1(2).Text & "*'"
      Else
        StrTemp = StrTemp & " and agencia like '" & MaskEdBox1(2).Text & "*'"
      End If
    End If
    If MaskEdBox1(3).Text <> "      - " Then
      If StrTemp = "" Then
        StrTemp = "conta like '" & MaskEdBox1(3).Text & "*'"
      Else
        StrTemp = StrTemp & " and conta like '" & MaskEdBox1(3).Text & "*'"
      End If
    End If
    If MaskEdBox1(4).Text <> "      " Then
      If StrTemp = "" Then
        StrTemp = "chequeNr like '" & MaskEdBox1(4).Text & "*'"
      Else
        StrTemp = StrTemp & " and chequeNr like '" & MaskEdBox1(4).Text & "*'"
      End If
    End If
    Select Case cboStatusCheque.Text
      Case "Compensados"
        If StrTemp = "" Then
          StrTemp = "compensado=-1 and devolvido=0"
        Else
          StrTemp = StrTemp & " and compensado=-1 and devolvido=0"
        End If
      Case "Custodiados"
        If StrTemp = "" Then
          StrTemp = "custodia=-1"
        Else
          StrTemp = StrTemp & " and custodia=-1"
        End If
      Case "Devolvidos"
        If StrTemp = "" Then
          StrTemp = "devolvido=-1"
        Else
          StrTemp = StrTemp & " and devolvido=-1"
        End If
      Case "Cobrando"
        If StrTemp = "" Then
          StrTemp = "cobrando=-1"
        Else
          StrTemp = StrTemp & " and cobrando=-1"
        End If
      Case "Protestados"
        If StrTemp = "" Then
          StrTemp = "protesto=-1"
        Else
          StrTemp = StrTemp & " and protesto=-1"
        End If
      Case "Todos"
          
    End Select
    
    If StrTemp <> "" Then
      .Recordset.Filter = StrTemp
    Else
      .Refresh
    End If
    .Recordset.Sort = Indice
    lblCheques.Caption = .Recordset.RecordCount
  End With
  
  With qCheques
    .ConnectionString = CaminhoADO
    If StrTemp = "" Then
      .RecordSource = "select sum(valor) as total from cheques"
    Else
      .RecordSource = "select sum(valor) as total from cheques where " & StrTemp
    End If
    .Refresh
    If IsNull(.Recordset!Total) = False Then
      lblTotal.Caption = Format(.Recordset!Total, "Currency")
    Else
      lblTotal.Caption = Format(0, "Currency")
    End If
  End With
End Sub



Private Sub cmdProcurar_Click()
ProcuraCheque
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
With dbCheques
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from cheques"
  .Refresh
End With

End Sub

Private Sub MaskEdBox1_GotFocus(Index As Integer)
With MaskEdBox1(Index)
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCMC7_Change()
With txtCMC7
  If Len(.Text) >= 34 Then
    If Mid(.Text, 1, 1) <> "<" Then
      MsgBox "Erro de leitura"
      Exit Sub
    End If
    If UCase(Mid(.Text, 34, 1)) <> "Ç" Then
      MsgBox "Erro de leitura"
      Exit Sub
    End If
    MaskEdBox1(0).Text = Mid(.Text, 11, 3)
    MaskEdBox1(1).Text = Mid(.Text, 2, 3)
    MaskEdBox1(2).Text = Mid(.Text, 5, 4)
    MaskEdBox1(3).Text = Mid(.Text, 26, 6) & "-" & Mid(.Text, 32, 1)
    MaskEdBox1(4).Text = Mid(.Text, 14, 6)
    ProcuraCheque
  End If
End With
End Sub

