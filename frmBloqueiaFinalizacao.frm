VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBloqueiaFinalizacao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bloqueia Finalização"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmBloqueiaFinalizacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker ctlData1 
      Bindings        =   "frmBloqueiaFinalizacao.frx":0442
      DataField       =   "Data1"
      DataSource      =   "dbBloquear"
      Height          =   300
      Left            =   1080
      TabIndex        =   11
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   39489
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   2520
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin MSAdodcLib.Adodc dbBloquear 
         Height          =   330
         Left            =   120
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Select *from bloqueiafechamento"
         Caption         =   "dbBloquear"
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
      Begin MSAdodcLib.Adodc dbTurno1 
         Height          =   330
         Left            =   120
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   "Select *from turnos order by horaini"
         Caption         =   "dbTurno1"
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
      Begin MSAdodcLib.Adodc dbTurno2 
         Height          =   330
         Left            =   120
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   "Select *from turnos order by horaini"
         Caption         =   "dbTurno2"
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
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmBloqueiaFinalizacao.frx":0466
      DataField       =   "CodigoTurno1"
      DataSource      =   "dbBloquear"
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      BoundColumn     =   "CodigoTurno"
      Text            =   ""
   End
   Begin VB.CheckBox chkBloqueia2 
      Caption         =   "Bloquear"
      DataField       =   "Bloqueia2"
      DataSource      =   "dbBloquear"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   930
   End
   Begin VB.CheckBox chkBloqueia1 
      Caption         =   "Bloquear"
      DataField       =   "Bloqueia1"
      DataSource      =   "dbBloquear"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   930
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frmBloqueiaFinalizacao.frx":047D
      DataField       =   "CoditoTurno2"
      DataSource      =   "dbBloquear"
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      BoundColumn     =   "CodigoTurno"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker ctlData 
      Bindings        =   "frmBloqueiaFinalizacao.frx":0494
      DataField       =   "Data2"
      DataSource      =   "dbBloquear"
      Height          =   300
      Left            =   1080
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   39489
   End
   Begin VB.Label Label4 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Não Finalizar Conferência de Caixa a partir de:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   3270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Não Finalizar Fechamento de Caixa a partir de:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3300
   End
End
Attribute VB_Name = "frmBloqueiaFinalizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancela_Click()
dbBloquear.Recordset.CancelUpdate
Unload Me
End Sub

Private Sub cmdOk_Click()
  dbBloquear.Recordset.Update
  Unload Me
  Screen.MousePointer = vbDefault
End Sub

Private Sub ctlData1_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub ctlData1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub ctlData1_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub ctlData2_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub ctlData2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub ctlData2_LostFocus()
Me.KeyPreview = True
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
With dbTurno1
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbTurno2
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbBloquear
  .ConnectionString = CaminhoADO
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    .Recordset!Data1 = CDate("31/12/2060")
    .Recordset!Data2 = CDate("31/12/2060")
    .Recordset!bloqueia1 = 0
    .Recordset!bloqueia2 = 0
    .Recordset!codigoturno1 = 0
    .Recordset!coditoturno2 = 0
    .Recordset.Update
  End If
  .Recordset.MoveFirst
End With
End Sub
