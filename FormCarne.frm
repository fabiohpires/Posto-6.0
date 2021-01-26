VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormCarne 
   Caption         =   "Impressão de Carnê"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   4320
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Info"
      TabPicture(0)   =   "FormCarne.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TextCliente"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TextId"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TextTexto"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TextCoupom"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ComboVias"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TextTitulo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Parcelas"
      TabPicture(1)   =   "FormCarne.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "ComboParcelas"
      Tab(1).Control(3)=   "Label6"
      Tab(1).ControlCount=   4
      Begin VB.TextBox TextTitulo 
         Height          =   375
         Left            =   960
         MaxLength       =   20
         TabIndex        =   28
         Top             =   600
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datas"
         Height          =   2895
         Left            =   -71760
         TabIndex        =   17
         Top             =   960
         Width           =   2895
         Begin VB.CommandButton Command6 
            Caption         =   "Pick..."
            Height          =   375
            Left            =   2040
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Delete"
            Height          =   375
            Left            =   2040
            TabIndex        =   25
            Top             =   2400
            Width           =   735
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Add"
            Height          =   375
            Left            =   2040
            TabIndex        =   24
            Top             =   720
            Width           =   735
         End
         Begin VB.ListBox ListDates 
            Height          =   2010
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox TextCurDate 
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Valores"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   16
         Top             =   960
         Width           =   2895
         Begin VB.CommandButton Command3 
            Caption         =   "Delete"
            Height          =   375
            Left            =   2040
            TabIndex        =   23
            Top             =   2280
            Width           =   735
         End
         Begin VB.ListBox ListValues 
            Height          =   2010
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Add"
            Height          =   375
            Left            =   2040
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TextCurValue 
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.ComboBox ComboParcelas 
         Height          =   315
         Left            =   -73080
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox ComboVias 
         Height          =   315
         Left            =   4920
         TabIndex        =   12
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TextCoupom 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Text            =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mostra Assinatura"
         Height          =   855
         Left            =   3600
         TabIndex        =   7
         Top             =   480
         Width           =   2295
         Begin VB.OptionButton OptionAssina 
            Caption         =   "Não"
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   29
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton OptionAssina 
            Caption         =   "Sim"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox TextTexto 
         Height          =   1335
         Left            =   960
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox TextId 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox TextCliente 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Título"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Número de Parcelas"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Número de vias"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Cupom Fiscal"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Texto"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "CPF/RG"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   615
      End
   End
End
Attribute VB_Name = "FormCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Dim cupom As Integer
Dim ret As Integer
Dim values As String
Dim dates As String
Dim quantidade As Integer
Dim vias As Integer
Dim assina As Integer




Value = ""
For i = 0 To ListValues.ListCount - 1
    If Len(values) Then
        values = values + ";"
    End If
    
    values = values + ListValues.List(i)

Next


dates = ""
For i = 0 To ListDates.ListCount - 1
    If Len(dates) Then
        dates = dates + ";"
    End If
    
    dates = dates + ListDates.List(i)
Next

quantidade = ComboParcelas.ListIndex + 1
vias = ComboVias.ListIndex + 1

If OptionAssina(0).Value Then
    assina = 1
Else
    assina = 0
End If

If IsNumeric(TextCoupom.Text) Then
    cupom = CInt(TextCoupom.Text)
Else
    cupom = 0
End If



ret = Bematech_FI_ImprimeCarne(TextTitulo.Text, values, dates, quantidade, _
                                TextTexto.Text, TextCliente.Text, _
                                TextId.Text, cupom, vias, assina)
                                
If ret <> 1 Then
    Select Case ret
    Case 0
        MsgBox "Erro de comunicação"
    Case -1
        MsgBox "Erro de execução"
    Case -2
        MsgBox "Erro de parâmetros"
    
    End Select
    
    
End If






End Sub

Private Sub Command2_Click()
    If ListValues.ListCount < ComboParcelas.ListIndex + 1 Then
        If Len(TextCurValue.Text) Then
            ListValues.AddItem TextCurValue.Text
        End If
    End If
    
    
End Sub

Private Sub Command3_Click()


    
    If ListValues.ListIndex <> -1 Then
        ListValues.RemoveItem ListValues.ListIndex
    End If

    
    
End Sub

Private Sub Command4_Click()
    If ListDates.ListCount < ComboParcelas.ListIndex + 1 Then
        If Len(TextCurDate.Text) Then
            ListDates.AddItem TextCurDate.Text
        End If
        
    End If
    
    
    
End Sub

Private Sub Command5_Click()

    If ListDates.ListIndex <> -1 Then
        ListDates.RemoveItem ListDates.ListIndex
    End If

    
    

End Sub

Private Sub Command6_Click()
    
    FormCalendar.Show 1
   
    TextCurDate = FormCalendar.Calendar1.Value
    Unload FormCalendar

End Sub

Private Sub Form_Load()
ComboVias.ListIndex = 0
ComboParcelas.ListIndex = 0
SSTab1.TabIndex = 0



End Sub

