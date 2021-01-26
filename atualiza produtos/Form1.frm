VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "NEWTREND2"
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim strConn As String

Dim db As New ADODB.Connection
Dim dbOrigem As New ADODB.Recordset
Dim dbDestino As New ADODB.Recordset

strConn = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=A30Sigpo;Data Source=" & Text1.Text

db.Open strConn

dbOrigem.CursorLocation = adUseClient
dbOrigem.Open "select *from a30produtotemp", db, adOpenForwardOnly, adLockReadOnly

dbDestino.CursorLocation = adUseClient
dbDestino.Open "select *from a30produto", db, adOpenKeyset, adLockOptimistic

Do While dbOrigem.EOF = False
    dbDestino.AddNew
    For i = 0 To dbOrigem.Fields.Count - 1
        Select Case UCase(dbOrigem.Fields(i).Name)
            Case "ID"
                
            Case "PRODUTO"
                dbDestino(i) = Format(dbOrigem(i), "000000")
            Case "POSTO"
                dbDestino(i) = Format(dbOrigem(i), "000")
            Case Else
                dbDestino(i) = dbOrigem(i)
        End Select
        
    Next i
    
    dbDestino.Update
    
    dbOrigem.MoveNext
    
Loop


End Sub

