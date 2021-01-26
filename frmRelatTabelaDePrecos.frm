VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmRelatTabelaDePrecos 
   Caption         =   "Tabela de Pre�os"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CR 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
End
Attribute VB_Name = "frmRelatTabelaDePrecos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CRReport As New CRAXDRT.Report
Dim CRApp As New CRAXDRT.Application

Private Sub Form_Load()

Set CRReport = CRApp.OpenReport("Relatorio de Altera��o de Pre�os.rpt")
'On Error GoTo ExitLabel
    With CRReport
        For i = 1 To .Database.Tables.Count
            .Database.Tables(i).Location = Caminho
        Next i
        .ParameterFields.GetItemByName("NomePosto").AddCurrentValue NomePosto
        .ParameterFields.GetItemByName("DataCaixa").AddCurrentValue CDate(frmCadProdutosPreco.dbProdutosAltera.Recordset!DataCaixa)
        .ParameterFields.GetItemByName("Turno").AddCurrentValue CStr(frmCadProdutosPreco.dbProdutosAltera.Recordset!Turno)
        .ParameterFields.GetItemByName("CodigoProdutoAltera").AddCurrentValue CDbl(frmCadProdutosPreco.dbProdutosAltera.Recordset!codigoprodutoaltera)
        
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

Private Sub Form_Paint()
CR.Width = Me.Width
CR.Height = Me.Height - 450
End Sub


Private Sub Form_Resize()
CR.Width = Me.Width
CR.Height = Me.Height - 450
End Sub
