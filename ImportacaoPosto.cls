VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImportacaoPosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private bolSoPrimeira As Boolean

Public Property Get SoPrimeira() As Boolean
SoPrimeira = bolSoPrimeira
End Property

Public Property Let SoPrimeira(ByVal vNewValue As Boolean)
bolSoPrimeira = vNewValue
End Property
