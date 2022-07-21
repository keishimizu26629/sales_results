VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Contractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public sash_amount As Long
Public exterior_amount As Long
Public sanitary_amount As Long
Public kitchen_amount As Long
Public ribiken_amount As Long
Public panel_amount As Long
Public electric_amount As Long
Public others_amount As Long
Public name As String

Public Sub init(contract_name)
    name = contract_name
End Sub

Public Function add_amount(kind, amount)
    Select Case kind
        Case "L90000"
            sash_amount = sash_amount + amount
        Case "L90005"
            exterior_amount = exterior_amount + amount
        Case "L90007"
            sanitary_amount = sanitary_amount + amount
        Case "L90008"
            kitchen_amount = kitchen_amount + amount
        Case "L90009"
            ribiken_amount = ribiken_amount + amount
        Case "L90012"
            panel_amount = panel_amount + amount
        Case "L90018"
            electric_amount = electric_amount + amount
        Case Else
            others_amount = others_amount + amount
    End Select

End Function

