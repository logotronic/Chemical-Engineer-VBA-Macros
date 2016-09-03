VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Choose2Show 
   Caption         =   "Choose what to show"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5505
   OleObjectBlob   =   "Choose2Show.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Choose2Show"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Canceler_Click()
Choose2Show.Hide
End Sub

Private Sub Shower_Click()

Dim Lbl As String
Dim Answer As Double

'Incoming Variable Names from other userform'
Lbl1 = "Atomic Number"
Answer1 = AtomNum
Lbl2 = "Atomic Weight [g]"
Answer2 = AtomWeight
Lbl3 = "Melting Point [K]"
Answer3 = MeltPt
Lbl4 = "Boiling Point [K]"
Answer4 = BoilPt
Lbl5 = "Atomic Density @300K [g/cm^3]"
Answer5 = AtomDensity
Lbl6 = "Electron Configuration"
Answer6 = ElectronConfig
Lbl7 = "Crystal Structure"
Answer7 = CrystalStrut
Lbl8 = "Electrical Conductivity @293K[10^6/ohm m]"
Answer8 = ElectConduct
Lbl9 = "Covalent Radius [Angstroms]"
Answer9 = CovRad
Lbl10 = "Atomic Radius [Angstroms]"
Answer10 = AtomRad
Lbl11 = "Atomic Volume [cm^3/mol]"
Answer11 = AtomVol
Lbl12 = "First Ionization Potential [eV]"
Answer12 = FirstIonPot
Lbl13 = "Specific Heat"
Answer13 = Cp
Lbl14 = "Heat of vaporization [kJ/mol]"
Answer14 = Hvapor
Lbl15 = "Heat of fusion [kJ/mol]"
Answer15 = Hfusion
Lbl16 = "Thermal Conductivity @300K[W/mK]"
Answer16 = ThermConduct
Lbl17 = "Electronegativity [Pauling's]"
Answer17 = Electroneg

Application.EnableEvents = False

If Variable1.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl1
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer1
ActiveCell.Offset(1, -2).Select
End If

If Variable2.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl2
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer2
ActiveCell.Offset(1, -2).Select
End If

If Variable3.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl3
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer3
ActiveCell.Offset(1, -2).Select
End If

If Variable4.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl4
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer4
ActiveCell.Offset(1, -2).Select
End If

If Variable5.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl5
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer5
ActiveCell.Offset(1, -2).Select
End If

If Variable6.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl6
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer6
ActiveCell.Offset(1, -2).Select
End If

If Variable7.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl7
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer7
ActiveCell.Offset(1, -2).Select
End If

If Variable8.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl8
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer8
ActiveCell.Offset(1, -2).Select
End If

If Variable9.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl9
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer9
ActiveCell.Offset(1, -2).Select
End If

If Variable10.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl10
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer10
ActiveCell.Offset(1, -2).Select
End If

If Variable11.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl11
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer11
ActiveCell.Offset(1, -2).Select
End If

If Variable12.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl12
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer12
ActiveCell.Offset(1, -2).Select
End If

If Variable13.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl13
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer13
ActiveCell.Offset(1, -2).Select
End If

If Variable14.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl14
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer14
ActiveCell.Offset(1, -2).Select
End If

If Variable15.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl15
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer15
ActiveCell.Offset(1, -2).Select
End If

If Variable16.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl16
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer16
ActiveCell.Offset(1, -2).Select
End If

If Variable17.Value = True Then
ActiveCell.Value = Element
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Lbl17
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = Answer17
ActiveCell.Offset(1, -2).Select
End If

Choose2Show.Hide
End Sub

Private Sub UserForm_Click()

End Sub
