Attribute VB_Name = "Runner4Ptable"
'Periodic Table Module and UserForm by Logan Boespflug (2016)
Option Explicit
Option Base 1

Public Element As String

'17 Variables in Total for the Periodic Table for each Element
Public AtomNum As Double
'Atomic Number (1)
Public AtomWeight As Double
'Atomic Weight based on Carbon-12 (2)
Public MeltPt As Double
'Variable (3)
Public BoilPt As Double
'Melting and Boiling point in deg K (4)
Public AtomDensity As Double
'Atomic density at 300K (g/cm^3), if gas then at 273K (g/L) (5)
Public ElectronConfig As String
'Electron Configuration (6)
Public CrystalStrut As String
'Variable (7)
Public ElectConduct As Double
'Electrical Conductivity in 10^-6 per ohm m (8)
Public CovRad As Double
'Covalent Radius in angstroms (9)
Public AtomRad As Double
'Atomic Radius in angstroms (10)
Public AtomVol As Double
'Atomic Volume in cm^3/mol (11)
Public FirstIonPot As Double
'First Ionization Potential as eV (12)
Public Cp As Double
'Specific heat (13)
Public Hvapor As Double
'Heat of vaporization in kJ/mol (14)
Public Hfusion As Double
'Heat of fusion in kJ/mol (15)
Public ThermConduct As Double
'Thermal Conductivity as W per m K (16)
Public Electroneg As Double
'Electronegativity as Paulings (17)

Public Sub getPtable()
Attribute getPtable.VB_Description = "Brings up the periodic table of elements"
Attribute getPtable.VB_ProcData.VB_Invoke_Func = "P\n14"
PeriodicTable.Show
End Sub
