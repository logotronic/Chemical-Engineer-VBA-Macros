VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PeriodicTable 
   Caption         =   "Periodic Table"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8985
   OleObjectBlob   =   "PeriodicTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PeriodicTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Periodic Table Module and UserForm by Logan Boespflug (2016)
Private Sub UserForm_Initialize()
AtomNum = 0
AtomWeight = 0
MeltPt = 0
BoilPt = 0
AtomDensity = 0
ElectronConfig = "1s1"
CrystalStrut = "hexagonal"
ElectConduct = 0
CovRad = 0
AtomRad = 0
AtomVol = 0
FirstIonPot = 0
Cp = 0
Hvapor = 0
Hfusion = 0
ThermConduct = 0
Electroneg = 0
End Sub

Private Sub H_Click()
Element = "Hydrogen"
AtomNum = 1
AtomWeight = 1.00794
MeltPt = 13.81
BoilPt = 20.28
AtomDensity = 0.0899
ElectronConfig = "1s1"
CrystalStrut = "hexagonal"
ElectConduct = 0
CovRad = 0.32
AtomRad = 0.79
AtomVol = 14.1
FirstIonPot = 13.598
Cp = 14.304
Hvapor = 0.4581
Hfusion = 0.0585
ThermConduct = 0.1815
Electroneg = 2.1
Choose2Show.Show
End Sub
Private Sub He_Click()
Element = "Helium"
AtomNum = 2
AtomWeight = 4.0026
MeltPt = 0.95
BoilPt = 4.216
AtomDensity = 0.1785
ElectronConfig = "1s2"
CrystalStrut = "hexagonal"
ElectConduct = 0
CovRad = 0.93
AtomRad = 0.49
AtomVol = 31.8
FirstIonPot = 24.587
Cp = 5.193
Hvapor = 0.084
Hfusion = 0.021
ThermConduct = 0.152
Electroneg = 0
Choose2Show.Show
End Sub
Private Sub Li_Click()
Element = "Lithium"
AtomNum = 3
AtomWeight = 6.941
MeltPt = 453.7
BoilPt = 1615
AtomDensity = 0.53
ElectronConfig = "1s2 2s1"
CrystalStrut = "cubic, body centered"
ElectConduct = 11.7
CovRad = 1.23
AtomRad = 2.05
AtomVol = 13.1
FirstIonPot = 5.392
Cp = 3.582
Hvapor = 147.1
Hfusion = 3#
ThermConduct = 84.7
Electroneg = 0.98
Choose2Show.Show
End Sub
Private Sub Be_Click()
Element = "Beryllium"
AtomNum = 4
AtomWeight = 9.01218
MeltPt = 1560
BoilPt = 3243
AtomDensity = 1.85
ElectronConfig = "1s1 2s2"
CrystalStrut = "hexagonal"
ElectConduct = 25
CovRad = 0.9
AtomRad = 1.4
AtomVol = 5#
FirstIonPot = 9.322
Cp = 1.825
Hvapor = 297
Hfusion = 11.71
ThermConduct = 200
Electroneg = 1.57
Choose2Show.Show
End Sub
Private Sub B_Click()
Element = "Boron"
AtomNum = 5
AtomWeight = 10.811
MeltPt = 2365
BoilPt = 4275
AtomDensity = 2.34
ElectronConfig = "1s2 2s2p1"
CrystalStrut = "rhombohedral"
ElectConduct = 0.000000000005
CovRad = 0.82
AtomRad = 1.17
AtomVol = 4.6
FirstIonPot = 8.298
Cp = 1.026
Hvapor = 507.8
Hfusion = 22.6
ThermConduct = 27#
Electroneg = 2.04
Choose2Show.Show
End Sub
Private Sub C_Click()
Element = "Carbon"
AtomNum = 6
AtomWeight = 12.011
MeltPt = 3825
BoilPt = 5100
AtomDensity = 2.26
ElectronConfig = "1s1 2s2p2"
CrystalStrut = "hexagonal"
ElectConduct = 0.07
CovRad = 0.77
AtomRad = 0.91
AtomVol = 5.3
FirstIonPot = 11.26
Cp = 0.709
Hvapor = -715
Hfusion = 0
ThermConduct = 230
Electroneg = 2.55
Choose2Show.Show
End Sub
Private Sub N_Click()
Element = "Nitrogen"
AtomNum = 7
AtomWeight = 14.0067
MeltPt = 63.15
BoilPt = 77.344
AtomDensity = 1.251
ElectronConfig = "1s1 2s2p3"
CrystalStrut = "hexagonal"
ElectConduct = 0
CovRad = 0.75
AtomRad = 0.75
AtomVol = 17.3
FirstIonPot = 14.534
Cp = 1.042
Hvapor = 2.7928
Hfusion = 0.36
ThermConduct = 0.02598
Electroneg = 3.04
Choose2Show.Show
End Sub
Private Sub O_Click()
Element = "Oxygen"
AtomNum = 8
AtomWeight = 15.9994
MeltPt = 54.8
BoilPt = 90.188
AtomDensity = 1.429
ElectronConfig = "1s1 2s2p4"
CrystalStrut = "cubic"
ElectConduct = 0
CovRad = 0.73
AtomRad = 0.65
AtomVol = 14#
FirstIonPot = 13.618
Cp = 0.92
Hvapor = 3.4109
Hfusion = 0.222
ThermConduct = 0.2674
Electroneg = 3.44
Choose2Show.Show
End Sub
Private Sub F_Click()
Element = "Fluorine"
AtomNum = 9
AtomWeight = 18.9984
MeltPt = 53.55
BoilPt = 85
AtomDensity = 1.696
ElectronConfig = "1s1 2s2p5"
CrystalStrut = "cubic"
ElectConduct = 0
CovRad = 0.72
AtomRad = 0.57
AtomVol = 17.1
FirstIonPot = 17.422
Cp = 0.824
Hvapor = 3.2698
Hfusion = 0.26
ThermConduct = 0.0279
Electroneg = 3.98
Choose2Show.Show
End Sub
Private Sub Ne_Click()
Element = "Neon"
AtomNum = 10
AtomWeight = 20.1797
MeltPt = 24.55
BoilPt = 27.1
AtomDensity = 0.9
ElectronConfig = "1s1 2s2p6"
CrystalStrut = "cubic, face centered"
ElectConduct = 0
CovRad = 0.71
AtomRad = 0.51
AtomVol = 16.9
FirstIonPot = 21.564
Cp = 1.03
Hvapor = 1.77
Hfusion = 0.34
ThermConduct = 0.0493
Electroneg = 0
Choose2Show.Show
End Sub
Private Sub Na_Click()
Element = "Sodium"
AtomNum = 11
AtomWeight = 22.98977
MeltPt = 371.6
BoilPt = 1156
AtomDensity = 0.97
ElectronConfig = "[Ne] 3s1"
CrystalStrut = "cubic, body centered"
ElectConduct = 20.1
CovRad = 1.54
AtomRad = 2.23
AtomVol = 23.7
FirstIonPot = 5.139
Cp = 1.23
Hvapor = 98.01
Hfusion = 2.601
ThermConduct = 141
Electroneg = 20.1
Choose2Show.Show
End Sub
Private Sub Mg_Click()
Element = "Magnesium"
AtomNum = 12
AtomWeight = 24.305
MeltPt = 922
BoilPt = 1380
AtomDensity = 1.74
ElectronConfig = "[Ne] 3s2"
CrystalStrut = "hexagonal"
ElectConduct = 22.4
CovRad = 1.36
AtomRad = 1.72
AtomVol = 14#
FirstIonPot = 7.646
Cp = 1.02
Hvapor = 127.6
Hfusion = 8.95
ThermConduct = 156
Electroneg = 1.31
Choose2Show.Show
End Sub
Private Sub Al_Click()
Element = "Aluminum"
AtomNum = 13
AtomWeight = 26.98154
MeltPt = 933.5
BoilPt = 2740
AtomDensity = 2.7
ElectronConfig = "[Ne] 3s2p1"
CrystalStrut = "cubic, face centered"
ElectConduct = 37.7
CovRad = 1.18
AtomRad = 1.82
AtomVol = 10#
FirstIonPot = 5.986
Cp = 0.9
Hvapor = 290.8
Hfusion = 10.7
ThermConduct = 237
Electroneg = 1.61
Choose2Show.Show
End Sub
Private Sub Si_Click()
Element = "Silicon"
AtomNum = 14
AtomWeight = 28.0855
MeltPt = 1683
BoilPt = 2630
AtomDensity = 2.33
ElectronConfig = "[Ne] 3s2p2"
CrystalStrut = "cubic, face centered"
ElectConduct = 0.0004
CovRad = 1.11
AtomRad = 1.46
AtomVol = 12.1
FirstIonPot = 8.151
Cp = 0.7
Hvapor = 359
Hfusion = 50.2
ThermConduct = 148
Electroneg = 1.9
Choose2Show.Show
End Sub
Private Sub P_Click()
Element = "Phosphorus"
AtomNum = 15
AtomWeight = 30.97376
MeltPt = 317.3
BoilPt = 553
AtomDensity = 1.82
ElectronConfig = "[Ne] 3s2p3"
CrystalStrut = "monoclinic"
ElectConduct = 0.000000000000001
CovRad = 1.06
AtomRad = 1.23
AtomVol = 17#
FirstIonPot = 10.486
Cp = 0.769
Hvapor = 12.4
Hfusion = 0.63
ThermConduct = 0.235
Electroneg = 2.19
Choose2Show.Show
End Sub
Private Sub S_Click()
Element = "Sulfur"
AtomNum = 16
AtomWeight = 32.066
MeltPt = 392.2
BoilPt = 717.82
AtomDensity = 2.07
ElectronConfig = "[Ne] 3s2p4"
CrystalStrut = "orthorhombic"
ElectConduct = 5E-16
CovRad = 1.02
AtomRad = 1.09
AtomVol = 15.5
FirstIonPot = 10.36
Cp = 0.71
Hvapor = 10
Hfusion = 1.73
ThermConduct = 0.269
Electroneg = 2.58
Choose2Show.Show
End Sub
Private Sub Cl_Click()
Element = "Chlorine"
AtomNum = 17
AtomWeight = 35.4527
MeltPt = 172.17
BoilPt = 239.18
AtomDensity = 3.214
ElectronConfig = "[Ne] 3s2p5"
CrystalStrut = "orthorhombic"
ElectConduct = 99999
CovRad = 0.99
AtomRad = 0.97
AtomVol = 18.7
FirstIonPot = 12.967
Cp = 0.48
Hvapor = 10.2
Hfusion = 3.21
ThermConduct = 0.0089
Electroneg = 3.16
Choose2Show.Show
End Sub
Private Sub Ar_Click()
Element = "Argon"
AtomNum = 18
AtomWeight = 39.948
MeltPt = 83.95
BoilPt = 87.45
AtomDensity = 1.784
ElectronConfig = "[Ne] 3s2p6"
CrystalStrut = "cubic, face centered"
ElectConduct = 99999
CovRad = 0.98
AtomRad = 0.88
AtomVol = 24.2
FirstIonPot = 15.759
Cp = 0.52
Hvapor = 6.506
Hfusion = 1.188
ThermConduct = 0.0177
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub K_Click()
Element = "Potassium"
AtomNum = 19
AtomWeight = 39.0983
MeltPt = 336.8
BoilPt = 1033
AtomDensity = 0.86
ElectronConfig = "[Ar] 4s1"
CrystalStrut = "cubic, body centered"
ElectConduct = 16.4
CovRad = 2.03
AtomRad = 2.77
AtomVol = 45.3
FirstIonPot = 4.341
Cp = 0.757
Hvapor = 76.9
Hfusion = 2.33
ThermConduct = 102.5
Electroneg = 0.82
Choose2Show.Show
End Sub
Private Sub Ca_Click()
Element = "Calcium"
AtomNum = 20
AtomWeight = 40.078
MeltPt = 1757
BoilPt = 1112
AtomDensity = 1.55
ElectronConfig = "[Ar] 4s2"
CrystalStrut = "cubic, face centered"
ElectConduct = 31.3
CovRad = 1.74
AtomRad = 2.23
AtomVol = 29.9
FirstIonPot = 6.113
Cp = 0.647
Hvapor = 154.67
Hfusion = 8.53
ThermConduct = 200
Electroneg = 1#
Choose2Show.Show
End Sub
Private Sub Sc_Click()
Element = "Scandium"
AtomNum = 21
AtomWeight = 44.9559
MeltPt = 1814
BoilPt = 3109
AtomDensity = 2.99
ElectronConfig = "[Ar] 3d1 4s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.5
CovRad = 1.44
AtomRad = 2.09
AtomVol = 15#
FirstIonPot = 6.54
Cp = 0.568
Hvapor = 304.8
Hfusion = 16.11
ThermConduct = 15.8
Electroneg = 1.36
Choose2Show.Show
End Sub
Private Sub Ti_Click()
Element = "Titanium"
AtomNum = 22
AtomWeight = 47.87
MeltPt = 1935
BoilPt = 3560
AtomDensity = 4.54
ElectronConfig = "[Ar] 3d2 4s2"
CrystalStrut = "hexagonal"
ElectConduct = 2.6
CovRad = 1.32
AtomRad = 2#
AtomVol = 10.6
FirstIonPot = 6.82
Cp = 0.523
Hvapor = 425.2
Hfusion = 18.6
ThermConduct = 21.9
Electroneg = 1.54
Choose2Show.Show
End Sub
Private Sub V_Click()
Element = "Vanadium"
AtomNum = 23
AtomWeight = 50.9415
MeltPt = 2163
BoilPt = 3650
AtomDensity = 6.11
ElectronConfig = "[Ar] 3d3 4s2"
CrystalStrut = "cubic, body centered"
ElectConduct = 4#
CovRad = 1.22
AtomRad = 1.92
AtomVol = 8.35
FirstIonPot = 6.74
Cp = 0.489
Hvapor = 446.7
Hfusion = 22.8
ThermConduct = 30.7
Electroneg = 1.63
Choose2Show.Show
End Sub
Private Sub Cr_Click()
Element = "Chromium"
AtomNum = 24
AtomWeight = 51.996
MeltPt = 2130
BoilPt = 2945
AtomDensity = 7.19
ElectronConfig = "[Ar] 3d5 4s1"
CrystalStrut = "cubic, body centered"
ElectConduct = 7.9
CovRad = 1.18
AtomRad = 1.85
AtomVol = 7.23
FirstIonPot = 6.768
Cp = 0.449
Hvapor = 339.5
Hfusion = 20
ThermConduct = 93.7
Electroneg = 1.66
Choose2Show.Show
End Sub
Private Sub Mn_Click()
Element = "Manganese"
AtomNum = 25
AtomWeight = 54.938
MeltPt = 1518
BoilPt = 2235
AtomDensity = 7.44
ElectronConfig = "[Ar] 3d5 4s2"
CrystalStrut = "cubic, body centered"
ElectConduct = 0.5
CovRad = 1.17
AtomRad = 1.65
AtomVol = 7.39
FirstIonPot = 7.435
Cp = 0.48
Hvapor = 219.74
Hfusion = 14.64
ThermConduct = 7.82
Electroneg = 1.55
Choose2Show.Show
End Sub
Private Sub Fe_Click()
Element = "Iron"
AtomNum = 26
AtomWeight = 55.845
MeltPt = 1808
BoilPt = 3023
AtomDensity = 7.874
ElectronConfig = "[Ar] 3d6 4s2"
CrystalStrut = "cubic, body centered"
ElectConduct = 11.2
CovRad = 1.17
AtomRad = 1.72
AtomVol = 7.1
FirstIonPot = 7.87
Cp = 0.449
Hvapor = 349.5
Hfusion = 13.8
ThermConduct = 80.2
Electroneg = 1.83
Choose2Show.Show
End Sub
Private Sub Co_Click()
Element = "Cobalt"
AtomNum = 27
AtomWeight = 58.9332
MeltPt = 1768
BoilPt = 3143
AtomDensity = 8.9
ElectronConfig = "[Ar] 3d7 4s2"
CrystalStrut = "hexagonal"
ElectConduct = 17.9
CovRad = 1.16
AtomRad = 1.67
AtomVol = 6.7
FirstIonPot = 7.86
Cp = 0.421
Hvapor = 373.3
Hfusion = 16.19
ThermConduct = 100
Electroneg = 1.88
Choose2Show.Show
End Sub
Private Sub Ni_Click()
Element = "Nickel"
AtomNum = 28
AtomWeight = 58.6934
MeltPt = 1726
BoilPt = 3005
AtomDensity = 8.9
ElectronConfig = "[Ar] 3d8 4s2"
CrystalStrut = "cubic, face centered"
ElectConduct = 14.6
CovRad = 1.15
AtomRad = 1.62
AtomVol = 6.6
FirstIonPot = 7.635
Cp = 0.444
Hvapor = 377.5
Hfusion = 17.2
ThermConduct = 90.7
Electroneg = 1.91
Choose2Show.Show
End Sub
Private Sub Cu_Click()
Element = "Copper"
AtomNum = 29
AtomWeight = 63.546
MeltPt = 1356.6
BoilPt = 2840
AtomDensity = 8.96
ElectronConfig = "[Ar] 3d10 4s1"
CrystalStrut = "cubic, face centered"
ElectConduct = 60.7
CovRad = 1.17
AtomRad = 1.57
AtomVol = 7.1
FirstIonPot = 7.726
Cp = 0.385
Hvapor = 300.5
Hfusion = 13.14
ThermConduct = 401
Electroneg = 1.9
Choose2Show.Show
End Sub
Private Sub Zn_Click()
Element = "Zinc"
AtomNum = 30
AtomWeight = 65.39
MeltPt = 692.73
BoilPt = 1180
AtomDensity = 7.13
ElectronConfig = "[Ar] 3d10 4s2"
CrystalStrut = "hexagonal"
ElectConduct = 16.9
CovRad = 1.25
AtomRad = 1.53
AtomVol = 9.2
FirstIonPot = 9.394
Cp = 0.388
Hvapor = 115.3
Hfusion = 7.38
ThermConduct = 116
Electroneg = 1.65
Choose2Show.Show
End Sub
Private Sub Ga_Click()
Element = "Gallium"
AtomNum = 31
AtomWeight = 69.723
MeltPt = 302.92
BoilPt = 2478
AtomDensity = 5.91
ElectronConfig = "[Ar] 3d10 4s2p1"
CrystalStrut = "orthorhombic"
ElectConduct = 1.8
CovRad = 1.26
AtomRad = 1.81
AtomVol = 11.8
FirstIonPot = 5.999
Cp = 0.371
Hvapor = 256.06
Hfusion = 5.59
ThermConduct = 40.6
Electroneg = 1.81
Choose2Show.Show
End Sub
Private Sub Ge_Click()
Element = "Germanium"
AtomNum = 32
AtomWeight = 72.61
MeltPt = 1211.5
BoilPt = 3107
AtomDensity = 5.32
ElectronConfig = "[Ar] 3d10 4s2p2"
CrystalStrut = "cubic, face centered"
ElectConduct = 0.000003
CovRad = 1.22
AtomRad = 1.52
AtomVol = 13.6
FirstIonPot = 7.899
Cp = 0.32
Hvapor = 334.3
Hfusion = 31.8
ThermConduct = 59.9
Electroneg = 2.01
Choose2Show.Show
End Sub
Private Sub Arsenic_Click()
Element = "Arsenic"
AtomNum = 33
AtomWeight = 74.9216
MeltPt = 876
BoilPt = 1090
AtomDensity = 5.78
ElectronConfig = "[Ar] 3d10 4s2p3"
CrystalStrut = "rhombohedral"
ElectConduct = 3.8
CovRad = 1.2
AtomRad = 1.33
AtomVol = 13.1
FirstIonPot = 9.81
Cp = 0.33
Hvapor = 32.4
Hfusion = 27.7
ThermConduct = 50
Electroneg = 2.18
Choose2Show.Show
End Sub
Private Sub Se_Click()
Element = "Selenium"
AtomNum = 34
AtomWeight = 78.96
MeltPt = 494
BoilPt = 958
AtomDensity = 4.79
ElectronConfig = "[Ar] 3d10 4s2p4"
CrystalStrut = "hexagonal"
ElectConduct = 8
CovRad = 1.16
AtomRad = 1.22
AtomVol = 16.5
FirstIonPot = 9.752
Cp = 0.32
Hvapor = 26.32
Hfusion = 5.54
ThermConduct = 2.04
Electroneg = 2.55
Choose2Show.Show
End Sub
Private Sub Br_Click()
Element = "Bromine"
AtomNum = 35
AtomWeight = 79.904
MeltPt = 265.95
BoilPt = 331.85
AtomDensity = 3.12
ElectronConfig = "[Ar] 3d10 4s2p5"
CrystalStrut = "orthorhombic"
ElectConduct = 0.000000000000001
CovRad = 1.14
AtomRad = 1.12
AtomVol = 23.5
FirstIonPot = 11.814
Cp = 0.226
Hvapor = 14.725
Hfusion = 5.286
ThermConduct = 0.122
Electroneg = 2.96
Choose2Show.Show
End Sub
Private Sub Kr_Click()
Element = "Krypton"
AtomNum = 36
AtomWeight = 83.8
MeltPt = 116
BoilPt = 120.85
AtomDensity = 3.75
ElectronConfig = "[Ar] 3d10 4s2p6"
CrystalStrut = "cubic, face centered"
ElectConduct = 99999
CovRad = 1.89
AtomRad = 1.03
AtomVol = 32.2
FirstIonPot = 13.999
Cp = 0.248
Hvapor = 9.029
Hfusion = 1.638
ThermConduct = 0.00949
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Rb_Click()
Element = "Rubidium"
AtomNum = 37
AtomWeight = 85.4678
MeltPt = 312.63
BoilPt = 961
AtomDensity = 1.532
ElectronConfig = "[Kr] 5s1"
CrystalStrut = "cubic, body centered"
ElectConduct = 47.8
CovRad = 2.16
AtomRad = 2.98
AtomVol = 55.9
FirstIonPot = 4.177
Cp = 0.363
Hvapor = 69.2
Hfusion = 2.34
ThermConduct = 58.2
Electroneg = 0.82
Choose2Show.Show
End Sub
Private Sub Sr_Click()
Element = "Strontium"
AtomNum = 38
AtomWeight = 87.62
MeltPt = 1042
BoilPt = 1655
AtomDensity = 2.54
ElectronConfig = "[Kr] 5s2"
CrystalStrut = "cubic, face centered"
ElectConduct = 5
CovRad = 1.91
AtomRad = 2.45
AtomVol = 33.7
FirstIonPot = 5.695
Cp = 0.3
Hvapor = 136.9
Hfusion = 8.2
ThermConduct = 35.3
Electroneg = 0.95
Choose2Show.Show
End Sub
Private Sub Y_Click()
Element = "Yttrium"
AtomNum = 39
AtomWeight = 88.9059
MeltPt = 1795
BoilPt = 3611
AtomDensity = 4.47
ElectronConfig = "[Kr] 4d1 5s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.8
CovRad = 1.62
AtomRad = 2.27
AtomVol = 19.8
FirstIonPot = 6.38
Cp = 0.3
Hvapor = 393.3
Hfusion = 17.15
ThermConduct = 17.2
Electroneg = 1.22
Choose2Show.Show
End Sub
Private Sub Zr_Click()
Element = "Zirconium"
AtomNum = 40
AtomWeight = 91.224
MeltPt = 2128
BoilPt = 4682
AtomDensity = 6.51
ElectronConfig = "[Kr] 4d2 5s2"
CrystalStrut = "hexagonal"
ElectConduct = 2.3
CovRad = 1.45
AtomRad = 2.16
AtomVol = 14.1
FirstIonPot = 6.84
Cp = 0.278
Hvapor = 590.5
Hfusion = 21
ThermConduct = 22.7
Electroneg = 1.33
Choose2Show.Show
End Sub
Private Sub Nb_Click()
Element = "Niobium"
AtomNum = 41
AtomWeight = 92.9064
MeltPt = 2742
BoilPt = 5015
AtomDensity = 8.57
ElectronConfig = "[Kr] 4d4 5s1"
CrystalStrut = "cubic, body centered"
ElectConduct = 6.6
CovRad = 1.34
AtomRad = 2.08
AtomVol = 10.8
FirstIonPot = 6.88
Cp = 0.265
Hvapor = 690.1
Hfusion = 26.9
ThermConduct = 53.7
Electroneg = 1.6
Choose2Show.Show
End Sub
Private Sub Mo_Click()
Element = "Molybdenum"
AtomNum = 42
AtomWeight = 95.94
MeltPt = 2896
BoilPt = 4912
AtomDensity = 10.22
ElectronConfig = "[Kr] 4d5 5s1"
CrystalStrut = "cubic, body centered"
ElectConduct = 17.3
CovRad = 1.3
AtomRad = 2.01
AtomVol = 9.4
FirstIonPot = 7.099
Cp = 0.25
Hvapor = 590.4
Hfusion = 36
ThermConduct = 138
Electroneg = 2.16
Choose2Show.Show
End Sub
Private Sub Tc_Click()
Element = "Technetium"
AtomNum = 43
AtomWeight = 98
MeltPt = 2477
BoilPt = 4538
AtomDensity = 11.5
ElectronConfig = "[Kr] 4d5 5s2"
CrystalStrut = "hexagonal"
ElectConduct = 0.001
CovRad = 1.27
AtomRad = 1.95
AtomVol = 8.5
FirstIonPot = 7.28
Cp = 0.24
Hvapor = 502
Hfusion = 23
ThermConduct = 50.6
Electroneg = 1.9
Choose2Show.Show
End Sub
Private Sub Ru_Click()
Element = "Ruthenium"
AtomNum = 44
AtomWeight = 101.07
MeltPt = 2610
BoilPt = 4425
AtomDensity = 12.37
ElectronConfig = "[Kr] 4d7 5s1"
CrystalStrut = "hexagonal"
ElectConduct = 14.9
CovRad = 1.25
AtomRad = 1.89
AtomVol = 8.3
FirstIonPot = 7.37
Cp = 0.238
Hvapor = 567.77
Hfusion = 25.52
ThermConduct = 117
Electroneg = 2.2
Choose2Show.Show
End Sub
Private Sub Rh_Click()
Element = "Rhodium"
AtomNum = 45
AtomWeight = 102.9055
MeltPt = 2236
BoilPt = 3970
AtomDensity = 12.41
ElectronConfig = "[Kr] 4d8 5s1"
CrystalStrut = "cubic, face centered"
ElectConduct = 23
CovRad = 1.25
AtomRad = 1.83
AtomVol = 8.3
FirstIonPot = 7.46
Cp = 0.242
Hvapor = 495.39
Hfusion = 21.76
ThermConduct = 150
Electroneg = 2.28
Choose2Show.Show
End Sub
Private Sub Pd_Click()
Element = "Palladium"
AtomNum = 46
AtomWeight = 106.42
MeltPt = 1825
BoilPt = 3240
AtomDensity = 12#
ElectronConfig = "[Kr] 4d10"
CrystalStrut = "cubic, face centered"
ElectConduct = 10
CovRad = 1.28
AtomRad = 1.79
AtomVol = 8.9
FirstIonPot = 4.34
Cp = 0.244
Hvapor = 393.3
Hfusion = 16.74
ThermConduct = 71.8
Electroneg = 2.2
Choose2Show.Show
End Sub
Private Sub Ag_Click()
Element = "Silver"
AtomNum = 47
AtomWeight = 107.868
MeltPt = 1235.08
BoilPt = 2436
AtomDensity = 10.5
ElectronConfig = "[Kr] 4d10 5s1"
CrystalStrut = "cubic, face centered"
ElectConduct = 62.9
CovRad = 1.34
AtomRad = 1.75
AtomVol = 10.3
FirstIonPot = 7.576
Cp = 0.235
Hvapor = 250.63
Hfusion = 11.3
ThermConduct = 429
Electroneg = 1.93
Choose2Show.Show
End Sub
Private Sub Cd_Click()
Element = "Cadmium"
AtomNum = 48
AtomWeight = 112.41
MeltPt = 594.26
BoilPt = 1040
AtomDensity = 8.65
ElectronConfig = "[Kr] 4d10 5s2"
CrystalStrut = "hexagonal"
ElectConduct = 14.7
CovRad = 1.41
AtomRad = 1.71
AtomVol = 13.1
FirstIonPot = 8.993
Cp = 0.232
Hvapor = 99.87
Hfusion = 6.07
ThermConduct = 96.8
Electroneg = 1.69
Choose2Show.Show
End Sub
Private Sub Indium_Click()
Element = "Indium"
AtomNum = 49
AtomWeight = 114.82
MeltPt = 429.78
BoilPt = 2350
AtomDensity = 7.31
ElectronConfig = "[Kr] 4d10 5s2p1"
CrystalStrut = "tetragonal"
ElectConduct = 3.4
CovRad = 1.44
AtomRad = 2
AtomVol = 15.7
FirstIonPot = 5.786
Cp = 0.233
Hvapor = 226.35
Hfusion = 3.26
ThermConduct = 81.6
Electroneg = 1.78
Choose2Show.Show
End Sub
Private Sub Sn_Click()
Element = "Tin"
AtomNum = 50
AtomWeight = 118.71
MeltPt = 505.12
BoilPt = 2876
AtomDensity = 7.31
ElectronConfig = "[Kr] 4d10 5s2p2"
CrystalStrut = "tetragonal"
ElectConduct = 8.7
CovRad = 1.41
AtomRad = 1.72
AtomVol = 16.3
FirstIonPot = 7.344
Cp = 0.228
Hvapor = 290.37
Hfusion = 7.2
ThermConduct = 66.6
Electroneg = 1.96
Choose2Show.Show
End Sub
Private Sub Sb_Click()
Element = "Antimony"
AtomNum = 51
AtomWeight = 121.76
MeltPt = 903.91
BoilPt = 1860
AtomDensity = 6.69
ElectronConfig = "[Kr] 4d10 5s2p3"
CrystalStrut = "rhombohedral"
ElectConduct = 2.6
CovRad = 1.4
AtomRad = 1.53
AtomVol = 18.4
FirstIonPot = 8.641
Cp = 0.207
Hvapor = 67.97
Hfusion = 19.83
ThermConduct = 24.3
Electroneg = 2.05
Choose2Show.Show
End Sub
Private Sub Te_Click()
Element = "Tellurium"
AtomNum = 52
AtomWeight = 127.6
MeltPt = 722.72
BoilPt = 1261
AtomDensity = 6.24
ElectronConfig = "[Kr] 4d10 5s2p4"
CrystalStrut = "hexagonal"
ElectConduct = 0.0002
CovRad = 1.36
AtomRad = 1.42
AtomVol = 20.5
FirstIonPot = 9.009
Cp = 0.202
Hvapor = 50.63
Hfusion = 17.49
ThermConduct = 2.35
Electroneg = 2.1
Choose2Show.Show
End Sub
Private Sub I_Click()
Element = "Iodine"
AtomNum = 53
AtomWeight = 126.9045
MeltPt = 386.7
BoilPt = 457.5
AtomDensity = 4.93
ElectronConfig = "[Kr] 4d10 5s2p5"
CrystalStrut = "orthorhombic"
ElectConduct = 0.0000000001
CovRad = 1.33
AtomRad = 1.32
AtomVol = 25.7
FirstIonPot = 10.451
Cp = 0.145
Hvapor = 20.9
Hfusion = 7.76
ThermConduct = 0.449
Electroneg = 2.66
Choose2Show.Show
End Sub
Private Sub Xe_Click()
Element = "Xenon"
AtomNum = 54
AtomWeight = 131.29
MeltPt = 161.39
BoilPt = 165.1
AtomDensity = 5.9
ElectronConfig = "[Kr] 4d10 5s2p6"
CrystalStrut = "cubic, face centered"
ElectConduct = 99999
CovRad = 1.31
AtomRad = 1.24
AtomVol = 42.9
FirstIonPot = 12.13
Cp = 0.158
Hvapor = 12.64
Hfusion = 2.3
ThermConduct = 0.00569
Electroneg = 2.6
Choose2Show.Show
End Sub
Private Sub Cs_Click()
Element = "Cesium"
AtomNum = 55
AtomWeight = 132.9054
MeltPt = 301.54
BoilPt = 944
AtomDensity = 1.87
ElectronConfig = "[Xe] 6s1"
CrystalStrut = "cubic, body centered"
ElectConduct = 5.3
CovRad = 2.35
AtomRad = 3.34
AtomVol = 70
FirstIonPot = 3.894
Cp = 0.24
Hvapor = 67.74
Hfusion = 2.092
ThermConduct = 35.9
Electroneg = 0.79
Choose2Show.Show
End Sub
Private Sub Ba_Click()
Element = "Barium"
AtomNum = 56
AtomWeight = 137.33
MeltPt = 1002
BoilPt = 2078
AtomDensity = 3.59
ElectronConfig = "[Xe] 6s2"
CrystalStrut = "cubic, body centered"
ElectConduct = 2.8
CovRad = 1.98
AtomRad = 2.78
AtomVol = 39
FirstIonPot = 5.212
Cp = 0.204
Hvapor = 140.2
Hfusion = 8.01
ThermConduct = 18.4
Electroneg = 0.89
Choose2Show.Show
End Sub
Private Sub La_Click()
Element = "Lanthanum"
AtomNum = 57
AtomWeight = 138.9055
MeltPt = 3737
BoilPt = 1191
AtomDensity = 6.15
ElectronConfig = "[Xe] 5d1 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.9
CovRad = 1.25
AtomRad = 2.74
AtomVol = 22.5
FirstIonPot = 5.58
Cp = 0.19
Hvapor = 399.57
Hfusion = 11.3
ThermConduct = 13.5
Electroneg = 1.1
Choose2Show.Show
End Sub
Private Sub Ce_Click()
Element = "Cerium"
AtomNum = 58
AtomWeight = 140.12
MeltPt = 1071
BoilPt = 3715
AtomDensity = 6.77
ElectronConfig = "[Xe] 4f1 5d1 6s2"
CrystalStrut = "cubic, face centered"
ElectConduct = 1.4
CovRad = 1.65
AtomRad = 2.7
AtomVol = 21
FirstIonPot = 5.47
Cp = 0.19
Hvapor = 313.8
Hfusion = 9.2
ThermConduct = 11.4
Electroneg = 1.12
Choose2Show.Show
End Sub
Private Sub Pr_Click()
Element = "Praseodymium"
AtomNum = 59
AtomWeight = 140.9077
MeltPt = 3785
BoilPt = 1204
AtomDensity = 6.77
ElectronConfig = "[Xe] 4f3 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.5
CovRad = 1.65
AtomRad = 2.67
AtomVol = 20.8
FirstIonPot = 5.42
Cp = 0.193
Hvapor = 332.63
Hfusion = 10.04
ThermConduct = 12.5
Electroneg = 1.13
Choose2Show.Show
End Sub
Private Sub Nd_Click()
Element = "Neodymium"
AtomNum = 60
AtomWeight = 144.24
MeltPt = 1294
BoilPt = 3347
AtomDensity = 7.01
ElectronConfig = "[Xe] 4f4 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.6
CovRad = 1.64
AtomRad = 2.64
AtomVol = 20.6
FirstIonPot = 5.49
Cp = 0.19
Hvapor = 283.68
Hfusion = 10.88
ThermConduct = 16.5
Electroneg = 1.14
Choose2Show.Show
End Sub
Private Sub Pm_Click()
Element = "Promethium"
AtomNum = 61
AtomWeight = 145
MeltPt = 1315
BoilPt = 3273
AtomDensity = 7.22
ElectronConfig = "[Xe] 4f5 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 2
CovRad = 1.63
AtomRad = 2.62
AtomVol = 22.4
FirstIonPot = 5.55
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 17.9
Electroneg = 1.13
Choose2Show.Show
End Sub
Private Sub Sm_Click()
Element = "Samarium"
AtomNum = 62
AtomWeight = 150.36
MeltPt = 1347
BoilPt = 2067
AtomDensity = 7.52
ElectronConfig = "[Xe] 4f6 6s2"
CrystalStrut = "rhombohedral"
ElectConduct = 1.1
CovRad = 1.62
AtomRad = 2.59
AtomVol = 19.9
FirstIonPot = 5.63
Cp = 0.197
Hvapor = 191.63
Hfusion = 11.09
ThermConduct = 13.3
Electroneg = 1.17
Choose2Show.Show
End Sub
Private Sub Eu_Click()
Element = "Europium"
AtomNum = 63
AtomWeight = 151.964
MeltPt = 1095
BoilPt = 1800
AtomDensity = 5.24
ElectronConfig = "[Xe] 4f7 6s2"
CrystalStrut = "cubic, body centered"
ElectConduct = 1.1
CovRad = 1.85
AtomRad = 2.56
AtomVol = 28.9
FirstIonPot = 5.67
Cp = 0.182
Hvapor = 175.73
Hfusion = 10.46
ThermConduct = 13.9
Electroneg = 1.2
Choose2Show.Show
End Sub
Private Sub Gd_Click()
Element = "Gadolinium"
AtomNum = 64
AtomWeight = 157.25
MeltPt = 1585
BoilPt = 3545
AtomDensity = 7.9
ElectronConfig = "[Xe] 4f7 5d1 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 0.8
CovRad = 1.61
AtomRad = 2.54
AtomVol = 19.9
FirstIonPot = 6.15
Cp = 0.236
Hvapor = 311.71
Hfusion = 15.48
ThermConduct = 10.6
Electroneg = 1.2
Choose2Show.Show
End Sub
Private Sub Tb_Click()
Element = "Terbium"
AtomNum = 65
AtomWeight = 158.9253
MeltPt = 1629
BoilPt = 3500
AtomDensity = 8.23
ElectronConfig = "[Xe] 4f9 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 0.9
CovRad = 1.59
AtomRad = 2.51
AtomVol = 19.2
FirstIonPot = 5.86
Cp = 0.18
Hvapor = 99999
Hfusion = 99999
ThermConduct = 11.1
Electroneg = 1.1
Choose2Show.Show
End Sub
Private Sub Dy_Click()
Element = "Dysprosium"
AtomNum = 66
AtomWeight = 162.5
MeltPt = 1685
BoilPt = 2840
AtomDensity = 8.55
ElectronConfig = "[Xe] 4f10 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.1
CovRad = 1.59
AtomRad = 2.49
AtomVol = 19
FirstIonPot = 5.93
Cp = 0.173
Hvapor = 230
Hfusion = 11.06
ThermConduct = 10.7
Electroneg = 1.22
Choose2Show.Show
End Sub
Private Sub Ho_Click()
Element = "Holmium"
AtomNum = 67
AtomWeight = 164.9303
MeltPt = 1747
BoilPt = 2968
AtomDensity = 8.8
ElectronConfig = "[Xe] 4f11 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.1
CovRad = 1.58
AtomRad = 2.47
AtomVol = 18.7
FirstIonPot = 6.02
Cp = 0.165
Hvapor = 251.04
Hfusion = 17.15
ThermConduct = 16.2
Electroneg = 1.23
Choose2Show.Show
End Sub
Private Sub Er_Click()
Element = "Erbium"
AtomNum = 68
AtomWeight = 167.26
MeltPt = 1802
BoilPt = 3140
AtomDensity = 9.07
ElectronConfig = "[Xe] 4f12 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.2
CovRad = 1.57
AtomRad = 2.45
AtomVol = 18.4
FirstIonPot = 6.101
Cp = 0.168
Hvapor = 292.88
Hfusion = 17.15
ThermConduct = 14.3
Electroneg = 1.24
Choose2Show.Show
End Sub
Private Sub Tm_Click()
Element = "Thulium"
AtomNum = 69
AtomWeight = 168.9342
MeltPt = 1818
BoilPt = 2223
AtomDensity = 9.32
ElectronConfig = "[Xe] 4f13 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.3
CovRad = 1.56
AtomRad = 2.42
AtomVol = 18.1
FirstIonPot = 6.184
Cp = 0.16
Hvapor = 191
Hfusion = 16.8
ThermConduct = 16.8
Electroneg = 1.25
Choose2Show.Show
End Sub
Private Sub Yb_Click()
Element = "Ytterbium"
AtomNum = 70
AtomWeight = 173.04
MeltPt = 1092
BoilPt = 1469
AtomDensity = 6.97
ElectronConfig = "[Xe] 4f14 6s2"
CrystalStrut = "cubic, face centered"
ElectConduct = 3.7
CovRad = 1.7
AtomRad = 2.4
AtomVol = 24.8
FirstIonPot = 6.254
Cp = 0.155
Hvapor = 128
Hfusion = 7.7
ThermConduct = 34.9
Electroneg = 1.1
Choose2Show.Show
End Sub
Private Sub Lu_Click()
Element = "Lutetium"
AtomNum = 71
AtomWeight = 174.967
MeltPt = 1936
BoilPt = 3668
AtomDensity = 9.84
ElectronConfig = "[Xe] 4f14 5d1 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 1.5
CovRad = 1.56
AtomRad = 2.25
AtomVol = 17.8
FirstIonPot = 5.43
Cp = 0.15
Hvapor = 355
Hfusion = 18.6
ThermConduct = 16.4
Electroneg = 1.27
Choose2Show.Show
End Sub
Private Sub Hf_Click()
Element = "Hafnium"
AtomNum = 72
AtomWeight = 178.49
MeltPt = 2504
BoilPt = 4875
AtomDensity = 13.31
ElectronConfig = "[Xe] 4f14 5d2 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 3.4
CovRad = 1.44
AtomRad = 2.16
AtomVol = 13.6
FirstIonPot = 6.65
Cp = 0.14
Hvapor = 661.07
Hfusion = 21.76
ThermConduct = 23
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Ta_Click()
Element = "Tantalum"
AtomNum = 73
AtomWeight = 180.9479
MeltPt = 3293
BoilPt = 5730
AtomDensity = 16.65
ElectronConfig = "[Xe] 4f14 5d3 6s2"
CrystalStrut = "cubic, body centered"
ElectConduct = 8.1
CovRad = 1.34
AtomRad = 2.09
AtomVol = 10.9
FirstIonPot = 7.89
Cp = 0.14
Hvapor = 737
Hfusion = 36
ThermConduct = 57.5
Electroneg = 1.5
Choose2Show.Show
End Sub
Private Sub W_Click()
Element = "Tungsten"
AtomNum = 74
AtomWeight = 183.84
MeltPt = 3695
BoilPt = 5825
AtomDensity = 19.3
ElectronConfig = "[Xe] 4f14 5d3 6s2"
CrystalStrut = "cubic, body centered"
ElectConduct = 18.2
CovRad = 1.3
AtomRad = 2.02
AtomVol = 9.53
FirstIonPot = 7.98
Cp = 0.13
Hvapor = 422.58
Hfusion = 35.4
ThermConduct = 174
Electroneg = 2.36
Choose2Show.Show
End Sub
Private Sub Re_Click()
Element = "Rhenium"
AtomNum = 75
AtomWeight = 186.207
MeltPt = 3455
BoilPt = 5870
AtomDensity = 21#
ElectronConfig = "[Xe] 4f14 5d5 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 5.8
CovRad = 1.28
AtomRad = 1.97
AtomVol = 8.85
FirstIonPot = 7.88
Cp = 0.137
Hvapor = 707.1
Hfusion = 33.05
ThermConduct = 47.9
Electroneg = 1.9
Choose2Show.Show
End Sub
Private Sub Os_Click()
Element = "Osmium"
AtomNum = 76
AtomWeight = 190.23
MeltPt = 3300
BoilPt = 5300
AtomDensity = 22.6
ElectronConfig = "[Xe] 4f14 5d6 6s2"
CrystalStrut = "hexagonal"
ElectConduct = 12.3
CovRad = 1.26
AtomRad = 1.92
AtomVol = 8.43
FirstIonPot = 8.7
Cp = 0.13
Hvapor = 627.6
Hfusion = 29.29
ThermConduct = 87.6
Electroneg = 2.2
Choose2Show.Show
End Sub
Private Sub Ir_Click()
Element = "Iridium"
AtomNum = 77
AtomWeight = 192.22
MeltPt = 2720
BoilPt = 4700
AtomDensity = 22.6
ElectronConfig = "[Xe] 4f14 5d7 6s2"
CrystalStrut = "cubic, face centered"
ElectConduct = 21.3
CovRad = 1.27
AtomRad = 1.87
AtomVol = 8.54
FirstIonPot = 9.1
Cp = 0.13
Hvapor = 563.58
Hfusion = 26.36
ThermConduct = 147
Electroneg = 2.2
Choose2Show.Show
End Sub
Private Sub Pt_Click()
Element = "Platinum"
AtomNum = 78
AtomWeight = 195.08
MeltPt = 4100
BoilPt = 2042.1
AtomDensity = 21.45
ElectronConfig = "[Xe] 4f14 5d8 6s2"
CrystalStrut = "cubic, face centered"
ElectConduct = 9.4
CovRad = 1.3
AtomRad = 1.83
AtomVol = 9.1
FirstIonPot = 9
Cp = 0.13
Hvapor = 510.45
Hfusion = 19.66
ThermConduct = 71.6
Electroneg = 2.28
Choose2Show.Show
End Sub
Private Sub Au_Click()
Element = "Gold"
AtomNum = 79
AtomWeight = 196.9665
MeltPt = 1337.58
BoilPt = 3130
AtomDensity = 19.3
ElectronConfig = "[Xe] 4f14 5d10 6s1"
CrystalStrut = "cubic, face centered"
ElectConduct = 48.8
CovRad = 1.34
AtomRad = 1.79
AtomVol = 10.2
FirstIonPot = 9.225
Cp = 0.128
Hvapor = 324.43
Hfusion = 12.36
ThermConduct = 317
Electroneg = 2.54
Choose2Show.Show
End Sub
Private Sub Hg_Click()
Element = "Mercury"
AtomNum = 80
AtomWeight = 200.59
MeltPt = 234.31
BoilPt = 629.88
AtomDensity = 13.55
ElectronConfig = "[Xe] 4f14 5d10 6s2"
CrystalStrut = "rhombohedral"
ElectConduct = 1
CovRad = 1.49
AtomRad = 1.76
AtomVol = 14.8
FirstIonPot = 10.437
Cp = 0.14
Hvapor = 59.3
Hfusion = 2.292
ThermConduct = 8.34
Electroneg = 2
Choose2Show.Show
End Sub
Private Sub Tl_Click()
Element = "Thallium"
AtomNum = 81
AtomWeight = 204.383
MeltPt = 577
BoilPt = 1746
AtomDensity = 11.85
ElectronConfig = "[Xe] 4f14 5d10 6s2p1"
CrystalStrut = "hexagonal"
ElectConduct = 5.6
CovRad = 1.48
AtomRad = 2.08
AtomVol = 17.2
FirstIonPot = 6.108
Cp = 0.129
Hvapor = 162.09
Hfusion = 4.27
ThermConduct = 46.1
Electroneg = 2.04
Choose2Show.Show
End Sub
Private Sub Pb_Click()
Element = "Lead"
AtomNum = 82
AtomWeight = 207.2
MeltPt = 600.65
BoilPt = 2023
AtomDensity = 11.35
ElectronConfig = "[Xe] 4f14 5d10 6s2p2"
CrystalStrut = "cubic, face centered"
ElectConduct = 4.8
CovRad = 1.47
AtomRad = 1.81
AtomVol = 18.3
FirstIonPot = 7.416
Cp = 0.129
Hvapor = 177.9
Hfusion = 4.77
ThermConduct = 35.3
Electroneg = 2.33
Choose2Show.Show
End Sub
Private Sub Bi_Click()
Element = "Bismuth"
AtomNum = 83
AtomWeight = 208.9804
MeltPt = 544.59
BoilPt = 1837
AtomDensity = 9.75
ElectronConfig = "[Xe] 4f14 5d10 6s2p3"
CrystalStrut = "rhombohedral"
ElectConduct = 0.9
CovRad = 1.46
AtomRad = 1.63
AtomVol = 21.3
FirstIonPot = 7.289
Cp = 0.122
Hvapor = 179
Hfusion = 11
ThermConduct = 7.87
Electroneg = 2.02
Choose2Show.Show
End Sub
Private Sub Po_Click()
Element = "Polonium"
AtomNum = 84
AtomWeight = 209
MeltPt = 527
BoilPt = 99999
AtomDensity = 9.3
ElectronConfig = "[Xe] 4f14 5d10 6s2p4"
CrystalStrut = "monoclinic"
ElectConduct = 0.7
CovRad = 1.53
AtomRad = 1.53
AtomVol = 22.7
FirstIonPot = 8.42
Cp = 99999
Hvapor = 120
Hfusion = 13
ThermConduct = 20
Electroneg = 2
Choose2Show.Show
End Sub
Private Sub At_Click()
Element = "Astatine"
AtomNum = 85
AtomWeight = 210
MeltPt = 575
BoilPt = 610
AtomDensity = 99999
ElectronConfig = "[Xe] 4f14 5d10 6s2p5"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 1.47
AtomRad = 1.43
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 30
Hfusion = 12
ThermConduct = 1.7
Electroneg = 2.2
Choose2Show.Show
End Sub
Private Sub Rn_Click()
Element = "Radon"
AtomNum = 86
AtomWeight = 222
MeltPt = 202
BoilPt = 211.4
AtomDensity = 9.73
ElectronConfig = "[Xe] 4f14 5d10 6s2p6"
CrystalStrut = "cubic, face centered"
ElectConduct = 99999
CovRad = 99999
AtomRad = 1.34
AtomVol = 50.5
FirstIonPot = 10.748
Cp = 0.094
Hvapor = 16.4
Hfusion = 2.9
ThermConduct = 0.00364
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Fr_Click()
Element = "Francium"
AtomNum = 87
AtomWeight = 223
MeltPt = 300
BoilPt = 950
AtomDensity = 99999
ElectronConfig = "[Rn] 7s1"
CrystalStrut = "cubic, body centered"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 64
Hfusion = 2.1
ThermConduct = 15
Electroneg = 0.7
Choose2Show.Show
End Sub
Private Sub Ra_Click()
Element = "Radium"
AtomNum = 88
AtomWeight = 226
MeltPt = 973
BoilPt = 1413
AtomDensity = 5#
ElectronConfig = "[Rn] 7s2"
CrystalStrut = "cubic, body centered"
ElectConduct = 1
CovRad = 99999
AtomRad = 99999
AtomVol = 45.2
FirstIonPot = 5.279
Cp = 0.094
Hvapor = 136.82
Hfusion = 8.37
ThermConduct = 18.6
Electroneg = 0.89
Choose2Show.Show
End Sub
Private Sub Ac_Click()
Element = "Actinium"
AtomNum = 89
AtomWeight = 227
MeltPt = 1324
BoilPt = 3470
AtomDensity = 10.07
ElectronConfig = "[Rn] 6d1 7s2"
CrystalStrut = "cubic, face centered"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 22.5
FirstIonPot = 5.17
Cp = 0.12
Hvapor = 99999
Hfusion = 99999
ThermConduct = 12
Electroneg = 1.1
Choose2Show.Show
End Sub
Private Sub Th_Click()
Element = "Thorium"
AtomNum = 90
AtomWeight = 232.0381
MeltPt = 2028
BoilPt = 5060
AtomDensity = 11.72
ElectronConfig = "[Rn] 6d2 7s2"
CrystalStrut = "cubic, face centered"
ElectConduct = 7.1
CovRad = 1.65
AtomRad = 99999
AtomVol = 19.9
FirstIonPot = 6.08
Cp = 0.113
Hvapor = 543.92
Hfusion = 15.65
ThermConduct = 54
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Pa_Click()
Element = "Protactinium"
AtomNum = 91
AtomWeight = 231.0359
MeltPt = 1845
BoilPt = 4300
AtomDensity = 15.4
ElectronConfig = "[Rn] 5f2 6d1 7s2"
CrystalStrut = "orthorhombic"
ElectConduct = 5.6
CovRad = 99999
AtomRad = 99999
AtomVol = 15
FirstIonPot = 5.88
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 47
Electroneg = 1.5
Choose2Show.Show
End Sub
Private Sub U_Click()
Element = "Uranium"
AtomNum = 92
AtomWeight = 238.029
MeltPt = 1408
BoilPt = 4407
AtomDensity = 18.95
ElectronConfig = "[Rn] 5f3 6d1 7s2"
CrystalStrut = "orthorhombic"
ElectConduct = 3.6
CovRad = 1.42
AtomRad = 99999
AtomVol = 12.5
FirstIonPot = 6.05
Cp = 0.12
Hvapor = 422.58
Hfusion = 15.48
ThermConduct = 27.6
Electroneg = 1.38
Choose2Show.Show
End Sub
Private Sub Np_Click()
Element = "Neptunium"
AtomNum = 93
AtomWeight = 237
MeltPt = 912
BoilPt = 4175
AtomDensity = 20.2
ElectronConfig = "[Rn] 5f4 6d1 7s2"
CrystalStrut = "orthorhombic"
ElectConduct = 0.8
CovRad = 99999
AtomRad = 99999
AtomVol = 21.1
FirstIonPot = 6.19
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 6.3
Electroneg = 1.36
Choose2Show.Show
End Sub
Private Sub Pu_Click()
Element = "Plutonium"
AtomNum = 94
AtomWeight = 244
MeltPt = 913
BoilPt = 3505
AtomDensity = 19.84
ElectronConfig = "[Rn] 5f6 7s2"
CrystalStrut = "monoclinic"
ElectConduct = 0.7
CovRad = 1.08
AtomRad = 99999
AtomVol = 12.32
FirstIonPot = 6.06
Cp = 0.13
Hvapor = 99999
Hfusion = 99999
ThermConduct = 6.74
Electroneg = 1.28
Choose2Show.Show
End Sub
Private Sub Am_Click()
Element = "Americium"
AtomNum = 95
AtomWeight = 243
MeltPt = 1449
BoilPt = 2880
AtomDensity = 13.7
ElectronConfig = "[Rn] 5f7 7s2"
CrystalStrut = "hexagonal"
ElectConduct = 0.7
CovRad = 99999
AtomRad = 99999
AtomVol = 20.8
FirstIonPot = 6
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 10
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Cm_Click()
Element = "Curium"
AtomNum = 96
AtomWeight = 247
MeltPt = 1620
BoilPt = 99999
AtomDensity = 13.5
ElectronConfig = "[Rn] 5f7 6d1 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 18.3
FirstIonPot = 6.02
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 10
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Bk_Click()
Element = "Berkelium"
AtomNum = 97
AtomWeight = 247
MeltPt = 99999
BoilPt = 99999
AtomDensity = 14
ElectronConfig = "[Rn] 5f9 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 6.23
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 10
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Cf_Click()
Element = "Californium"
AtomNum = 98
AtomWeight = 251
MeltPt = 1170
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f10 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 6.3
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 10
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Es_Click()
Element = "Einsteinium"
AtomNum = 99
AtomWeight = 252
MeltPt = 1130
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f11 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 6.42
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 10
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Fm_Click()
Element = "Fermium"
AtomNum = 100
AtomWeight = 257
MeltPt = 1800
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f12 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 6.5
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 10
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Md_Click()
Element = "Mendelevium"
AtomNum = 101
AtomWeight = 258
MeltPt = 1100
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f13 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 6.58
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 10
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub No_Click()
Element = "Nobelium"
AtomNum = 102
AtomWeight = 259
MeltPt = 1100
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 6.65
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 10
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Lr_Click()
Element = "Lawrencium"
AtomNum = 103
AtomWeight = 262
MeltPt = 1900
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d1 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 10
Electroneg = 1.3
Choose2Show.Show
End Sub
Private Sub Rf_Click()
Element = "Rutherfordium"
AtomNum = 104
AtomWeight = 261
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d2 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Db_Click()
Element = "Dubnium"
AtomNum = 105
AtomWeight = 262
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d3 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Sg_Click()
Element = "Seaborgium"
AtomNum = 106
AtomWeight = 263
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d4 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Bh_Click()
Element = "Bohrium"
AtomNum = 107
AtomWeight = 264
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d5 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Hs_Click()
Element = "Hassium"
AtomNum = 108
AtomWeight = 265
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d6 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Mt_Click()
Element = "Meitnerium"
AtomNum = 109
AtomWeight = 268
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d7 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Ds_Click()
Element = "Darmstadtium"
AtomNum = 110
AtomWeight = 269
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d8 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Rg_Click()
Element = "Roentgenium"
AtomNum = 111
AtomWeight = 272
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d9 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Cn_Click()
Element = "Copernicium"
AtomNum = 112
AtomWeight = 277
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d10 7s2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Uut_Click()
Element = "Ununtrium"
AtomNum = 113
AtomWeight = 99999
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "Unknown"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Fl_Click()
Element = "Flerovium"
AtomNum = 114
AtomWeight = 285
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "[Rn] 5f14 6d10 7s2p2"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Uup_Click()
Element = "Ununpentium"
AtomNum = 115
AtomWeight = 99999
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "Unknown"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Lv_Click()
Element = "Livermorium"
AtomNum = 116
AtomWeight = 99999
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "Unknown"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Uus_Click()
Element = "Ununseptium"
AtomNum = 117
AtomWeight = 99999
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "Unknown"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub
Private Sub Uuo_Click()
Element = "Ununoctium"
AtomNum = 118
AtomWeight = 99999
MeltPt = 99999
BoilPt = 99999
AtomDensity = 99999
ElectronConfig = "Unknown"
CrystalStrut = "unknown"
ElectConduct = 99999
CovRad = 99999
AtomRad = 99999
AtomVol = 99999
FirstIonPot = 99999
Cp = 99999
Hvapor = 99999
Hfusion = 99999
ThermConduct = 99999
Electroneg = 99999
Choose2Show.Show
End Sub

Private Sub Closer_Click()
PeriodicTable.Hide
Unload PeriodicTable
End Sub
