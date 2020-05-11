Attribute VB_Name = "Funciones"
Option Compare Text

Public Function Marshall(PaFi, Creatinina, Tension_Arterial_Sistolica, Optional pH = 7.4)
On Error Resume Next

Dim P, C, TA
If PaFi = "" Or Creatinina = "" Or Tension_Arterial_Sistolica = "" Then
Marshall = "N/A"
Else

   If PaFi >= 400 Then P = 0
   If PaFi >= 301 And PaFi < 400 Then P = 1
   If PaFi >= 201 And PaFi < 301 Then P = 2
   If PaFi >= 101 And PaFi < 201 Then P = 3
   If PaFi < 101 Then P = 4


    If Creatinina < 1.4 Then C = 0
    If Creatinina >= 1.4 And Creatinina < 1.9 Then C = 1
    If Creatinina >= 1.9 And Creatinina < 3.7 Then C = 2
    If Creatinina >= 3.7 And Creatinina < 4.9 Then C = 3
    If Creatinina >= 4.9 Then C = 4

 If Tension_Arterial_Sistolica >= 90 Then TA = 0

If Tension_Arterial_Sistolica < 90 Then
    If pH <> "" Then
        If pH >= 7.3 Then TA = 1
        If pH >= 7.2 And pH < 7.3 Then TA = 3
        If pH < 7.2 Then TA = 4
    End If
Else
    TA = 1
End If
Marshall = P + C + TA
End If

End Function


Public Function SIRS(FreC, FreR, Temp, Leu, Optional CO2 = "")
On Error Resume Next

Dim FC, FR, T, L, C
If FreC <> "" Then FC = 1 Else FC = 0:      If FreR <> "" Then FR = 1 Else FR = 0
If Temp <> "" Then T = 1 Else T = 0:       If Leu <> "" Then L = 1 Else L = 0
If FC + FR + T + L < 2 Then
SIRS = 0
Else
    If FreC > 90 Then FC = 1
    If FreC <= 90 Then FC = 0

    If FreR >= 21 Then FR = 1
    If FreR < 21 Then FR = 0

    If Temp >= 38 Then T = 1
    If Temp >= 36 And Temp < 38 Then T = 0
    If Temp < 36 Then T = 1

    If Leu >= 12 Then L = 1
    If Leu >= 4 And Leu < 12 Then L = 0
    If Leu < 4 Then L = 1

SIRS = FC + FR + T + L
If SIRS < 2 Then SIRS = 0 Else SIRS = 1
End If
End Function


Public Function BISAP(BUN, Estado_mental, SIRS, Edad, Derrame_pleural)
On Error Resume Next

Dim B, EM, SI, Ed, DP
If BUN > 25 Then B = 1 Else B = 0
If Estado_mental <= 13 Then EM = 1 Else EM = 0
If SIRS > 0 Then SI = 1 Else SI = 0
If Edad > 60 Then Ed = 1 Else Ed = 0
If Derrame_pleural > 0 Then DP = 1 Else DP = 0

BISAP = B + EM + SI + Ed + DP

End Function

Public Function APACHE_II(Temperatura, Pres_art_med, Fre_C, _
    Fre_R, PaO2, p_H, Sodio, Potadio, Creatinina, Hematocrito, Leucocitos, Glasgow, _
    Edad, Optional Diabetes = 0, Optional Cirrosis = 0, Optional Ins_Cardiaca = 0, _
    Optional Neumopatia = 0, Optional ERC_V = 0, Optional InmunoDef = 0, Optional QX = 0)
    
On Error Resume Next
Dim Temp, PAM, FC, FR, PO, pH, Na, K, Cr, Hto, Leu, ECG, Ed, DM2, HAS, EHep, ICC, EPOC, ERC, Idef, CHP

If Temperatura <> "" And Pres_art_med <> "" And Fre_C <> "" And Fre_R <> "" And PaO2 <> "" _
    And p_H <> "" And Sodio <> "" And Potadio <> "" And Creatinina <> "" And Hematocrito <> "" And Leucocitos <> "" And Glasgow <> "" And Edad <> "" Then

If Diabetes + Cirrosis + Ins_Cardiaca + Neumopatia + ERC_V + InmunoDef + QX = 0 Then
CHP = 0
Else
Select Case QX
    Case 0: CHP = 5:     ' No quirœrgico
    Case 1: CHP = 2      ' Cirug’a electiva
    Case 2: CHP = 5      ' Cirug’a urgencia
    Case Else: CHP = 0  'Error
End Select
End If
'================Temperatura===============================
Select Case Temperatura
  Case Is >= 41: Temp = 4
  Case 39 To 40.9: Temp = 3
  Case 38.5 To 39: Temp = 1
  Case 36 To 38.4: Temp = 0
  Case 34 To 35.9: Temp = 1
  Case 32 To 33.9: Temp = 2
  Case 30 To 31.2: Temp = 3
  Case Is <= 29.9: Temp = 4
End Select
'  ======= Otra forma de temperatura=======================
'  If Temperatura >= 41 Then Temp = 4
'  If Temperatura >= 39 And Temperatura < 41 Then Temp = 3
'  If Temperatura >= 38.5 And Temperatura < 39 Then Temp = 1
'  If Temperatura >= 36 And Temperatura < 38.5 Then Temp = 0
'  If Temperatura >= 34 And Temperatura < 36 Then Temp = 1
'  If Temperatura >= 32 And Temperatura < 34 Then Temp = 2
'  If Temperatura >= 30 And Temperatura < 32 Then Temp = 3
'  If Temperatura < 30 Then Temp = 4
'==========================================================
'================TAM===============================

Select Case Pres_art_med
  Case Is > 159: PAM = 4
  Case 130 To 159: PAM = 3
  Case 110 To 129.9: PAM = 2
  Case 70 To 109.9: PAM = 0
  Case 50 To 69.9: PAM = 2
  Case Is < 50: PAM = 4
End Select
'  If Pres_art_med >= 159 Then PAM = 4
'  If Pres_art_med >= 130 And Pres_art_med < 159 Then PAM = 3
'  If Pres_art_med >= 110 And Pres_art_med < 130 Then PAM = 2
'  If Pres_art_med >= 70 And Pres_art_med < 110 Then PAM = 0
'  If Pres_art_med >= 50 And Pres_art_med < 70 Then PAM = 2
'  If Pres_art_med < 50 Then PAM = 4
'==========================================================
Select Case Fre_C
  Case Is > 179: FC = 4
  Case 140 To 178: FC = 3
  Case 110 To 139: FC = 2
  Case 70 To 109: FC = 0
  Case 55 To 69: FC = 2
  Case 40 To 54: FC = 3
  Case Is < 40: FC = 4
End Select

'If Fre_C >= 179 Then FC = 4
'If Fre_C >= 140 And Fre_C < 179 Then FC = 3
'If Fre_C >= 110 And Fre_C < 140 Then FC = 2
'If Fre_C >= 70 And Fre_C < 110 Then FC = 0
'If Fre_C >= 55 And Fre_C < 70 Then FC = 2
'If Fre_C >= 40 And Fre_C < 55 Then FC = 3
'If Fre_C < 40 Then FC = 4
'==========================================================
Select Case Fre_R
  Case Is > 49: FR = 4
  Case 35 To 49: FR = 3
  Case 25 To 34: FR = 1
  Case 12 To 24: FR = 0
  Case 10 To 11: FR = 1
  Case 6 To 9: FR = 2
  Case Is < 6: FR = 4
End Select
'If Fre_R >= 49 Then FR = 4
'If Fre_R >= 35 And Fre_R < 49 Then FR = 3
'If Fre_R >= 25 And Fre_R < 35 Then FR = 1
'If Fre_R >= 12 And Fre_R < 25 Then FR = 0
'If Fre_R >= 10 And Fre_R < 12 Then FR = 1
'If Fre_R >= 6 And Fre_R < 10 Then FR = 2
'If Fre_R < 6 Then FR = 4
''==========================================================
Select Case PaO2
  Case Is > 70: PO = 0
  Case 61 To 70: PO = 1
  Case 56 To 60: PO = 3
  Case Is < 56: PO = 4
End Select
'If PaO2 >= 70 Then PO = 0
'If PaO2 >= 61 And PaO2 < 70 Then PO = 1
'If PaO2 >= 56 And PaO2 < 61 Then PO = 3
'If PaO2 < 56 Then PO = 4
''==========================================================
Select Case p_H
  Case Is > 7.69: pH = 4
  Case 7.6 To 7.69: pH = 3
  Case 7.5 To 7.59: pH = 1
  Case 7.33 To 7.49: pH = 0
  Case 7.25 To 7.32: pH = 2
  Case 7.15 To 7.24: pH = 3
  Case Is < 7.15: pH = 4
End Select
'If p_H >= 7.69 Then pH = 4
'If p_H >= 7.6 And p_H < 7.69 Then pH = 3
'If p_H >= 7.5 And p_H < 7.6 Then pH = 1
'If p_H >= 7.33 And p_H < 7.5 Then pH = 0
'If p_H >= 7.25 And p_H < 7.33 Then pH = 2
'If p_H >= 7.15 And p_H < 7.25 Then pH = 3
'If p_H < 7.15 Then pH = 4
''==========================================================
Select Case Sodio
  Case Is > 179: Na = 4
  Case 160 To 179: Na = 3
  Case 155 To 159.9: Na = 2
  Case 150 To 154.9: Na = 1
  Case 130 To 149.9: Na = 0
  Case 120 To 129.9: Na = 2
  Case 111 To 119.9: Na = 3
  Case Is < 110: Na = 4
End Select
'If Sodio >= 179 Then Na = 4
'If Sodio >= 160 And Sodio < 179 Then Na = 3
'If Sodio >= 155 And Sodio < 160 Then Na = 2
'If Sodio >= 150 And Sodio < 155 Then Na = 1
'If Sodio >= 130 And Sodio < 150 Then Na = 0
'If Sodio >= 120 And Sodio < 130 Then Na = 2
'If Sodio >= 111 And Sodio < 120 Then Na = 3
'If Sodio < 111 Then Na = 4
''==========================================================
Select Case Potadio
  Case Is > 6.9: K = 4
  Case 6 To 6.9: K = 3
  Case 5.5 To 5.9: K = 2
  Case 3.5 To 5.4: K = 1
  Case 3 To 3.4: K = 0
  Case 2.5 To 2.9: K = 2
  Case Is < 2.5: K = 4
End Select
'If Potadio >= 6.9 Then K = 4
'If Potadio >= 6 And Potadio < 6.9 Then K = 3
'If Potadio >= 5.5 And Potadio < 6 Then K = 1
'If Potadio >= 3.5 And Potadio < 5.5 Then K = 0
'If Potadio >= 3 And Potadio < 3.5 Then K = 1
'If Potadio >= 2.5 And Potadio < 3 Then K = 2
'If Potadio < 2.5 Then K = 4
''==========================================================
Select Case Creatinina
  Case Is > 3.4: Cr = 4
  Case 2 To 3.3: Cr = 3
  Case 1.5 To 1.9: Cr = 2
  Case 0.6 To 1.4: Cr = 0
  Case Is < 0.6: Cr = 2
End Select
'If Creatinina >= 3.4 Then Cr = 4
'If Creatinina >= 2 And Creatinina < 3.4 Then Cr = 3
'If Creatinina >= 1.5 And Creatinina < 2 Then Cr = 2
'If Creatinina >= 0.6 And Creatinina < 1.5 Then Cr = 0
'If Creatinina < 0.6 Then Cr = 2
''==========================================================
Select Case Hematocrito
  Case Is > 6.9: K = 4
  Case 6 To 6.9: K = 3
  Case 5.5 To 5.9: K = 2
  Case 3.5 To 5.4: K = 1
  Case 3 To 3.4: K = 0
  Case 2.5 To 2.9: K = 2
  Case Is < 2.5: K = 4
End Select

'If Hematocrito >= 59.9 Then Hto = 4
'If Hematocrito >= 50 And Hematocrito < 59.9 Then Hto = 2
'If Hematocrito >= 46 And Hematocrito < 50 Then Hto = 1
'If Hematocrito >= 30 And Hematocrito < 46 Then Hto = 0
'If Hematocrito >= 20 And Hematocrito < 30 Then Hto = 2
'If Hematocrito < 20 Then Hto = 4
''==========================================================
Select Case Potadio
  Case Is > 6.9: K = 4
  Case 6 To 6.9: K = 3
  Case 5.5 To 5.9: K = 2
  Case 3.5 To 5.4: K = 1
  Case 3 To 3.4: K = 0
  Case 2.5 To 2.9: K = 2
  Case Is < 2.5: K = 4
End Select

'If Leucocitos >= 39.9 Then Leu = 4
'If Leucocitos >= 20 And Leucocitos < 39.9 Then Leu = 2
'If Leucocitos >= 15 And Leucocitos < 20 Then Leu = 1
'If Leucocitos >= 3 And Leucocitos < 15 Then Leu = 0
'If Leucocitos >= 1 And Leucocitos < 3 Then Leu = 2
'If Leucocitos < 1 Then Leu = 4
''==========================================================

ECG = 15 - Glasgow
'==========================================================
Select Case Potadio
  Case Is > 6.9: K = 4
  Case 6 To 6.9: K = 3
  Case 5.5 To 5.9: K = 2
  Case 3.5 To 5.4: K = 1
  Case 3 To 3.4: K = 0
  Case 2.5 To 2.9: K = 2
  Case Is < 2.5: K = 4
End Select

'    If Edad < 45 Then Ed = 0
'    If Edad >= 45 And Edad < 55 Then Ed = 2
'    If Edad >= 55 And Edad < 65 Then Ed = 3
'    If Edad >= 65 And Edad <= 74 Then Ed = 5
'    If Edad > 74 Then Ed = 6
'

APACHE_II = Temp + PAM + FC + FR + PO + pH + Na + K + Cr + Hto + Leu + ECG + Ed + CHP
Else
APACHE_II = "N/A"
End If

End Function


