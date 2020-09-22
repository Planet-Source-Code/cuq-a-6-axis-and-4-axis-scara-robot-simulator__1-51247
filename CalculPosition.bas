Attribute VB_Name = "CalculPosition"
Option Explicit

'Public Const Pi = 3.14159265358979
Public PI As Double
Public RADTODEG As Double
Public DEGTORAD As Double

Function calcul_position_polymorh(PtA As Point3, A As Double, B As Double, C As Double, position_calculée As position) As Boolean
Dim position_calculée1 As position
Dim position_calculée2 As position
Dim Poignet As Point3
Dim LeVecteur As Point3
Dim P0 As Point3
Dim VPX As Point3
Dim VPY As Point3
Dim VPZ As Point3

Dim P1 As Point3
Dim P2 As Point3
Dim P3 As Point3

Dim retour As Integer
Dim Pabs As Point3
Dim Pdiff As Point3
Dim qdiff As Point3
Dim dist As Double

calcul_position_polymorh = False


Poignet.X = -Machine.Element(7).Origine.X - Machine.Element(6).Origine.X
Poignet.Y = -Machine.Element(7).Origine.Y - Machine.Element(6).Origine.Y
Poignet.Z = -Machine.Element(7).Origine.Z - Machine.Element(6).Origine.Z


' calcul la position du poignet

Call CalculSuivantABC(Poignet, PtA, A, B, C)
Debug.Print "Poignet"
Debug.Print " X " & Poignet.X & " Y " & Poignet.Y & " Z " & Poignet.Z

retour = resoudreQ1Q2Q3(Poignet, position_calculée1, position_calculée2)

Select Case retour
   
' une solution pour join1
Case 0
   ROBOT_SIMUL_FRM.MessageLOG.CurrentY = 0
   ROBOT_SIMUL_FRM.MessageLOG.Cls
   ROBOT_SIMUL_FRM.MessageLOG.Print "INACCESSIBLE POINT !!!"

   ROBOT_SIMUL_FRM.PictureVoyant(2).Visible = True
   Exit Function
   
' une solution pour join2-join3
Case 1

' deux solutions pour join2-join3
Case 2

'erreur inconnu
Case Else

End Select

'Selection de la position Haut/Bas
If ROBOT_SIMUL_FRM.PositionHaute Then
    position_calculée = position_calculée2
Else
    position_calculée = position_calculée1
End If



'-------------------------------------------------------------------
'  Calcul de join 4
'
'  Calcul avec le vecteur Z projeté sur plan normal a VPX du bras
'     après mise en position suivant Q1 Q2 Q3
'-----------------------------------------------------------------
Call GETPointPoignet(position_calculée, P0, VPX, VPY, VPZ)

LeVecteur.X = 0
LeVecteur.Y = 0
LeVecteur.Z = 1
' calcul le vecteur Z a obtenir
Call CalculSuivantABC(LeVecteur, Pabs, A, B, C)
Debug.Print " LeVecteur   | " & Format(LeVecteur.X, "#,###0.0000") & " | " & Format(LeVecteur.Y, "#,###0.0000") & " | " & Format(LeVecteur.Z, "#,###0.0000") & " | "

' projection de LeVecteur sur le plan
 Call VecSub(LeVecteur, Pabs, Pdiff)
 dist = Dot(VPX, Pdiff)
' projection avec VPX=normale plan
 Call SubVect(Pdiff, VPX, dist, qdiff)
    
Debug.Print " VPZ         | " & Format(VPZ.X, "#,###0.0000") & " | " & Format(VPZ.Y, "#,###0.0000") & " | " & Format(VPZ.Z, "#,###0.0000") & " | "
Debug.Print " qdiff       | " & Format(qdiff.X, "#,###0.0000") & " | " & Format(qdiff.Y, "#,###0.0000") & " | " & Format(qdiff.Z, "#,###0.0000") & " | "

position_calculée1.Join(4) = AngleVect(VPZ, qdiff, VPX)
    
    ' Calcul avec + ou - PI
If position_calculée1.Join(4) + 180 > Machine.Element(5).MaxiAxe Then
    position_calculée2.Join(4) = position_calculée1.Join(4) - 180
    Else
    position_calculée2.Join(4) = position_calculée1.Join(4) + 180
End If
    

'-------------------------------------------------
'  Choix de join 4
'-------------------------------------------------
If ROBOT_SIMUL_FRM.PositionGauche Then
    position_calculée.Join(4) = position_calculée2.Join(4)
Else
    position_calculée.Join(4) = position_calculée1.Join(4)
End If


'-------------------------------------------------
'  Calcul de join 5
'  Calcul avec le vecteur Z projeté sur plan normal a VPY du bras
'     après mise en position suivant Q1 Q2 Q3 Q4
'-------------------------------------------------
'Vecteur apres determination Join4
Call GETPointPoignet(position_calculée, P0, VPX, VPY, VPZ)
Debug.Print " VPY         | " & Format(VPY.X, "#,###0.0000") & " | " & Format(VPY.Y, "#,###0.0000") & " | " & Format(VPY.Z, "#,###0.0000") & " | "

' Vecteur du join5 en position 0
LeVecteur.X = 0
LeVecteur.Y = 0
LeVecteur.Z = 1
' calcul le vecteur Z a obtenir
Call CalculSuivantABC(LeVecteur, Pabs, A, B, C)
Debug.Print " LeVecteur   | " & Format(LeVecteur.X, "#,###0.0000") & " | " & Format(LeVecteur.Y, "#,###0.0000") & " | " & Format(LeVecteur.Z, "#,###0.0000") & " | "

' projection de LeVecteur (Z) sur le plan
Call VecSub(LeVecteur, Pabs, Pdiff)
dist = Dot(VPY, Pdiff)
' projection avec VPY=normale plan
 Call SubVect(Pdiff, VPY, dist, qdiff)
 
Debug.Print " qdiff       | " & Format(qdiff.X, "#,###0.0000") & " | " & Format(qdiff.Y, "#,###0.0000") & " | " & Format(qdiff.Z, "#,###0.0000") & " | "
Debug.Print " VPZ         | " & Format(VPZ.X, "#,###0.0000") & " | " & Format(VPZ.Y, "#,###0.0000") & " | " & Format(VPZ.Z, "#,###0.0000") & " | "


 position_calculée.Join(5) = AngleVect(VPZ, qdiff, VPY)


'----------------------------------------------------------------
'  Calcul de join 6
'  Calcul avec le vecteur X projeté sur plan normal a VPZ du bras
'     après mise en position suivant Q1 Q2 Q3 Q4 Q5
'-----------------------------------------------------------------
'Vecteur apres determination Join5
Call GETPointPoignet(position_calculée, P0, VPX, VPY, VPZ)
Debug.Print " VPX        | " & Format(VPX.X, "#,###0.0000") & " | " & Format(VPX.Y, "#,###0.0000") & " | " & Format(VPX.Z, "#,###0.0000") & " | "


LeVecteur.X = 1
LeVecteur.Y = 0
LeVecteur.Z = 0

' calcul le vecteur VPX a obtenir
Call CalculSuivantABC(LeVecteur, Pabs, A, B, C)
Debug.Print " LeVecteur   | " & Format(LeVecteur.X, "#,###0.0000") & " | " & Format(LeVecteur.Y, "#,###0.0000") & " | " & Format(LeVecteur.Z, "#,###0.0000") & " | "

' projection de LeVecteur (x) sur le plan
Call VecSub(LeVecteur, Pabs, Pdiff)
dist = Dot(VPZ, Pdiff)
' projection avec VPZ=normale plan
 Call SubVect(Pdiff, VPZ, dist, qdiff)
Debug.Print " VPY         | " & Format(VPY.X, "#,###0.0000") & " | " & Format(VPY.Y, "#,###0.0000") & " | " & Format(VPY.Z, "#,###0.0000") & " | "
Debug.Print " VPZ         | " & Format(VPZ.X, "#,###0.0000") & " | " & Format(VPZ.Y, "#,###0.0000") & " | " & Format(VPZ.Z, "#,###0.0000") & " | "
Debug.Print " LeVecteur   | " & Format(LeVecteur.X, "#,###0.0000") & " | " & Format(LeVecteur.Y, "#,###0.0000") & " | " & Format(LeVecteur.Z, "#,###0.0000") & " | "

Debug.Print " VPX        | " & Format(VPX.X, "#,###0.0000") & " | " & Format(VPX.Y, "#,###0.0000") & " | " & Format(VPX.Z, "#,###0.0000") & " | "
Debug.Print " qdiff       | " & Format(qdiff.X, "#,###0.0000") & " | " & Format(qdiff.Y, "#,###0.0000") & " | " & Format(qdiff.Z, "#,###0.0000") & " | "

 position_calculée.Join(6) = AngleVect(VPX, qdiff, VPZ)


calcul_position_polymorh = True

End Function

Function calcul_position_scara(PtA As Point3, C As Double, position_calculée As position) As Boolean
Dim position_calculée1 As position
Dim position_calculée2 As position
Dim Poignet As Point3
Dim LeVecteur As Point3
Dim P0 As Point3
Dim VPX As Point3
Dim VPY As Point3
Dim VPZ As Point3

Dim P1 As Point3
Dim P2 As Point3
Dim P3 As Point3

Dim retour As Integer
Dim Pabs As Point3
Dim Pdiff As Point3
Dim qdiff As Point3
Dim dist As Double

calcul_position_scara = False


' calcul la position du point pivot

Poignet = PtA
'Debug.Print "Poignet"
'Debug.Print " X " & Poignet.X & " Y " & Poignet.Y & " Z " & Poignet.Z

retour = resoudreQ2Q3(Poignet, position_calculée1, position_calculée2)

Select Case retour
   
' une solution pour join1
Case 0
   ROBOT_SIMUL_FRM.MessageLOG.CurrentY = 0
   ROBOT_SIMUL_FRM.MessageLOG.Cls
   ROBOT_SIMUL_FRM.MessageLOG.Print " INACCESSIBLE POINT"
   ROBOT_SIMUL_FRM.PictureVoyant(2).Visible = True
   Exit Function
   
' une solution pour join2-join3
Case 1

' deux solutions pour join2-join3
Case 2

'erreur inconnu
Case Else

End Select


'-------------------------------------------------
'  Choix de suivant position
'-------------------------------------------------
If ROBOT_SIMUL_FRM.PositionGauche Then
    position_calculée.Join(1) = position_calculée2.Join(1)
    position_calculée.Join(2) = position_calculée2.Join(2)
Else
    position_calculée.Join(1) = position_calculée1.Join(1)
    position_calculée.Join(2) = position_calculée1.Join(2)
End If

'-------------------------------------------------
'  autre valeur
'-------------------------------------------------
position_calculée.Join(3) = PtA.Z - (Machine.Element(1).Origine.Z + Machine.Element(2).Origine.Z + Machine.Element(3).Origine.Z + Machine.Element(4).Origine.Z)
position_calculée.Join(4) = C - (position_calculée.Join(1) + position_calculée.Join(2))


calcul_position_scara = True

End Function
Function resoudreQ2Q3(O As Point3, position_c1 As position, position_c2 As position) As Integer

Dim Q1 As Double
Dim Q2 As Double
Dim Q3 As Double
Dim D As Double
Dim D1 As Double
Dim A1 As Double
Dim A2 As Double
Dim D4 As Double
Dim K1 As Double
Dim K2 As Double
Dim K3 As Double
Dim P1 As Point3
Dim P2 As Point3
Dim P3 As Point3
Dim P4 As Point3
Dim AngleCompensation As Double

resoudreQ2Q3 = 0

'Debug.Print "O X " & O.x & " Y " & O.y & " Z " & O.Z


A1 = (Machine.Element(1).Origine.X)
D1 = (Machine.Element(1).Origine.Y)


P1 = Machine.Element(1).Origine
P1.Z = 0
Call VecAdd(P1, Machine.Element(2).Origine, 1, P2)
P2.Z = 0
A2 = Distance(P1, P2)

Call VecAdd(P2, Machine.Element(3).Origine, 1, P3)
P3.Z = 0
D4 = Distance(P2, P3)

D = (((O.Y - D1) ^ 2 + (Sqr(O.X ^ 2) - A1) ^ 2 - A2 ^ 2 - D4 ^ 2)) / (2 * A2 * D4)
If Abs(D) > 1 Then
'Pas de solution car point inaccessible par le robot
   resoudreQ2Q3 = 0
   Exit Function
End If

'Premiere solution sur Q3
Q3 = Atan2(Sqr(1 - D ^ 2), D) * RADTODEG

K1 = A2 + (D4 * (Sin(Q3 * DEGTORAD)))
K2 = D4 * Cos(Q3 * DEGTORAD)

Q2 = Atan2((K2 * (O.Y - D1)) + (K1 * (Sqr(O.X ^ 2) - A1)), (K1 * (O.Y - D1)) - (K2 * (Sqr(O.X ^ 2) - A1))) * RADTODEG

position_c1.Join(1) = Q2
position_c1.Join(2) = 90 - Q3




resoudreQ2Q3 = resoudreQ2Q3 + 11
'Deuxieme solution sur Q3
Q3 = Atan2(-1 * Sqr(1 - D ^ 2), D) * RADTODEG  ' - 11.482
K1 = A2 + (D4 * (Sin(Q3 * DEGTORAD)))
K2 = D4 * Cos(Q3 * DEGTORAD)
Q2 = Atan2((K2 * (O.Y - D1)) + (K1 * (Sqr(O.X ^ 2) - A1)), (K1 * (O.Y - D1)) - (K2 * (Sqr(O.X ^ 2) - A1))) * RADTODEG

'Q2 = Atan2(K2 * (O.Y - D1) + K1 * (Sqr(O.X ^ 2) - A1), K1 * (O.Z - D1) - K2 * (Sqr(O.X ^ 2) - A1)) * RADTODEG

position_c2.Join(1) = Q2
position_c2.Join(2) = -270 - Q3


resoudreQ2Q3 = resoudreQ2Q3 + 1


End Function
Function resoudreQ1Q2Q3(O As Point3, position_c1 As position, position_c2 As position) As Integer

Dim Q1 As Double
Dim Q2 As Double
Dim Q3 As Double
Dim D As Double
Dim D1 As Double
Dim A1 As Double
Dim A2 As Double
Dim D4 As Double
Dim K1 As Double
Dim K2 As Double
Dim K3 As Double
Dim P1 As Point3
Dim P2 As Point3
Dim P3 As Point3
Dim P4 As Point3
Dim AngleCompensation As Double

resoudreQ1Q2Q3 = 0

'Debug.Print "O X " & O.x & " Y " & O.y & " Z " & O.Z

' Calcul venant de http://eavr.u-strasbg.fr/library/teaching/robotics/chap2/siframes.htm
If O.Y <> 0 Then
    position_c1.Join(1) = Atan2(O.X, O.Y) * RADTODEG
Else
    'Oy = 0
    position_c1.Join(1) = 0
End If
position_c2 = position_c1


D1 = (Machine.Element(1).Origine.Z + Machine.Element(2).Origine.Z)
A1 = (Machine.Element(1).Origine.X + Machine.Element(2).Origine.X)

'AngleCompensation =11.482
'A2 = 570
'D4 = 653
Call VecAdd(Machine.Element(1).Origine, Machine.Element(2).Origine, 1, P1)
P1.Y = 0
Call VecAdd(P1, Machine.Element(3).Origine, 1, P2)
P2.Y = 0
A2 = Distance(P1, P2)

Call VecAdd(P2, Machine.Element(4).Origine, 1, P4)
Call VecAdd(P4, Machine.Element(5).Origine, 1, P3)
P3.Y = 0
P4.Y = 0
D4 = Distance(P2, P3)

AngleCompensation = 180 - Angle3Pt(P2, P3, P4)


D = (((O.Z - D1) ^ 2 + (Sqr((O.X ^ 2 + O.Y ^ 2)) - A1) ^ 2 - A2 ^ 2 - D4 ^ 2)) / (2 * A2 * D4)
If Abs(D) > 1 Then
'Pas de solution car point inaccessible par le robot
   resoudreQ1Q2Q3 = 0
   Exit Function
End If

'Premiere solution sur Q3
Q3 = Atan2(Sqr(1 - D ^ 2), D) * RADTODEG '- 11.482

K1 = A2 + (D4 * (Sin(Q3 * DEGTORAD)))
K2 = D4 * Cos(Q3 * DEGTORAD)

Q2 = Atan2((K2 * (O.Z - D1)) + (K1 * (Sqr(O.X ^ 2 + O.Y ^ 2) - A1)), (K1 * (O.Z - D1)) - (K2 * (Sqr(O.X ^ 2 + O.Y ^ 2) - A1))) * RADTODEG
'Q2 = Atan2((K1 * (O.Z - D1)) - (K2 * (Sqr(O.x ^ 2 + O.y ^ 2) - A1)), (K2 * (O.Z - D1)) + (K1 * (Sqr(O.x ^ 2 + O.y ^ 2) - A1)))
position_c1.Join(3) = -1 * (Q3 - AngleCompensation)
position_c2.Join(2) = -1 * (Q2 - 90)
resoudreQ1Q2Q3 = resoudreQ1Q2Q3 + 11
'Deuxieme solution sur Q3
Q3 = Atan2(-1 * Sqr(1 - D ^ 2), D) * RADTODEG  ' - 11.482
K1 = A2 + (D4 * (Sin(Q3 * DEGTORAD)))
K2 = D4 * Cos(Q3 * DEGTORAD)

Q2 = Atan2(K2 * (O.Z - D1) + K1 * (Sqr(O.X ^ 2 + O.Y ^ 2) - A1), K1 * (O.Z - D1) - K2 * (Sqr(O.X ^ 2 + O.Y ^ 2) - A1)) * RADTODEG
position_c2.Join(3) = -1 * (Q3 - AngleCompensation)
position_c1.Join(2) = -1 * (Q2 - 90)
resoudreQ1Q2Q3 = resoudreQ1Q2Q3 + 1

resoudreQ1Q2Q3 = 2

End Function
'
Sub CalculSuivantABC(Pt As Point3, Origine As Point3, XangleIn As Double, YangleIn As Double, ZangleIn As Double)

Dim Point_0 As Point3
Dim Point_1 As Point3
Dim Point_2 As Point3
Dim Point_3 As Point3

Dim Xangle As Double
Dim Yangle As Double
Dim Zangle As Double
    
Dim X(2, 2) As Double 'X Matrix
Dim Y(2, 2) As Double 'Y Matrix
Dim Z(2, 2) As Double 'Z Matrix
   
        
        Point_0 = Pt
        
        Xangle = Val(XangleIn) * DEGTORAD
        Yangle = Val(YangleIn) * DEGTORAD
        Zangle = Val(ZangleIn) * DEGTORAD
        
        
     ' Autour de X
    '³   1           0           0       ³
    '³                                   ³
    '³   0        cos(Zan)   -sin(Zan)   ³
    '³                                   ³
    '³   0        sin(Zan)    cos(Zan)   ³
    X(2, 2) = Cos(Xangle) 'X matrice
    X(1, 1) = Cos(Xangle)
    X(2, 1) = Sin(Xangle)
    X(1, 2) = -Sin(Xangle)
    X(0, 0) = 1
    
    ' Autour de Y
    '³  cos(Yan)     0       sin(Yan)   ³
    '³                                  ³
    '³   0           1           0      ³
    '³                                  ³
    '³ -sin(Yan)     0       cos(Yan)   ³
    Y(0, 0) = Cos(Yangle) 'Y matrice
    Y(2, 2) = Cos(Yangle)
    Y(0, 2) = Sin(Yangle)
    Y(2, 0) = -Sin(Yangle)
    Y(1, 1) = 1
    
    'Rotation autour de Z Axis
    '³  cos(Xan)  -sin(Xan)      0  ³
    '³                              ³
    '³  sin(Xan)   cos(Xan)      0  ³
    '³                              ³
    '³   0           0           1  ³
    Z(0, 0) = Cos(Zangle) 'Z matrice
    Z(1, 0) = Sin(Zangle)
    Z(0, 1) = -Sin(Zangle)
    Z(1, 1) = Cos(Zangle)
    Z(2, 2) = 1
            
            
    Call Trans_Matrix(X, Point_0, Point_1)
    Call Trans_Matrix(Y, Point_1, Point_2)
    Call Trans_Matrix(Z, Point_2, Point_3)
                       
    'Debug.Print Point_0.x; Point_0.y; Point_0.Z
    'Debug.Print Point_3.x; Point_3.y; Point_3.Z
     Call VecAdd(Point_3, Origine, 1, Point_2)
            
     Pt = Point_2


End Sub

Sub Trans_Matrix(Mx2() As Double, P1 As Point3, P2 As Point3)
'Calculate the Matrix
 
               P2.X = Mx2(0, 0) * P1.X + Mx2(0, 1) * P1.Y + Mx2(0, 2) * P1.Z
               P2.Y = Mx2(1, 0) * P1.X + Mx2(1, 1) * P1.Y + Mx2(1, 2) * P1.Z
               P2.Z = Mx2(2, 0) * P1.X + Mx2(2, 1) * P1.Y + Mx2(2, 2) * P1.Z
'debug.print P2.Z

End Sub

