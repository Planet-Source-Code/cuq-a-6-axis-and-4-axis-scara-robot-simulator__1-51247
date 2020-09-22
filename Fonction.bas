Attribute VB_Name = "Fonction"
Public Declare Sub SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'INI File Functions
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long


'Angle entre 2 Vecteurs formés de 3 point ( attention ne gère pas de valeu + ou -)

Function Angle3Pt(Po1 As Point3, Po2 As Point3, Po3 As Point3) As Double
Dim VEC1 As Point3
Dim VEC2 As Point3

Call VecSub(Po1, Po2, VEC1)
Call VecSub(Po2, Po3, VEC2)
If Distance(Po3, Po2) = 0 Or Distance(Po1, Po2) = 0 Then
    Angle3Pt = 0
Else
    If Distance(Po1, Po3) = 0 Then
        Angle3Pt = 180
    Else
        Angle3Pt = RADTODEG * (ACOS((VEC1.X * VEC2.X + VEC1.Y * VEC2.Y + VEC1.Z * VEC2.Z) / (Distance(Po3, Po2) * Distance(Po1, Po2))))
    End If
End If
                    
End Function

Function Atan2(ByVal X, ByVal Y)
    'On Error Resume Next
    '0 a PI
    If Y = 0 Then
        If X = 0 Then
            Atan2 = 0
        ElseIf X > 0 Then
            Atan2 = 0 'PI / 2
        Else
            Atan2 = PI '-PI / 2
        End If
    ElseIf X = 0 Then
        If Y > 0 Then
            Atan2 = PI / 2
        Else
            Atan2 = -PI / 2
        End If
    ElseIf X > 0 Then
        Atan2 = Atn(Y / X)
    Else
        Atan2 = (PI - Atn(Abs(Y) / Abs(X))) * Sgn(Y)
    End If
End Function


' Suppression dans une chaine de charactères des éventuels charactères parasites
Function Nettoyage_texte(Chaine As String) As String

  Dim pos_char As Integer
  
    Chaine = Replace(Chaine, Chr(9), Chr(32)) 'Remplace les tabulations par des espaces
    Chaine = LTrim(Chaine) ' Suprime les espaces de gauche.
    Chaine = Replace(Chaine, Chr(13), "") 'Supprime Chr(13)
    Chaine = Replace(Chaine, Chr(10), "") 'Supprime Chr(10)
    Chaine = RTrim(Chaine) ' Suprime les espaces de droite.
    
    ' corrige le ( sans espace avant pour analyse du select case
    pos_char = InStr(1, Chaine, "(")
    If pos_char > 1 Then
        If Mid$(Chaine, (pos_char - 1), 1) <> " " Then
            Chaine = Left(Chaine, pos_char - 1) + " (" + Right(Chaine, Len(Chaine) - pos_char)
        End If
    End If
        
    ' Suppression du numero pour saut dans programme
    Select Case Mid$(Chaine, 1, 1)
        Case "0" To "9"
            If InStr(2, Chaine, ":") Then
             Nettoyage_texte = LTrim(TokLeftRight(Chaine, ":"))
            Else
             Nettoyage_texte = Chaine
            End If
            
        Case Else
            Nettoyage_texte = Chaine
    End Select
End Function

'****************************************************************
' Name: A 'strtok' function for VB
' Description:I wrote four functions to tokenize strings. He
'     re they are...
'The functions work like this TokLeftLeft finds the leftmost token and
'then returns the left part of the string (empty if not there). You
'can figure out the rest. Note that if the token is more than 1 character
'then the function will always return "".
'****************************************************************
Public Function TokLeftLeft(Source As String, token As String) As String

       Dim I As Integer
       TokLeft = Source

              For I = 1 To Len(Source)

                            If Mid(Source, I, 1) = token Then
                                   TokLeftLeft = Left(Source, I - 1)
                                   Exit Function
                            End If

              Next I

End Function
Public Function TokLeftRight(Source As String, token As String) As String

       Dim I As Integer
       TokRight = Source

              For I = 1 To Len(Source)

                            If Mid(Source, I, 1) = token Then
                                   TokLeftRight = Right(Source, Len(Source) - I)
                                   Exit Function
                            End If

              Next I

End Function
Public Function TokRightLeft(Source As String, token As String) As String

       Dim I As Integer
       TokRightLeft = ""

              For I = Len(Source) To 1 Step -1

                            If Mid(Source, I, 1) = token Then
                                   TokRightLeft = Left(Source, I - 1)
                                   Exit Function
                            End If

              Next I

End Function
Public Function TokRightRight(Source As String, token As String) As String

       Dim I As Integer
       TokRightRight = ""

              For I = Len(Source) To 1 Step -1

                            If Mid(Source, I, 1) = token Then
                                   TokRightRight = Right(Source, Len(Source) - I)
                                   Exit Function
                            End If

              Next I

End Function





Function longueur(P1 As Point3) As Double
  longueur = Sqr((P1.X ^ 2) + (P1.Y ^ 2) + (P1.Z ^ 2))
End Function

Function Distance(P1 As Point3, P2 As Point3) As Double
    Distance = Sqr((P2.X - P1.X) ^ 2 + (P2.Y - P1.Y) ^ 2 + (P2.Z - P1.Z) ^ 2)
End Function

Sub VecAdd(P1 As Point3, P2 As Point3, f As Double, P3 As Point3)
 P3.X = P1.X + f * P2.X
 P3.Y = P1.Y + f * P2.Y
 P3.Z = P1.Z + f * P2.Z
End Sub

'Produit Scalaire
Function Dot(p As Point3, q As Point3) As Double
    Dot = p.X * q.X + p.Y * q.Y + p.Z * q.Z
End Function
Sub SubVect(P1 As Point3, P2 As Point3, f As Double, P3 As Point3)
 P3.X = P1.X - P2.X * f
 P3.Y = P1.Y - P2.Y * f
 P3.Z = P1.Z - P2.Z * f
End Sub

Sub VecSub(P1 As Point3, P2 As Point3, P3 As Point3)
 P3.X = P1.X - P2.X
 P3.Y = P1.Y - P2.Y
 P3.Z = P1.Z - P2.Z
End Sub
'Produit vectoriel
Sub VecProd(P1 As Point3, P2 As Point3, P3 As Point3)

Dim P4 As Point3
 

 P4.X = (P1.Y * P2.Z) - (P1.Z * P2.Y)
 P4.Y = (P1.Z * P2.X) - (P1.X * P2.Z)
 P4.Z = (P1.X * P2.Y) - (P1.Y * P2.X)
 P3 = P4
 
End Sub
'récupération du vecteur normal de 3 points
Sub NormVec(P1 As Point3, P2 As Point3, P3 As Point3, Nv As Point3)
 Dim A As Point3
 Dim B As Point3
 A.X = P1.X - P2.X
 A.Y = P1.Y - P2.Y
 A.Z = P1.Z - P2.Z
 B.X = P3.X - P2.X
 B.Y = P3.Y - P2.Y
 B.Z = P3.Z - P2.Z
 Call VecProd(A, B, Nv)
 Call VecteurUnitaire(Nv)
 
End Sub

' transforme un vecteur en vecteur unitaire
Sub VecteurUnitaire(P1 As Point3)
Dim Norm As Double
Norm = Sqr((P1.X) ^ 2 + (P1.Y) ^ 2 + (P1.Z) ^ 2)
If Norm = 0 Then
    
    Exit Sub
End If
    
    P1.X = P1.X / Norm
    P1.Y = P1.Y / Norm
    P1.Z = P1.Z / Norm

End Sub
' Coordonées du point Milieu
Sub PointMillieu(P1 As Point3, P2 As Point3, P3 As Point3)

 P3.X = 0.5 * (P1.X + P2.X)
 P3.Y = 0.5 * (P1.Y + P2.Y)
 P3.Z = 0.5 * (P1.Z + P2.Z)
End Sub
'****************************************************************
' Name: Round
'
' Inputs:DP is the decimal place to round to (0 to 14) e.g
' Round (3.56376, 3) will give the result 3.564
' Round (3.56376, 1) will give the result 3.6
' Round (3.56376, 0) will give the result 4
' Round (3.56376, 2) will give the result 3.56
' Round (1.4999, 3) will give the result 1.5
' Round (1.4899, 2) will give the result 1.49
' Returns:None
' Assumes:None
' Side Effects:None
'
'****************************************************************
Function Round(x1 As Double, DP As Integer) As Double
    Round = Int((x1 * 10 ^ DP) + 0.5) / 10 ^ DP
End Function

Function ACOS(Ang As Double) As Double
    Select Case Val(Ang)
        Case 1
            ACOS = 0 '0
        Case -1
            ACOS = 4 * Atn(1) 'PI
        Case Else
            ACOS = 2 * Atn(1) - Atn(Ang / Sqr(1 - Ang ^ 2))
    End Select
End Function

Function ASIN(Ang As Double) As Double
    Select Case Val(Ang)
        Case 1
            ASIN = 2 * Atn(1)
        Case -1
            ASIN = -2 * Atn(1)
        Case Else
            ASIN = Atn(Ang / Sqr(1 - Ang ^ 2))
    End Select
End Function

'Angle entre 2 Vecteurs formés concourant en 0
Function AngleVect(Po1 As Point3, Po2 As Point3, Normal As Point3) As Double
Dim VEC1 As Point3
Dim VEC2 As Point3
Dim Po3 As Point3
Dim Signe As Double

VEC1 = Po1
VEC2 = Po2
If longueur(VEC1) = 0 Or longueur(VEC2) = 0 Then
    AngleVect = 0
Else
    If Distance(Po1, Po2) = 0 Then
        AngleVect = 0
    Else

        Call VecProd(VEC1, VEC2, Po3)
        Signe = Sgn(Dot(Po3, Normal))
        
        'Debug.Print " Dot/Normal         | " & Signe
        
        'Debug.Print " Po3         | " & Format(Po3.x, "#,###0.0000") & " | " & Format(Po3.y, "#,###0.0000") & " | " & Format(Po3.Z, "#,###0.0000") & " | "
        AngleVect = Signe * RADTODEG * (ACOS((VEC1.X * VEC2.X + VEC1.Y * VEC2.Y + VEC1.Z * VEC2.Z) / (longueur(VEC1) * longueur(VEC2))))
    End If
End If
                    
End Function


