Attribute VB_Name = "Import_STL"

'--------------
' Point 3D
'-------------
Public Type Point3
    X As Double
    Y As Double
    Z As Double
End Type


'--------------
' 3 couleurs
'-------------
Public Type CouleurRVB
  Rouge As Long
  Vert As Long
  Bleu As Long
End Type


Public Type Facette
    NmbVertex As Integer
    NmbNormal As Integer
    Normal() As Point3
    Vertex() As Point3
End Type


Public Type Element3D
    Name As String
    File As String
    STL_def As Facette
    Origine As Point3
    Type_axe As Integer
    Vecteur As Point3
    Valeur_axe As Double
    Color As CouleurRVB
    MiniAxe As Long
    MaxiAxe As Long
End Type

Public Type Machine3D
    Name As String
    Type As Integer   '1=SCARA 2=POLYMORPH
    NB_axe As Integer ' Nb axe sur la machine
    Element() As Element3D
    Accessoire As Integer ' Présence d'un accessoire (pince ou autre)
End Type

Public Type position
    Join(6) As Double
End Type


Public Machine As Machine3D


Global Position_precedente As position
' Global Position_courante As position
Global Position_suivante As position
' Affiche les coordonnées
Public Sub Affiche_coord()
    Call Sub_Affiche_coord(Pt0, 0)
    Call Sub_Affiche_coord(Vx, 3)
    Call Sub_Affiche_coord(Vy, 6)
    Call Sub_Affiche_coord(Vz, 9)
End Sub

' Affiche les coordonnees d'un point dans le controle ROBOT_SIMUL_FRM.AxeAbsolu(Index)
Sub Sub_Affiche_coord(P1 As Point3, Index As Integer)
    ROBOT_SIMUL_FRM.AxeAbsolu(Index) = Round(P1.X, 3)
    ROBOT_SIMUL_FRM.AxeAbsolu(Index + 1) = Round(P1.Y, 3)
    ROBOT_SIMUL_FRM.AxeAbsolu(Index + 2) = Round(P1.Z, 3)
End Sub

'Recuperation du File STL
Sub ChargeFichier(File As String, Mesh As Facette, Optional Facteur As Integer = 1)

Dim Donnéeslues As String
Dim Poi As Point3
Loading "Loading " & File


On Error GoTo fin

Open File For Input As #1   ' Ouvre le File en lecture
    Do While Not EOF(1) ' Cherche la fin du File.
  
    Line Input #1, Donnéeslues  ' Lit une ligne de données.
    
    If InStr(1, Donnéeslues, "facet normal") Then
        Mesh.NmbNormal = Mesh.NmbNormal + 1
         ReDim Preserve Mesh.Normal(Mesh.NmbNormal)
        Call Decodage(Donnéeslues, Poi)
        Mesh.Normal(Mesh.NmbNormal - 1) = Poi
        Line Input #1, Donnéeslues  ' Lit une ligne de données.
    End If
    
    If InStr(1, Donnéeslues, "outer loop") Then
    Mesh.NmbVertex = Mesh.NmbVertex + 3
    ReDim Preserve Mesh.Vertex(Mesh.NmbVertex)
        Line Input #1, Donnéeslues  ' 1Er vertex
        Call Decodage(Donnéeslues, Poi)
        Mesh.Vertex(Mesh.NmbVertex - 3) = Poi
        Line Input #1, Donnéeslues  ' 2Eme vertex
        Call Decodage(Donnéeslues, Poi)
        Mesh.Vertex(Mesh.NmbVertex - 2) = Poi
        Line Input #1, Donnéeslues  ' 3Eme vertex
        Call Decodage(Donnéeslues, Poi)
        Mesh.Vertex(Mesh.NmbVertex - 1) = Poi
    End If
    
  Loop ' fin boucle traitement File
Close #1

Exit Sub


fin:
    MsgBox Err.Description, 16, "Error #" & Err.Number

    
End Sub

Public Function Decodage(ByVal Ligne As String, Poi As Point3) As Boolean
Dim Chaine As String
Dim Pt As Point3

Decodage = False

    Chaine = Replace(Ligne, Chr(9), Chr(32)) 'Remplace les tabulations par des espaces
    'Chaine = LTrim(Chaine) ' Suprime les espaces de gauche.
    'Chaine = mReplaceCharacter(Chr(13), "", Chaine) 'Supprime Chr(13)
    'Chaine = mReplaceCharacter(Chr(10), "", Chaine) 'Supprime Chr(10)
    'Chaine = RTrim(Chaine) ' Suprime les espaces de droite.
    
    If InStr(1, Ligne, "normal") Then
       Chaine = LTrim((TokRightRight(Chaine, "l")))
       Pt.X = Val(TokLeftLeft(Chaine, " "))
       Chaine = LTrim(TokLeftRight(Chaine, " "))
       
       Pt.Y = Val(TokLeftLeft(Chaine, " "))
       Chaine = LTrim(TokLeftRight(Chaine, " "))
    
       Pt.Z = Val(Chaine)
       ' normaliser le vecteur
       'non necessaire
       ' Call VecteurUnitaire(Pt)
       Poi = Pt
       Decodage = True
    End If
    
    If InStr(1, Ligne, "vertex") Then
       Chaine = LTrim((TokRightRight(Chaine, "x")))
       Poi.X = Val(TokLeftLeft(Chaine, " "))
       Chaine = LTrim(TokLeftRight(Chaine, " "))
       
       Poi.Y = Val(TokLeftLeft(Chaine, " "))
       Chaine = LTrim(TokLeftRight(Chaine, " "))
    
       Poi.Z = Val(Chaine)
       Decodage = True
    End If
  
FinSub:

 
End Function

'Transformation d'une couleur long en RVB
Public Function RVB(Couleur_long As Long) As CouleurRVB
Dim TempColor As CouleurRVB
    
    TempColor.Bleu = Int(Couleur_long / 65536)
    TempColor.Vert = Int((Couleur_long - (65536 * TempColor.Bleu)) / 256)
    TempColor.Rouge = Couleur_long - (65536 * TempColor.Bleu + 256 * TempColor.Vert)

RVB = TempColor
  
End Function
