VERSION 5.00
Begin VB.Form PreVisu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PreView Robot"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   FillColor       =   &H00808080&
   ForeColor       =   &H00808080&
   Icon            =   "PreVisu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   4800
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1020
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4575
   End
   Begin VB.PictureBox PicPrevisu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   120
      Width           =   4605
   End
End
Attribute VB_Name = "PreVisu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim File_ini As String

Dim Machine_tempo As Machine3D


Private Sub Form_Load()
    'Me.Show
    'PreVisu.SetFocus
ROBOT_SIMUL_FRM.MessageLOG.Cls

    SetWindowPos PreVisu.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    Me.Top = ROBOT_SIMUL_FRM.Top + 150
     Me.Left = ROBOT_SIMUL_FRM.Left + ROBOT_SIMUL_FRM.Width / 3
       

    ' permet d'initialiser la grille
    xm = 45
ym = 0
zm = 45

Zoom = 0.0018

PosX = 0
PosY = 0

' Initialisation du controle picturebox en opengl
LoadGL Me.PicPrevisu
    
Call PrevisuRobot(Me.PicPrevisu)
     
File_ini = App.Path + "\Robot_simul.ini"

Call Init_Liste_robot(File_ini)
'init sur premier robot par defaut
If File_robot = "" Then
File_robot = App.Path + "\Robot_def\" + mfncGetFromIni("Robot1", "File", File_ini)
End If

End Sub

Sub Init_Liste_robot(File As String)
Dim nb_robot As Integer
Dim I As Integer
Dim Name_Item As String

    nb_robot = Val(mfncGetFromIni("ROBOT_SIMUL", "NB_robot", File))

List1.Clear

For I = 1 To nb_robot
Name_Item = "Robot" & I

List1.AddItem (mfncGetFromIni(Name_Item, "Name", File))

    
Next I


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim J As Integer

'Charge le robot
 ROBOT_SIMUL_FRM.MessageLOG.Cls

 LoadGL ROBOT_SIMUL_FRM.Pic


    Call ROBOT_SIMUL_FRM.init_robot(File_robot)
    
    Call DessineRobot(ROBOT_SIMUL_FRM.Pic, False, 1, False)
    Call GETPoint
    Call Affiche_coord
    
    For J = 1 To 2
        ROBOT_SIMUL_FRM.PictureVoyant(J).Visible = False
    Next
    

'Call ROBOT_SIMUL_FRM.Pic_Paint
End Sub

Private Sub List1_Click()
Dim Name_Item As String
Dim I As Integer

xm = 45
ym = 0
zm = 45

Zoom = 0.0018

PosX = 0
PosY = 0


I = List1.ListIndex + 1
Name_Item = "Robot" & I

   File_robot = App.Path + "\Robot_def\" + mfncGetFromIni(Name_Item, "File", File_ini)
   
    Call init_previsu_robot(File_robot)
    Call PrevisuRobot(PicPrevisu)
    
End Sub


'Initialisation du robot
Sub init_previsu_robot(File_Def_machine As String)
'Dim File_STL As String
Dim I As Integer


'Chargement caractéristique robot
Call Charger_robot(File_Def_machine, Machine_tempo)



'Recuperation de la definition geométrique via File STL ascii
For I = 0 To UBound(Machine_tempo.Element)
    Call Reinit_Element(Machine_tempo.Element(I))
    'File_STL = "bras" & I & "_o.stl"
    Call ChargeFichier(App.Path + "\Robot_def\" + Machine_tempo.Element(I).File, Machine_tempo.Element(I).STL_def)
    'Machine.Element(I).Color = RVB(QBColor(I + 1))
Next

End Sub

Sub PrevisuRobot(Pict As PictureBox)

Dim MMatrix(1 To 16) As Double
Dim T As Integer
Dim Z As Integer
Dim I As Integer


On Error Resume Next



'---DEBUT :

    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    glMatrixMode mmModelView

    
'---POINT DE VUE : Camera

    gluLookAt 20, 0, 0, 0, 0, 0, 0, 0, 20




'---PETITE INITALISATION :
        
    'glMaterialiv GL_FRONT_AND_BACK, mprAmbientAndDiffuse, GL_SPECULAR
    'glMateriali GL_FRONT_AND_BACK, GL_SHININESS, 100
     glViewport 0, 0, Pict.ScaleWidth, Pict.ScaleHeight

    glScalef Zoom, Zoom, Zoom
    
        glRotatef xm, 0, 1, 0  ' rotation en y
        glRotatef ym, 1, 0, 0  ' rotation en x
        glRotatef zm, 0, 0, 1  ' rotation en z


    T = 50
    Z = 20
    glTranslatef 0, 0, 0
    glBegin bmLines
    glColor3d 1, 1, 1
    glEnable GL_DEPTH_TEST
    For I = -T To T Step 10
        glVertex3f I * Z, -T * Z, 0
        glVertex3f I * Z, T * Z, 0
    Next
    For I = -T To T Step 10
        glVertex3f -T * Z, I * Z, 0
        glVertex3f T * Z, I * Z, 0
    Next
    glEnd

 
'---AFFICHAGE
' Init des traits cachés
glEnable GL_DEPTH_TEST
glEnable GL_POLYGON_OFFSET_FILL
glPolygonOffset 1, 2

glPushMatrix



'Affichage element
For I = 0 To UBound(Machine_tempo.Element)
    'Debug.Print Machine.Element(I).Valeur_axe
     Call Affiche_Element(Machine_tempo.Element(I), 1)
Next
'gluLookAt Pt0.X * Zoom + 20, Pt0.Y * Zoom, Pt0.Z * Zoom, 0, 0, 0, 0, 0, 1




glPopMatrix



SwapBuffers Pict.hdc
'reinitialise
glLoadIdentity
End Sub

Private Sub PicPrevisu_Paint()
        LoadGL PicPrevisu
        Call PrevisuRobot(PicPrevisu)
End Sub
