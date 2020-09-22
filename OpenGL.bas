Attribute VB_Name = "OpenGL"
Public HGLRC As Long
Global f As String

Global Yangle As GLfloat
Global Langle As Integer
Global PosX As GLfloat
Global PosY As GLfloat
Global SVGposX As GLfloat
Global SVGposY As GLfloat
Global Zoom As Single

Global Pt0 As Point3
Global Vx As Point3
Global Vy As Point3
Global Vz As Point3


Global xm As Single, ym As Single, zm As Single


Sub SetupPixelFormat(ByVal hdc As Long)
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim PixelFormat As Integer
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 16
    pfd.cDepthBits = 16
    pfd.iLayerType = PFD_MAIN_PLANE
    PixelFormat = ChoosePixelFormat(hdc, pfd)
    If PixelFormat = 0 Then MsgBox ("format de pixels inconnu")
    SetPixelFormat hdc, PixelFormat, pfd
End Sub

Public Sub Lampe()
    Dim aflLightAmbient(4) As GLfloat
    Dim aflLightDiffuse(4) As GLfloat
    Dim aflLightPosition(4) As GLfloat
    Dim aflLightSpecular(4) As GLfloat
    
   
   glDisable glcLighting 'temporarily disable lighting
   glPushMatrix 'push the matrix stack down by one
   
    glEnable GL_LIGHT0
    
    glShadeModel smSmooth
    glEnable glcColorMaterial
    
    


    ' Smooth shading
    'glShadeModel smSmooth
    glShadeModel smFlat
    
    ' Set the clear colour gris
    glClearColor 0.5, 0.5, 0.5, 0 '
    ' Set the clear depth
    glClearDepth 1#
    
    ' Enable Z-buffer
    glEnable glcDepthTest
    ' Set test type
    glDepthFunc cfLEqual
    ' Best perspective correction
    glHint htPerspectiveCorrectionHint, hmNicest
      
    ' Ambient light settings
    aflLightAmbient(0) = 1
    aflLightAmbient(1) = 1
    aflLightAmbient(2) = 0.85
    aflLightAmbient(3) = 1
    ' Diffuse light settings
    aflLightDiffuse(0) = 1
    aflLightDiffuse(1) = 1
    aflLightDiffuse(2) = 1
    aflLightDiffuse(3) = 1
    ' Light position settings
    aflLightPosition(0) = 3000
    aflLightPosition(1) = 3000
    aflLightPosition(2) = 3000
    aflLightPosition(3) = 1
      
    ' Light position Specular
    aflLightSpecular(0) = 0
    aflLightSpecular(1) = 0
    aflLightSpecular(2) = 0
    aflLightSpecular(3) = 1
    
    ' Set the light's ambient and diffuse levels and its position
    glLightfv ltLight0, lpmAmbient, aflLightAmbient(0)
    glLightfv ltLight0, lpmDiffuse, aflLightDiffuse(0)
    glLightfv ltLight0, lpmSpecular, aflLightSpecular(0)
    glLightfv ltLight0, lpmPosition, aflLightPosition(0)
    
    ' Enable light0
    'glEnable glcLight0
    
            glPopMatrix 'pop the matrix stack up by one
    glEnable GL_LIGHTING 're-enable lighting
    
End Sub
Public Sub LoadGL(p As PictureBox)

    
    SetupPixelFormat p.hdc
    HGLRC = wglCreateContext(p.hdc)
    wglMakeCurrent p.hdc, HGLRC
    glEnable glcDepthTest
    glDepthFunc cfLEqual
    

    glMatrixMode mmProjection
    glLoadIdentity
    gluPerspective 10, p.ScaleWidth / p.ScaleHeight, 1, 1000

    
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
    
    Call Lampe
End Sub

Sub Loading(T As String, Optional ByVal Visible = True)

With ROBOT_SIMUL_FRM
    .MessageLOG.Print T
End With
DoEvents

End Sub

Sub DessineRobot(Pict As PictureBox, Grille As Boolean, Render_mode As Integer, Option_tracer As Boolean)

Dim MMatrix(1 To 16) As Double
       

On Error Resume Next



'---DEBUT :

    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    glMatrixMode mmModelView

    
'---POINT DE VUE : Camera
' Vue Standard
If ROBOT_SIMUL_FRM.OptionVue(0).Value = True Then
    gluLookAt 20, 0, 0, 0, 0, 0, 0, 0, 20
End If
' Vue Pince
If ROBOT_SIMUL_FRM.OptionVue(1).Value = True Then
    gluLookAt Pt0.X * Zoom + 20, Pt0.Y * Zoom, Pt0.Z * Zoom, Pt0.X * Zoom, Pt0.Y * Zoom, Pt0.Z * Zoom, 0, 0, 1
End If
'Vue Dessus
If ROBOT_SIMUL_FRM.OptionVue(2).Value = True Then
    gluLookAt 0, 0, 22, 0, 0, 0, -20, 0, 0
End If
'Vue Desous
If ROBOT_SIMUL_FRM.OptionVue(3).Value = True Then
    gluLookAt 0, 0, -22, 0, 0, 0, 20, 0, 0
End If



'---PETITE INITALISATION :
        
    'glMaterialiv GL_FRONT_AND_BACK, mprAmbientAndDiffuse, GL_SPECULAR
    'glMateriali GL_FRONT_AND_BACK, GL_SHININESS, 100
     glViewport 0, 0, Pict.ScaleWidth, Pict.ScaleHeight

    glScalef Zoom, Zoom, Zoom
    
    If ROBOT_SIMUL_FRM.OptionVue(0).Value = True Then
        glRotatef xm, 0, 1, 0  ' rotation en y
        glRotatef ym, 1, 0, 0  ' rotation en x
        glRotatef zm, 0, 0, 1  ' rotation en z
    End If
    
    If ROBOT_SIMUL_FRM.OptionVue(2).Value = True Or ROBOT_SIMUL_FRM.OptionVue(3).Value = True Then
        glTranslatef PosX, PosY, 0
    Else
        glTranslatef 0, 0, PosX ', 0
        glTranslatef PosY, PosY, 0
    End If
    

        

     
     'glViewport -Zoom1x, -Zoom1y, Pict.ScaleWidth + Zoom2x, Pict.ScaleHeight + Zoom2y
     
'---GRILLE :
        
If Grille Then
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
End If
 
'---AFFICHAGE
' Init des traits cachés
glEnable GL_DEPTH_TEST
glEnable GL_POLYGON_OFFSET_FILL
glPolygonOffset 1, 2

glPushMatrix



' Tracer des vecteurs de localisation du point de controle principale
If Option_tracer Then
glBegin bmLines
            glColor3d 1, 0, 0
            glBegin bmLines
            glVertex3f 0, 0, 0
            glVertex3f Pt0.X, Pt0.Y, Pt0.Z
            glEnd
            
            glColor3d 1, 1, 1
            glBegin bmLines
            glVertex3f Pt0.X, Pt0.Y, Pt0.Z
            glVertex3f Pt0.X + Vz.X * 200, Pt0.Y + Vz.Y * 200, Pt0.Z + Vz.Z * 200
            glEnd
            
            glColor3d 0, 1, 0
            glBegin bmLines
            glVertex3f Pt0.X, Pt0.Y, Pt0.Z
            glVertex3f Pt0.X + Vx.X * 200, Pt0.Y + Vx.Y * 200, Pt0.Z + Vx.Z * 200
            glEnd
            
            glColor3d 0, 0, 1
            glBegin bmLines
            glVertex3f Pt0.X, Pt0.Y, Pt0.Z
            glVertex3f Pt0.X + Vy.X * 200, Pt0.Y + Vy.Y * 200, Pt0.Z + Vy.Z * 200
            glEnd

End If

'Affichage element
For I = 0 To UBound(Machine.Element)
    'Debug.Print Machine.Element(I).Valeur_axe
    Call Affiche_Element(Machine.Element(I), Render_mode)
Next
'gluLookAt Pt0.X * Zoom + 20, Pt0.Y * Zoom, Pt0.Z * Zoom, 0, 0, 0, 0, 0, 1




glPopMatrix



SwapBuffers Pict.hdc
glLoadIdentity
End Sub


Public Sub GETPoint()

Dim MMatrix(1 To 16) As Double
Dim Element As Element3D
On Error Resume Next



For I = 0 To UBound(Machine.Element)
    Element = Machine.Element(I)
    Select Case Element.Type_axe
    
    ' Element Fixe (outil, torche...)
    Case 0
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        
    'Rotation
    Case 1
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glRotatef Element.Vecteur.X * Element.Valeur_axe, 1, 0, 0 ' rotation en x
        glRotatef Element.Vecteur.Y * Element.Valeur_axe, 0, 1, 0  ' rotation en y
        glRotatef Element.Vecteur.Z * Element.Valeur_axe, 0, 0, 1  ' rotation en z
    'Translation
    Case 2
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glTranslatef Element.Vecteur.X * Element.Valeur_axe, Element.Vecteur.Y * Element.Valeur_axe, Element.Vecteur.Z * Element.Valeur_axe

    'pince
    Case 3
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        
    ' Rotule !!!
    Case Else
    
    
    End Select
Next
    
'Recuperation de la matrice OpengL
glGetDoublev glgModelViewMatrix, MMatrix(1)

Debug.Print
Debug.Print "--- T MMatrix-------------------------"
For I = 1 To 16 Step 4
Debug.Print " | " & Format(MMatrix(I), "#,###0.0000") & " | " & Format(MMatrix(I + 1), "#,###0.0000") & " | " & Format(MMatrix(I + 2), "#,###0.0000") & " | " & Format(MMatrix(I + 3), "#,###0.0000") & " | "
Next
Debug.Print "------------------------------------"

Pt0.X = MMatrix(13)
Pt0.Y = MMatrix(14)
Pt0.Z = MMatrix(15)

Vx.X = MMatrix(1)
Vx.Y = MMatrix(2)
Vx.Z = MMatrix(3)

Vy.X = MMatrix(5)
Vy.Y = MMatrix(6)
Vy.Z = MMatrix(7)

Vz.X = MMatrix(9)
Vz.Y = MMatrix(10)
Vz.Z = MMatrix(11)


'P2.X = MMatrix(13) + MMatrix(1) * P3.X + MMatrix(2) * P3.Y + MMatrix(3) * P3.Z
'P2.Y = MMatrix(14) + MMatrix(5) * P3.X + MMatrix(6) * P3.Y + MMatrix(7) * P3.Z
'P2.Z = MMatrix(15) + MMatrix(9) * P3.X + MMatrix(10) * P3.Y + MMatrix(11) * P3.Z

' Reinit de la matrice  pour la suite
glLoadIdentity
End Sub

Sub GETPointPoignet(PositionR As position, Po As Point3, VPX As Point3, VPY As Point3, VPZ As Point3, Optional Niveau As Integer = 5)
Dim MMatrix(1 To 16) As Double
Dim Element As Element3D
On Error Resume Next



For I = 0 To Niveau ' UBound(Machine.Element)
    Element = Machine.Element(I)
    Select Case Element.Type_axe
    
    ' Element Fixe (outil, torche...)
    Case 0
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        
    'Rotation
    Case 1
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glRotatef Element.Vecteur.X * PositionR.Join(I), 1, 0, 0 ' rotation en x
        glRotatef Element.Vecteur.Y * PositionR.Join(I), 0, 1, 0  ' rotation en y
        glRotatef Element.Vecteur.Z * PositionR.Join(I), 0, 0, 1  ' rotation en z
    'Translation
    Case 2
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glTranslatef Element.Vecteur.X * PositionR.Join(I), Element.Vecteur.Y * PositionR.Join(I), Element.Vecteur.Z * PositionR.Join(I)

    'pince
    Case 3
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        
    ' Rotule !!!
    Case Else
    
    
    End Select
Next
    
'Recuperation de la matrice OpengL
glGetDoublev glgModelViewMatrix, MMatrix(1)

Debug.Print
Debug.Print "--- T MMatrix-------------------------"
For I = 1 To 16 Step 4
Debug.Print " | " & Format(MMatrix(I), "#,###0.0000") & " | " & Format(MMatrix(I + 1), "#,###0.0000") & " | " & Format(MMatrix(I + 2), "#,###0.0000") & " | " & Format(MMatrix(I + 3), "#,###0.0000") & " | "
Next
Debug.Print "------------------------------------"

Po.X = MMatrix(13)
Po.Y = MMatrix(14)
Po.Z = MMatrix(15)

VPX.X = MMatrix(1)
VPX.Y = MMatrix(2)
VPX.Z = MMatrix(3)

VPY.X = MMatrix(5)
VPY.Y = MMatrix(6)
VPY.Z = MMatrix(7)

VPZ.X = MMatrix(9)
VPZ.Y = MMatrix(10)
VPZ.Z = MMatrix(11)


'P2.X = MMatrix(13) + MMatrix(1) * P3.X + MMatrix(2) * P3.Y + MMatrix(3) * P3.Z
'P2.Y = MMatrix(14) + MMatrix(5) * P3.X + MMatrix(6) * P3.Y + MMatrix(7) * P3.Z
'P2.Z = MMatrix(15) + MMatrix(9) * P3.X + MMatrix(10) * P3.Y + MMatrix(11) * P3.Z

' Reinit de la matrice  pour la suite
glLoadIdentity
End Sub

Sub Affiche_Element(Element As Element3D, Render_mode As Integer)
Dim P1 As Point3
 
    
If Element.STL_def.NmbVertex > 0 Then
    Select Case Element.Type_axe
    
    
    ' Element Fixe (outil, torche...)
    Case 0
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        
    'Rotation
    Case 1
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glRotatef Element.Vecteur.X * Element.Valeur_axe, 1, 0, 0 ' rotation en x
        glRotatef Element.Vecteur.Y * Element.Valeur_axe, 0, 1, 0  ' rotation en y
        glRotatef Element.Vecteur.Z * Element.Valeur_axe, 0, 0, 1  ' rotation en z

    'Translation
    Case 2
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glTranslatef Element.Vecteur.X * Element.Valeur_axe, Element.Vecteur.Y * Element.Valeur_axe, Element.Vecteur.Z * Element.Valeur_axe
        
    'pliers
    Case 3
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        'P1= point de deplacement de la pince
        P1.X = Element.Vecteur.X * Element.Valeur_axe
        P1.Y = Element.Vecteur.Y * Element.Valeur_axe
        P1.Z = Element.Vecteur.Z * Element.Valeur_axe
    
    ' Rotule !!!
    Case Else
    
    
    End Select

    
    
mode = bmTriangles
If Render_mode = 2 Then
     mode = bmLineLoop
End If

glEnable glcNormalize


    
With Element.STL_def
    For J = 0 To .NmbVertex - 3 Step 3
        glBegin mode
        
        'glColor4f ObjDiffuse(0), ObjDiffuse(1), ObjDiffuse(2), ObjDiffuse(3) 'set the object's diffuse color
        'glMaterialfv faceFront, mprAmbient, ObjAmbient(0) 'set the object's ambient color
        'glMaterialfv faceFront, mprSpecular, ObjSpecular(0) 'set the object's specular color
    
        glColor3d Element.Color.Rouge / 255, Element.Color.Vert / 255, Element.Color.Bleu / 255
        glNormal3f .Normal(J / 3).X, .Normal(J / 3).Y, .Normal(J / 3).Z
        glVertex3f .Vertex(J).X + P1.X, .Vertex(J).Y + P1.Y, .Vertex(J).Z + P1.Z
        glVertex3f .Vertex(J + 1).X + P1.X, .Vertex(J + 1).Y + P1.Y, .Vertex(J + 1).Z + P1.Z
        glVertex3f .Vertex(J + 2).X + P1.X, .Vertex(J + 2).Y + P1.Y, .Vertex(J + 2).Z + P1.Z
        glEnd
    Next
End With

    

'Depth whire
If Render_mode = 1 Then
    With Element.STL_def
        For J = 0 To .NmbVertex - 3 Step 3
    
            glColor3f 0, 0, 0
            glBegin bmLineLoop
            glVertex3f .Vertex(J).X + P1.X, .Vertex(J).Y + P1.Y, .Vertex(J).Z + P1.Z
            glVertex3f .Vertex(J + 1).X + P1.X, .Vertex(J + 1).Y + P1.Y, .Vertex(J + 1).Z + P1.Z
            glVertex3f .Vertex(J + 2).X + P1.X, .Vertex(J + 2).Y + P1.Y, .Vertex(J + 2).Z + P1.Z
            glEnd
        Next
    End With
End If

End If
End Sub


'Reinit objet
Sub Reinit_Element(Elem As Element3D)

    Elem.STL_def.NmbNormal = 0
    Elem.STL_def.NmbVertex = 0
    ReDim Elem.STL_def.Vertex(0)
    ReDim Elem.STL_def.Normal(0)

End Sub
