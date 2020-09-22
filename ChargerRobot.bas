Attribute VB_Name = "ChargerRobot"
Option Explicit

Public File_robot As String

Public Sub Charger_robot(File_Def As String, Mach As Machine3D)
Dim NBElement As Integer
Dim I As Integer
Dim SectionElement As String
' VB QbColor
' ==============
'0   Noir
'1   Bleu
'2   Vert
'3   bleu
'4   Rose
'5   Magenta
'6   jaune
'7   Blanc sale
'8   Gris
'9   Bleu Clair
'10  Vert Clair
'11  bleu Clair
'12  Rose Clair
'13  Magenta Clair
'14  Jaune Clair
'15  Blanc brillant

'
NBElement = Val(mfncGetFromIni("Robot", "Element", File_Def))
ReDim Mach.Element(NBElement)

Mach.Name = mfncGetFromIni("Robot", "Name", File_Def)
Mach.Type = Val(mfncGetFromIni("Robot", "Type", File_Def))
Mach.NB_axe = Val(mfncGetFromIni("Robot", "NB_axe", File_Def))
Mach.Accessoire = Val(mfncGetFromIni("Robot", "Accessoire", File_Def))
    
' Reinit objet
'Type_axe = 0 => Translation
'Type_axe = 1 => rotation
For I = 0 To NBElement
    'Name de la section
    SectionElement = "Element" & I
    
    Mach.Element(I).Name = mfncGetFromIni(SectionElement, "Name", File_Def)
    Mach.Element(I).File = mfncGetFromIni(SectionElement, "File", File_Def)

    Mach.Element(I).Color = RVB(QBColor(Val(mfncGetFromIni(SectionElement, "Couleur", File_Def))))

    Mach.Element(I).MiniAxe = Val(mfncGetFromIni(SectionElement, "Mini_axe", File_Def))
    Mach.Element(I).MaxiAxe = Val(mfncGetFromIni(SectionElement, "Maxi_axe", File_Def))
    
    Mach.Element(I).Type_axe = Val(mfncGetFromIni(SectionElement, "Type_axe", File_Def))
    Mach.Element(I).Origine.X = Val(mfncGetFromIni(SectionElement, "Origine_X", File_Def))
    Mach.Element(I).Origine.Y = Val(mfncGetFromIni(SectionElement, "Origine_Y", File_Def))
    Mach.Element(I).Origine.Z = Val(mfncGetFromIni(SectionElement, "Origine_Z", File_Def))
    Mach.Element(I).Vecteur.X = Val(mfncGetFromIni(SectionElement, "Vecteur_X", File_Def))
    Mach.Element(I).Vecteur.Y = Val(mfncGetFromIni(SectionElement, "Vecteur_Y", File_Def))
    Mach.Element(I).Vecteur.Z = Val(mfncGetFromIni(SectionElement, "Vecteur_Z", File_Def))
Next I


End Sub


'****************************************************************
' Name: .INI read/write routines
' Description:.INI read/write routines
'mfncGetFromIni-- Reads from an *.INI file strFileName(full path & file name)
'mfncWriteIni--Writes to an *.INI file called strFileName (full path & file name)
'****************************************************************
Function mfncGetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
       '*** DESCRIPTION:Reads from an *.INI file strFileName (fullpath & file name)
       '     '*** RETURNS:The string stored in [strSectionHeader], line  beginning strVariableName=
       '     '*** NOTE: Requires declaration of API call GetPrivateProfileString
       '     'Initialise variable
       Dim strReturn As String
       '     'Blank the return string
       strReturn = String(255, Chr(0))
       '     'Get requested information, trimming the returned string
       mfncGetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function
Function mfncWriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
       '*** DESCRIPTION:Writes to an *.INI file called strFileName
       '     (full       path & file name)
       '*** RETURNS:Integer indicating failure (0) or success (other)       to write
       '     '*** NOTE: Requires declaration of API call     WritePrivateProfileString
       '     'Call the API
       mfncWriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function
Function mfncDeleteIniKey(strSectionHeader As String, strVariableName As String, strFileName As String) As Integer
       '*** DESCRIPTION:Writes to an *.INI file called strFileName
       '     (full       path & file name)
       '*** RETURNS:Integer indicating failure (0) or success (other)       to write
       '     '*** NOTE: Requires declaration of API call     WritePrivateProfileString
       '     'Call the API
       mfncDeleteIniKey = WritePrivateProfileString(strSectionHeader, strVariableName, 0&, strFileName)
End Function
Function mfncDeleteIniSection(strSectionHeader As String, strFileName As String) As Integer
       '*** DESCRIPTION:Writes to an *.INI file called strFileName
       '     (full       path & file name)
       '*** RETURNS:Integer indicating failure (0) or success (other)       to write
       '     '*** NOTE: Requires declaration of API call     WritePrivateProfileString
       '     'Call the API
       mfncDeleteIniSection = WritePrivateProfileString(strSectionHeader, 0&, 0&, strFileName)
End Function



