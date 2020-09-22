VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ROBOT_SIMUL_FRM 
   Caption         =   "Robot_simul"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   14475
   FillColor       =   &H00FF0000&
   Icon            =   "Robot_simul.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleMode       =   0  'User
   ScaleWidth      =   965
   Begin VB.PictureBox Support_BTN 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   961
      TabIndex        =   0
      Top             =   3720
      Width           =   14415
      Begin VB.Frame FrameVue 
         Caption         =   "Vue"
         Height          =   1575
         Left            =   120
         TabIndex        =   81
         Top             =   3120
         Width           =   1695
         Begin VB.OptionButton OptionVue 
            Caption         =   "Bottom View"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   85
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton OptionVue 
            Caption         =   "Top View"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   84
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton OptionVue 
            Caption         =   "Pliers View"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   83
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton OptionVue 
            Caption         =   "Standard View"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame FrameAffichage 
         Caption         =   "Display"
         Height          =   2055
         Left            =   120
         TabIndex        =   72
         Top             =   960
         Width           =   1695
         Begin VB.OptionButton OPT 
            Caption         =   "Depth Wire"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OPT 
            Caption         =   "Wire"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton OPT 
            Caption         =   "Shade"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   75
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox CHK 
            Caption         =   "Raster"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   1320
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox Optiontracer 
            Caption         =   "Vector"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1560
            Width           =   1215
         End
      End
      Begin VB.Frame FrameDestination 
         Caption         =   "Goto Point"
         ForeColor       =   &H00000000&
         Height          =   4455
         Left            =   6000
         TabIndex        =   54
         Top             =   120
         Width           =   2415
         Begin VB.PictureBox PictureVoyant 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   470
            Index           =   2
            Left            =   720
            Picture         =   "Robot_simul.frx":030A
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   80
            Top             =   3840
            Width           =   470
         End
         Begin VB.PictureBox PictureVoyant 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   470
            Index           =   1
            Left            =   120
            Picture         =   "Robot_simul.frx":0EEC
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   79
            Top             =   3840
            Width           =   470
         End
         Begin VB.PictureBox PictureVoyant 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   470
            Index           =   0
            Left            =   1800
            Picture         =   "Robot_simul.frx":1ACE
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   78
            Top             =   2700
            Width           =   470
         End
         Begin VB.CommandButton CommandUndo 
            Height          =   375
            Left            =   120
            Picture         =   "Robot_simul.frx":26B0
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   3360
            Width           =   615
         End
         Begin VB.CommandButton CommandRedo 
            Height          =   375
            Left            =   1680
            Picture         =   "Robot_simul.frx":3232
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   3360
            Width           =   615
         End
         Begin VB.CheckBox PositionGauche 
            Caption         =   "Left Position"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   3000
            Width           =   1695
         End
         Begin VB.CheckBox PositionHaute 
            Caption         =   "Top Position"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   2640
            Width           =   1695
         End
         Begin VB.CommandButton CommandGoto 
            Height          =   375
            Left            =   800
            Picture         =   "Robot_simul.frx":3E2C
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   3360
            Width           =   855
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   5
            Left            =   720
            TabIndex        =   65
            Text            =   "0"
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   4
            Left            =   720
            TabIndex        =   63
            Text            =   "0"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   61
            Text            =   "0"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   57
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   56
            Text            =   "0"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   55
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   66
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   64
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   62
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   60
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   59
            Top             =   600
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   58
            Top             =   960
            Width           =   255
         End
      End
      Begin VB.CommandButton Bouton2 
         Caption         =   "Execute mvt"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   1455
      End
      Begin VB.PictureBox MessageLOG 
         AutoRedraw      =   -1  'True
         Height          =   4575
         Left            =   8640
         ScaleHeight     =   4515
         ScaleWidth      =   5475
         TabIndex        =   42
         Top             =   120
         Width           =   5535
      End
      Begin VB.Frame FramePince 
         Caption         =   "Pliers Control"
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   3840
         TabIndex        =   9
         Top             =   3120
         Width           =   2175
         Begin VB.TextBox Pince 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Text            =   "0"
            Top             =   600
            Width           =   1935
         End
         Begin MSComctlLib.Slider SliderPince 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            Max             =   15
         End
      End
      Begin VB.Frame FrameXYZABC 
         Caption         =   "Absolute Position"
         ForeColor       =   &H00000000&
         Height          =   2895
         Left            =   3840
         TabIndex        =   7
         Top             =   120
         Width           =   2175
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   8
            Left            =   960
            TabIndex        =   53
            Text            =   "0"
            Top             =   2400
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   52
            Text            =   "0"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   11
            Left            =   1560
            TabIndex        =   51
            Text            =   "0"
            Top             =   2400
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   10
            Left            =   1560
            TabIndex        =   50
            Text            =   "0"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   9
            Left            =   1560
            TabIndex        =   49
            Text            =   "0"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   7
            Left            =   960
            TabIndex        =   48
            Text            =   "0"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   6
            Left            =   960
            TabIndex        =   47
            Text            =   "0"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   39
            Text            =   "0"
            Top             =   2400
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   37
            Text            =   "0"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   34
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   32
            Text            =   "0"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   30
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Vz"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1680
            TabIndex        =   46
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Vy"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1080
            TabIndex        =   45
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Vx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   44
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "K"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "J"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   33
            Top             =   960
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   31
            Top             =   600
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame FrameControl 
         Caption         =   "Manual Control"
         ForeColor       =   &H00000000&
         Height          =   4575
         Left            =   1920
         TabIndex        =   2
         Top             =   120
         Width           =   1935
         Begin VB.TextBox AxeRobot 
            Height          =   285
            Index           =   6
            Left            =   840
            TabIndex        =   27
            Text            =   "0"
            Top             =   3840
            Width           =   855
         End
         Begin VB.PictureBox PictureAxeRobot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   6
            Left            =   120
            Picture         =   "Robot_simul.frx":4C5A
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   26
            Top             =   4200
            Width           =   330
         End
         Begin VB.TextBox AxeRobot 
            Height          =   285
            Index           =   5
            Left            =   840
            TabIndex        =   23
            Text            =   "0"
            Top             =   3120
            Width           =   855
         End
         Begin VB.PictureBox PictureAxeRobot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   5
            Left            =   120
            Picture         =   "Robot_simul.frx":514C
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   22
            Top             =   3480
            Width           =   330
         End
         Begin VB.TextBox AxeRobot 
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   19
            Text            =   "0"
            Top             =   2400
            Width           =   855
         End
         Begin VB.PictureBox PictureAxeRobot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   4
            Left            =   120
            Picture         =   "Robot_simul.frx":563E
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   18
            Top             =   2760
            Width           =   330
         End
         Begin VB.TextBox AxeRobot 
            Height          =   285
            Index           =   3
            Left            =   840
            TabIndex        =   15
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.PictureBox PictureAxeRobot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   120
            Picture         =   "Robot_simul.frx":5B30
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   14
            Top             =   2040
            Width           =   330
         End
         Begin VB.PictureBox PictureAxeRobot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   120
            Picture         =   "Robot_simul.frx":6022
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   11
            Top             =   1320
            Width           =   330
         End
         Begin VB.TextBox AxeRobot 
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   10
            Text            =   "0"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox AxeRobot 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   6
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.PictureBox PictureAxeRobot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   120
            Picture         =   "Robot_simul.frx":6514
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   5
            Top             =   600
            Width           =   330
         End
         Begin MSComctlLib.Slider SliderAxeRobot 
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   3
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeRobot 
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   12
            Top             =   1320
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeRobot 
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   16
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeRobot 
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   20
            Top             =   2760
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeRobot 
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   24
            Top             =   3480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeRobot 
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   28
            Top             =   4200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin VB.Label LabelAxeRobot 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   29
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label LabelAxeRobot 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   25
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label LabelAxeRobot 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   21
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label LabelAxeRobot 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   17
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label LabelAxeRobot 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   255
         End
         Begin VB.Label LabelAxeRobot 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   120
      ScaleHeight     =   235
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   364
      TabIndex        =   1
      Top             =   120
      Width           =   5490
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&File"
      Begin VB.Menu mnuCharger 
         Caption         =   "&Load Robot"
      End
      Begin VB.Menu mnuExecute 
         Caption         =   "&Execute Movement"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "ROBOT_SIMUL_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Chargement As Boolean


Private Sub AxeRobot_Change(Index As Integer)
    If Val(AxeRobot(Index)) <> 0 Then
        AxeRobot(Index) = Round(Val(AxeRobot(Index)), 4)
    End If
End Sub


'Initialisation du robot
Public Sub init_robot(File_Def_machine As String)
'Dim File_STL As String
Dim I As Integer

'reinit affichage si plusieurs chargement de robot different
For I = 4 To 6
    SliderAxeRobot(I).Visible = False
    AxeRobot(I).Visible = False
    LabelAxeRobot(I).Visible = False
    PictureAxeRobot(I).Visible = False
Next


'Chargement caractéristique robot
Call Charger_robot(File_Def_machine, Machine)



'Recuperation de la definition geométrique via File STL ascii
For I = 0 To UBound(Machine.Element)
    Call Reinit_Element(Machine.Element(I))
    Call ChargeFichier(App.Path + "\Robot_def\" + Machine.Element(I).File, Machine.Element(I).STL_def)
Next

For I = 1 To Machine.NB_axe
    PictureAxeRobot(I).Picture = LoadResPicture(Machine.Element(I).Type_axe, vbResBitmap)
    SliderAxeRobot(I).Min = Machine.Element(I).MiniAxe
    SliderAxeRobot(I).Max = Machine.Element(I).MaxiAxe
    SliderAxeRobot(I).Visible = True
    AxeRobot(I).Visible = True
    LabelAxeRobot(I).Visible = True
    PictureAxeRobot(I).Visible = True
Next

'modif axe modifiable suivant type robot
Select Case Machine.Type
Case 1 'scara
    AxeDestination(3).Visible = False
    AxeDestination(4).Visible = False
    LabelAxeDestination(3).Visible = False
    LabelAxeDestination(4).Visible = False
    PositionHaute.Visible = False
    
Case 2 'Polymorph
    AxeDestination(3).Visible = True
    AxeDestination(4).Visible = True
    LabelAxeDestination(3).Visible = True
    LabelAxeDestination(4).Visible = True
    PositionHaute.Visible = True
End Select

' presence Accessoire
If Machine.Accessoire Then
    SliderPince.Min = Machine.Element(8).MiniAxe
    SliderPince.Max = Machine.Element(8).MaxiAxe
Else
    FramePince.Visible = False
End If


Chargement = True
End Sub

Private Sub CHK_Click()
    Call Pic_Paint
End Sub

Private Sub Bouton2_Click()
    Call mnuExecute_Click
End Sub

'Execute un mouvement entre deux positions du robot
Sub Execute_mouvement()
Dim Increment_max As Integer
Dim Increment_controle As Integer
Dim J As Integer
Dim I As Integer

For J = 1 To 2
    ROBOT_SIMUL_FRM.PictureVoyant(J).Visible = False
    ROBOT_SIMUL_FRM.Refresh
Next
For J = 1 To Machine.NB_axe
    If Position_suivante.Join(J) > Machine.Element(J).MaxiAxe Then
        MessageLOG.Print "Axis " & J & " Out of limits  !!!!!"
        ROBOT_SIMUL_FRM.PictureVoyant(2).Visible = True
        Exit Sub
    End If
    
    If Position_suivante.Join(J) < Machine.Element(J).MiniAxe Then
        MessageLOG.Print "Axis " & J & " Out of limits  !!!!!"
        ROBOT_SIMUL_FRM.PictureVoyant(2).Visible = True
        Exit Sub
    End If
    
Next

Increment_max = 10
'Cherche l'incrément suer les axes "majeurs" = ceux qui provoquent les plus grand déplacement
For J = 1 To 3
     Increment_controle = Abs(Position_suivante.Join(J) - Position_precedente.Join(J))
     If Increment_controle > Increment_max Then
     Increment_max = Increment_controle
     End If
Next

For I = 0 To Increment_max - 1
    For J = 1 To Machine.NB_axe
        Machine.Element(J).Valeur_axe = Machine.Element(J).Valeur_axe + ((Position_suivante.Join(J) - Position_precedente.Join(J))) / Increment_max
    Next
    ' cherche la position du point de controle
    Call GETPoint
    Call Pic_Paint
Next

' Position exacte demandée
For J = 1 To Machine.NB_axe
    Machine.Element(J).Valeur_axe = Position_suivante.Join(J)
Next
Call GETPoint
Call Pic_Paint

' Affiche la coordonnées obtenues
Call Affiche_coord
' Position exacte affichée
For J = 1 To Machine.NB_axe
  AxeRobot(J) = Machine.Element(J).Valeur_axe
Next

ROBOT_SIMUL_FRM.PictureVoyant(1).Visible = True

End Sub

Private Sub CommandGoto_Click()
Dim pos As position
Dim retour As Boolean
Dim Pt As Point3
Dim J As Integer

For J = 1 To 2
ROBOT_SIMUL_FRM.PictureVoyant(J).Visible = False
Next

If Not Chargement Then
MessageLOG.Print "First you must load a Robot !!!!!"
ROBOT_SIMUL_FRM.PictureVoyant(2).Visible = True
Exit Sub
End If




Pt.X = Val(AxeDestination(0))
Pt.Y = Val(AxeDestination(1))
Pt.Z = Val(AxeDestination(2))
Select Case Machine.Type
Case 1 'scara
     retour = calcul_position_scara(Pt, Val(AxeDestination(5)), pos)
Case 2 ' polymorh
    retour = calcul_position_polymorh(Pt, Val(AxeDestination(3)), Val(AxeDestination(4)), Val(AxeDestination(5)), pos)
End Select
' Position trouvée
If retour Then

For J = 1 To Machine.NB_axe
Position_precedente.Join(J) = Machine.Element(J).Valeur_axe
Next
Position_suivante = pos
Call Execute_mouvement
End If

 

End Sub

Private Sub CommandRedo_Click()
Dim J As Integer

If Not Chargement Then
MessageLOG.Print "First you must load a robot !!!!!"
Exit Sub
End If

Position_suivante = Position_precedente
For J = 1 To 6
Position_precedente.Join(J) = Machine.Element(J).Valeur_axe
Next
Call Execute_mouvement
End Sub

Private Sub CommandUndo_Click()
Dim J As Integer

If Not Chargement Then
MessageLOG.Print "First you must load a robot !!!!!"
Exit Sub
End If

Position_suivante = Position_precedente
For J = 1 To 6
Position_precedente.Join(J) = Machine.Element(J).Valeur_axe
Next
Call Execute_mouvement

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then PosX = PosX + 1
If KeyCode = vbKeyUp Then PosX = PosX - 1
If KeyCode = vbKeyRight Then PosY = PosY + 1
If KeyCode = vbKeyLeft Then PosY = PosY - 1


Call Pic_Paint

End Sub

Private Sub Form_Load()
Dim I As Integer

' Init pour Calcul de convertion degree radian
    RADTODEG = 180 / (4 * Atn(1))
    DEGTORAD = (4 * Atn(1)) / 180
    PI = (4 * Atn(1))
    
    ' init des voyants
    For I = 1 To 2
        PictureVoyant(I).Left = PictureVoyant(0).Left
        PictureVoyant(I).Top = PictureVoyant(0).Top
        PictureVoyant(I).Visible = False
    Next
    
'Init des types d'axes
For I = 1 To 6
PictureAxeRobot(I).Picture = LoadResPicture(3, vbResBitmap)
Next


For I = 4 To 6
SliderAxeRobot(I).Visible = False
AxeRobot(I).Visible = False
LabelAxeRobot(I).Visible = False
PictureAxeRobot(I).Visible = False
Next

'Previsu et chargement Robot
PreVisu.Show

' Initialisation du controle picturebox en opengl
'LoadGL Pic


'Call Pic_DblClick
'Call Pic_Paint



End Sub


Private Sub Form_Resize()
Dim W, H As Integer

W = Me.ScaleWidth
H = Me.ScaleHeight
Pic.Width = W - 15
With Support_BTN
    .Width = W - 15
    .Top = H - .ScaleHeight - 5
    Pic.Height = H - .ScaleHeight - 20
End With


LoadGL Pic
Call Pic_Paint
End Sub






Private Sub mnuCharger_Click()

 PreVisu.Show
End Sub

Private Sub mnuExecute_Click()
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim NB_mvt As Integer
Dim NB_axes As Integer
Dim PosX As String
Dim Tab_split

On Error GoTo fin

If Not Chargement Then
MessageLOG.Print "First you must load a robot !!!!!"
Exit Sub
End If

NB_axes = Val(mfncGetFromIni("Robot", "NB_axe", File_robot))
For J = 1 To NB_axes
Position_precedente.Join(J) = Machine.Element(J).Valeur_axe
Next
For J = 1 To NB_axes
Position_suivante.Join(J) = 0
Next

'retour a zero
MessageLOG.CurrentY = 0
MessageLOG.Cls


'MessageLOG.Print "Goto Position 0 ..."
'Call Execute_mouvement
'Position_precedente = Position_suivante

NB_mvt = Val(mfncGetFromIni("MvtDemo", "Nb_point", File_robot))

For I = 1 To NB_mvt
   PosX = "Position" & I
   Tab_split = Split(mfncGetFromIni("MvtDemo", PosX, File_robot), ",")

    For K = 1 To NB_axes
        Position_suivante.Join(K) = Tab_split(K - 1)
    Next K

    
    MessageLOG.Print "Goto Position " & I & " ..."
    Call Execute_mouvement
    
    Position_precedente = Position_suivante
Next I

fin:

End Sub

Private Sub mnuQuitter_Click()
Unload Me
End Sub

Private Sub OPT_Click(Index As Integer)
Call Pic_Paint
End Sub

Private Sub Optiontracer_Click()
Call Pic_Paint
End Sub

Private Sub OptionVue_Click(Index As Integer)
Call Pic_Paint
End Sub

Private Sub Pic_DblClick()
xm = 45
ym = 0
zm = 45

Zoom = 0.0018

PosX = 0
PosY = 0

End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim diffx As Single
Dim diffy As Single
Dim facteur_diff
Dim facteur_zoom
facteur_diff = 20
facteur_zoom = 0.0001

'Translation
If Button = 1 And Shift = 0 Then
' Recherche la direction principale
  diffx = Abs(SVGposX - X)
  diffy = Abs(SVGposY - Y)
  
  If diffx > diffy Then
       If X > SVGposX Then
       PosY = PosY + facteur_diff
       Else
       PosY = PosY - facteur_diff
      End If
  Else
       If Y > SVGposY Then
       PosX = PosX - facteur_diff
      Else
       PosX = PosX + facteur_diff
      End If
  End If
End If

'  Zoom global
If Button = 2 And Shift = 0 Then
 If Y < SVGposY Then
 Zoom = Zoom - facteur_zoom
 Else
 Zoom = Zoom + facteur_zoom
 End If
 
 If Zoom < 0 Then Zoom = 0.0001
 'Zoom = Y / 10000
End If

'  Zoom fenetre
If Button = 3 And Shift = 0 Then

End If

'rotation
If Button = 1 And Shift = 1 Then
' Recherche la direction principale
  diffx = Abs(SVGposX - X)
  diffy = Abs(SVGposY - Y)
  
  If diffx > diffy Then
       If X > SVGposX Then
       ym = ym - X * 0.005
       Else
       ym = ym + X * 0.005
       End If
  Else
       If Y > SVGposY Then
       xm = xm - Y * 0.005
       Else
       xm = xm + Y * 0.005
       End If
  End If
End If

If Button = 2 And Shift = 1 Then
       If X > SVGposX Then
       zm = zm - X * 0.005
       Else
       zm = zm + X * 0.005
       End If
End If

If Button > 0 Then
SVGposX = X
SVGposY = Y
End If


Call Pic_Paint

End Sub


Private Sub Pic_Paint()
    Dim render As Integer
    'trait caché
    If Abs(CInt(OPT(0).Value)) = 1 Then
    render = 1
    End If
    'Fil de fer
    If Abs(CInt(OPT(1).Value)) = 1 Then
    render = 2
    End If
    'ombrée
    If Abs(CInt(OPT(2).Value)) = 1 Then
    render = 3
    End If
    
    Call DessineRobot(Pic, CBool(CHK.Value), render, Optiontracer)
End Sub

Private Sub PictureAxeRobot_dblClick(Index As Integer)
Dim Cancel As Boolean
    AxeRobot(Index) = 0
    Call AxeRobot_Validate(Index, Cancel)
End Sub

Private Sub Pince_Change()

If Chargement Then
    Machine.Element(7).Valeur_axe = Pince
    Machine.Element(8).Valeur_axe = Pince
    Call Pic_Paint
End If

End Sub



Private Sub SliderAxeRobot_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cancel As Boolean
AxeRobot(Index) = SliderAxeRobot(Index).Value
Call AxeRobot_Validate(Index, Cancel)
End Sub

Private Sub AxeRobot_Validate(Index As Integer, Cancel As Boolean)
Dim I As Integer

   MessageLOG.Print "Validate " & Index
    ' force un nombre
   AxeRobot(Index) = Val(AxeRobot(Index))

If Chargement Then

For I = 1 To Machine.NB_axe
   Position_precedente.Join(I) = Machine.Element(I).Valeur_axe
Next


If Val(AxeRobot(Index)) < SliderAxeRobot(Index).Min Then
    MessageLOG.CurrentY = 0
    MessageLOG.Cls
    AxeRobot(Index) = SliderAxeRobot(Index).Min
    MessageLOG.Print "Axis " & Index & " < limit " & SliderAxeRobot(Index).Min
End If

If Val(AxeRobot(Index)) > SliderAxeRobot(Index).Max Then
    MessageLOG.CurrentY = 0
    MessageLOG.Cls
    AxeRobot(Index) = SliderAxeRobot(Index).Max
    MessageLOG.Print "Axis " & Index & " > limit " & SliderAxeRobot(Index).Max
End If



For I = 1 To 6
    Position_suivante.Join(I) = AxeRobot(I)
Next



Call Execute_mouvement


    'Machine.Element(Index).Valeur_axe = Val(AxeRobot(Index))
    'Recuperation de position via OpenGl fonction
    'Call GETPoint

    
End If

End Sub



Private Sub SliderPince_Click()
Pince = SliderPince.Value
End Sub

