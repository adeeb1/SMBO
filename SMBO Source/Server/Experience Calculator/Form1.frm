VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Experience Calculator"
   ClientHeight    =   2580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtCurrentExp 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtCurrentLvl 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Submit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblCurrentExp 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Exp Here:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblCurrentLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Level Here:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1530
      Width           =   1575
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Submit_Click()
    Dim CurrentLevel As Double, CurrentExp As Double, NextLevel As Double, Result As Double
    
    CurrentLevel = Val(txtCurrentLvl.Text)
    NextLevel = CurrentLevel + 1
    CurrentExp = Val(txtCurrentExp.Text)
    
    Result = (((3.22 * (CurrentLevel + NextLevel)) + ((CurrentLevel / 2.5) ^ 2)) * (Sqr((CurrentExp * (CurrentLevel + 2))))) / (NextLevel)
    Result = Round(Result, 0)
    
    lblResult.Caption = "You will need " & Result & " exp for level " & NextLevel
    txtCurrentLvl.Text = NextLevel
    txtCurrentExp.Text = Result
End Sub
