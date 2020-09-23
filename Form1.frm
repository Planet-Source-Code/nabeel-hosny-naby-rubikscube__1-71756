VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9420
   ClientLeft      =   15
   ClientTop       =   -180
   ClientWidth     =   6975
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   60
      TabIndex        =   21
      Top             =   7800
      Width           =   6840
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   600
         Top             =   120
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Easy "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   0
         Left            =   80
         MouseIcon       =   "Form1.frx":0BC2
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Tag             =   "5"
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Midium "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   1
         Left            =   80
         MouseIcon       =   "Form1.frx":0D14
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Tag             =   "10"
         Top             =   625
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hard "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   2
         Left            =   80
         MouseIcon       =   "Form1.frx":0E66
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Tag             =   "20"
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Scramble 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scramble"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   1750
         MouseIcon       =   "Form1.frx":0FB8
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   200
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moves No :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4815
         TabIndex        =   29
         Top             =   945
         Width           =   1080
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1100
         Left            =   3800
         MouseIcon       =   "Form1.frx":110A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":125C
         Stretch         =   -1  'True
         ToolTipText     =   " Nabeel Hosny Cairo / 2004 Click to Exit"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6060
         TabIndex        =   28
         Top             =   945
         Width           =   480
      End
      Begin VB.Label RestMe 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reset Me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   435
         Left            =   1750
         MouseIcon       =   "Form1.frx":25C8E
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   560
         Width           =   1850
      End
      Begin VB.Label SolveMe 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solve Me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   1750
         MouseIcon       =   "Form1.frx":25DE0
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   900
         Width           =   1850
      End
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   4800
         TabIndex        =   25
         Top             =   370
         Width           =   1900
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1100
         Left            =   1750
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   450
         Left            =   4740
         Shape           =   4  'Rounded Rectangle
         Top             =   280
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   450
         Left            =   4740
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdTopCCWise 
      Caption         =   ">>TCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdTopCWise 
      Caption         =   "<<TC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "t"
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFMBCCWise 
      Caption         =   ">>FMBCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      TabIndex        =   20
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   990
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdFMBCWise 
      Caption         =   "<<FMBC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      TabIndex        =   19
      Tag             =   "f"
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   660
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdTMBCCWise 
      Caption         =   ">>TMBCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   990
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdTMBCWise 
      Caption         =   "<<TMBC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "t"
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   660
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdLMRCCWise 
      Caption         =   ">>LMRCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5280
      TabIndex        =   16
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   990
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdLMRCWise 
      Caption         =   "<<LMRCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5280
      TabIndex        =   15
      Tag             =   "i"
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   660
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdBCCWise 
      Caption         =   ">>BCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      TabIndex        =   10
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   1650
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdBCWise 
      Caption         =   "<<BC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      TabIndex        =   9
      Tag             =   "b"
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFCWise 
      Caption         =   "<<FC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      TabIndex        =   12
      Tag             =   "f"
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFCCWise 
      Caption         =   ">>FCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      TabIndex        =   11
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdBottomCCWise 
      Caption         =   ">>DCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   1650
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdBottomCWise 
      Caption         =   "<<DC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "d"
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdRCWise 
      Caption         =   "<<RC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5640
      TabIndex        =   4
      Tag             =   "r"
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdRCCWise 
      Caption         =   ">>RCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5640
      TabIndex        =   3
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdLCWise 
      Caption         =   "<<LC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   2
      Tag             =   "i"
      ToolTipText     =   "Clock-Wise Rotation"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdLCCWise 
      Caption         =   ">>LCC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5880
      TabIndex        =   1
      ToolTipText     =   "Counter Clock-Wise Rotation"
      Top             =   1650
      Visible         =   0   'False
      Width           =   1095
   End
   Begin RubiksCube.RubikCube RubikCube1 
      Height          =   6300
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11113
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   7775
      Left            =   120
      Picture         =   "Form1.frx":25F32
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Dim ElapsedSeconds  As Long     'time elapsed
Dim level As Integer
Dim kk As Integer
Private Sub cmdBCCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "BA", "CCWise"
If kk = 0 Then List1.AddItem 16: gome
End Sub

Private Sub cmdBCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "BA", "CWise"
If kk = 0 Then List1.AddItem 7: gome
End Sub

Private Sub cmdBottomCCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "BO", "CCWise"
If kk = 0 Then List1.AddItem 18: gome
End Sub

Private Sub cmdBottomCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "BO", "CWise"
If kk = 0 Then List1.AddItem 11: gome
End Sub

Private Sub cmdFCCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "F", "CCWise"
If kk = 0 Then List1.AddItem 13: gome
End Sub

Private Sub cmdFCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "F", "CWise"
If kk = 0 Then List1.AddItem 1: gome
End Sub

Private Sub cmdFMBCCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "FMB", "CCWise"
If kk = 0 Then List1.AddItem 21: gome
End Sub

Private Sub cmdFMBCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "FMB", "CWise"
If kk = 0 Then List1.AddItem 22: gome
End Sub

Private Sub cmdLCCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "L", "CCWise"
If kk = 0 Then List1.AddItem 15: gome
End Sub

Private Sub cmdLCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "L", "CWise"
If kk = 0 Then List1.AddItem 5: gome
End Sub

Private Sub cmdLMRCCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "LMR", "CCWise"
If kk = 0 Then List1.AddItem 23: gome
End Sub

Private Sub cmdLMRCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "LMR", "CWise"
If kk = 0 Then List1.AddItem 24: gome
End Sub

Private Sub cmdRCCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "R", "CCWise"
If kk = 0 Then List1.AddItem 14: gome
End Sub

Private Sub cmdRCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "R", "CWise"
If kk = 0 Then List1.AddItem 3: gome
End Sub

Private Sub cmdTMBCCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "TMB", "CCWise"
If kk = 0 Then List1.AddItem 20: gome
End Sub

Private Sub cmdTMBCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "TMB", "CWise"
If kk = 0 Then List1.AddItem 19: gome
End Sub

Private Sub cmdTopCCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "T", "CCWise"
If kk = 0 Then List1.AddItem 17: gome
End Sub

Private Sub cmdTopCWise_Click()
Dim Z As String

X = RubikCube1.GetXCoor
Y = RubikCube1.GetYCoor
Z = X & "/" & Y
RubikCube1.GetRotation Z, "T", "CWise"
If kk = 0 Then List1.AddItem 9: gome
End Sub

Private Sub Command1_Click()
Label3.Caption = "1"
List1.Clear
Label2.Caption = "0000"
'if Time Elapsed is enabled..
If Timer2.Enabled = True Then
    
    'display elapsed time..
    MsgBox lblTimer.Caption, , Me.Caption

    'ask if user want to scramble cube again..
    ans = MsgBox("Are you sure?", _
        vbQuestion + vbYesNo, "Scramble Cube")

    'if user answered "NO", exit to this event..
    If ans = vbNo Then Exit Sub

End If
 RubikCube1.ResetCube
'scramble cube..
RubikCube1.ScrambleCube

'display that timer has been started..
MsgBox "Timer starts now!!", vbInformation, Me.Caption

lblTimer.Caption = IIf(level = 0, "Easy :", IIf(level = 1, "Midium :", "Hard :")) & " - " & "00:00:00"

'enable Timer2 control..
Timer2.Enabled = True
For i = 0 To List1.ListCount - 1
Select Case List1.List(i)
Case 1, 3, 5, 7, 9, 11
Label2.Caption = Format$(Val(Label2.Caption) + 1, "0000")
Case 2, 4, 6, 8, 10, 12
Label2.Caption = Format$(Val(Label2.Caption) + 3, "0000")
End Select
Next

Option1(0).Enabled = False: Option1(1).Enabled = False: Option1(2).Enabled = False
End Sub

Private Sub Command2_Click()
Label3.Caption = "0"
'ask if Time Elapsed is enabled..
If Timer2.Enabled = True Then
    
    'ask if user wants to reset cube..
    ans = MsgBox("Are you sure?", _
        vbQuestion + vbYesNo, "Reset Cube")

    'if user answered "NO", exit to this event..
    If ans = vbNo Then Exit Sub
    
    'display elpased time..
    MsgBox lblTimer.Caption, , Me.Caption
    
    'disable Timer2..
    Timer2.Enabled = False

End If

'reset cube..
RubikCube1.ResetCube

'reset lblTimer to default caption..
lblTimer.Caption = IIf(level = 0, "Easy :", IIf(level = 1, "Midium :", "Hard :")) & " - " & "00:00:00"

Label2.Caption = "0000"
'display message..
MsgBox "Rubik's cube has been reset.", vbInformation, Me.Caption

Option1(0).Enabled = True: Option1(1).Enabled = True: Option1(2).Enabled = True
End Sub

Private Sub Command3_Click()

Label3.Caption = "0"
kk = 1
For i = List1.ListCount - 1 To 0 Step -1
Select Case List1.List(i)
        
        Case 1
         cmdFCCWise_Click
         dome
         
        Case 2
         For ii = 1 To 3
          cmdFCCWise_Click
          dome
         Next

        Case 3
         cmdRCCWise_Click
         dome
         
        Case 4
          For ii = 1 To 3
           cmdRCCWise_Click
           dome
          Next
          
        Case 5
         cmdLCCWise_Click
         dome
         
        Case 6
          For ii = 1 To 3
           cmdLCCWise_Click
           dome
          Next
         
        Case 7
          cmdBCCWise_Click
          dome
        
        Case 8
          For ii = 1 To 3
           cmdBCCWise_Click
           dome
          Next
          
        Case 9
         cmdTopCCWise_Click
          dome
          
        Case 10
          For ii = 1 To 3
           cmdTopCCWise_Click
           dome
          Next

        Case 11
          cmdBottomCCWise_Click
          dome
        Case 12
          For ii = 1 To 3
           cmdBottomCCWise_Click
           dome
          Next
          
        Case 13
          cmdFCWise_Click
          dome
          
        Case 14
          cmdRCWise_Click
          dome
          
         Case 15
          cmdLCWise_Click
          dome
          
        Case 16
         cmdBCWise_Click
          dome
          
        Case 17
          cmdTopCWise_Click
          dome
          
         Case 18
          cmdBottomCWise_Click
          dome
          
         Case 19
         cmdTMBCCWise_Click
         dome
         
         Case 20
         cmdTMBCWise_Click
         dome
         
         Case 21
         cmdFMBCWise_Click
         dome
         
         Case 22
         cmdFMBCCWise_Click
         dome
         
         Case 23
         cmdLMRCWise_Click
         dome
         
         Case 24
         cmdLMRCCWise_Click
         dome
         
         
    End Select

Next
List1.Clear
Option1(0).Enabled = True: Option1(1).Enabled = True: Option1(2).Enabled = True
End Sub
Sub dome()
Label2.Caption = Format$(Val(Label2.Caption) - 1, "0000")
End Sub
Sub gome()
Label2.Caption = Format$(Val(Label2.Caption) + 1, "0000")
End Sub


Private Sub Form_Load()
SetWindowRgn Me.hwnd, CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 75, 75), True
'SetWindowRgn Frame1.hwnd, CreateRoundRectRgn(0, 0, Frame1.Width / Screen.TwipsPerPixelX, Frame1.Height / Screen.TwipsPerPixelY, 50, 50), True
Label3.Caption = "0"
kk = Val(Label3.Caption)
'level = 0
Option1_Click (Index)
Me.Move 2500, -100
End Sub

Private Sub Image2_Click()
Unload Me
End
End Sub

Private Sub Option1_Click(Index As Integer)
level = Index
lblTimer.Caption = IIf(level = 0, "Easy :", IIf(level = 1, "Midium :", "Hard :")) & " - " & "00:00:00"
End Sub

Private Sub RestMe_Click()
Command2_Click
End Sub

Private Sub Scramble_Click()
kk = 0
Command1_Click
End Sub

Private Sub SolveMe_Click()
Command3_Click
End Sub

Private Sub Timer2_Timer()

    Dim T As Date
    Dim M As Integer
    Dim S As Integer
  
    'increase elapsed time
    ElapsedSeconds = ElapsedSeconds + 1
    'show elapsed time in status window
    T = TimeSerial(0, 0, ElapsedSeconds)
    lblTimer.Caption = IIf(level = 0, "Easy :", IIf(level = 1, "Midium :", "Hard :")) & " - " & Format(T, "hh:nn:ss")

Check_Answer
End Sub

Sub Check_Answer()

If RubikCube1.GetCube = "RRRRRRRRRYYYYYYYYYPPPPPPPPPWWWWWWWWWBBBBBBBBBGGGGGGGGG" Then
    Timer2.Enabled = False
    MsgBox "You Have Complete My Cube!!" & _
        vbCrLf & lblTimer.Caption, vbInformation, "Congratulations"

    'reset lblTimer to default caption..
    lblTimer.Caption = IIf(level = 0, "Easy :", IIf(level = 1, "Midium :", "Hard :")) & " - " & "00:00:00"
' sndPlaySound App.Path & "\sound\won.wav", 1
End If

End Sub



