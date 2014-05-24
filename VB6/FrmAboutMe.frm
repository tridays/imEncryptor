VERSION 5.00
Begin VB.Form FrmAboutMe 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "关于"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Label BG 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "徐普"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(auto)"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2013-10-15"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "作者"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "版本"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "imEncryptor"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1965
   End
End
Attribute VB_Name = "FrmAboutMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const vTransparent As Byte = 200

Private Sub BG_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Line (0, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1), vbBlack, B
    BG.Move 0, 0, Me.ScaleWidth - 1, Me.ScaleHeight - 1
    Label7 = App.Major & "." & App.Minor & "." & App.Revision
End Sub
