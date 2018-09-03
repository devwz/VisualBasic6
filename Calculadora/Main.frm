VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Calculadora"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3495
   LinkTopic       =   "Main"
   ScaleHeight     =   4455
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_equals 
      Caption         =   "="
      Height          =   615
      Left            =   2640
      TabIndex        =   20
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton bot_dot 
      Caption         =   "."
      Height          =   615
      Left            =   1800
      TabIndex        =   19
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton btn_0 
      Appearance      =   0  'Flat
      Caption         =   "0"
      Height          =   615
      Left            =   960
      TabIndex        =   18
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton btn_sum 
      Caption         =   "+"
      Height          =   615
      Left            =   2640
      TabIndex        =   16
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton btn_3 
      Caption         =   "3"
      Height          =   615
      Left            =   1800
      TabIndex        =   15
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton btn_2 
      Caption         =   "2"
      Height          =   615
      Left            =   960
      TabIndex        =   14
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton btn_1 
      Caption         =   "1"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton btn_sub 
      Caption         =   "-"
      Height          =   615
      Left            =   2640
      TabIndex        =   12
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton btn_6 
      Caption         =   "6"
      Height          =   615
      Left            =   1800
      TabIndex        =   11
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton btn_5 
      Caption         =   "5"
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton btn_4 
      Caption         =   "4"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton btn_mult 
      Caption         =   "X"
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton btn_9 
      Caption         =   "9"
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton btn_8 
      Caption         =   "8"
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton btn_7 
      Caption         =   "7"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton btn_div 
      Caption         =   "/"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton btn_clean 
      Caption         =   "C"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label label_display 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
