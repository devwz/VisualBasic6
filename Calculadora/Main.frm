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
   Begin VB.Label main_label 
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

Dim operation As String
Dim tVal As Double
Dim fVal As Double

Private Sub bot_dot_Click()

    If Not InStr(1, main_label.Caption, ".") > 0 Then
        main_label.Caption = main_label.Caption + "."
    End If

End Sub

Private Sub btn_0_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If
    
    If main_label.Caption = "0" Then
        main_label.Caption = "0"
    Else
        main_label.Caption = main_label.Caption + "0"
    End If

End Sub

Private Sub btn_1_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If
    
    If main_label.Caption = "0" Then
        main_label.Caption = "1"
    Else
        main_label.Caption = main_label.Caption + "1"
    End If
    
End Sub

Private Sub btn_2_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If
    
    If main_label.Caption = "0" Then
        main_label.Caption = "2"
    Else
        main_label.Caption = main_label.Caption + "2"
    End If
    
End Sub

Private Sub btn_3_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If

    If main_label.Caption = "0" Then
        main_label.Caption = "3"
    Else
        main_label.Caption = main_label.Caption + "3"
    End If
    
End Sub

Private Sub btn_4_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If

    If main_label.Caption = "0" Then
        main_label.Caption = "4"
    Else
        main_label.Caption = main_label.Caption + "4"
    End If
    
End Sub

Private Sub btn_5_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If

    If main_label.Caption = "0" Then
        main_label.Caption = "5"
    Else
        main_label.Caption = main_label.Caption + "5"
    End If
    
End Sub

Private Sub btn_6_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If

    If main_label.Caption = "0" Then
        main_label.Caption = "6"
    Else
        main_label.Caption = main_label.Caption + "6"
    End If
    
End Sub

Private Sub btn_7_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If

    If main_label.Caption = "0" Then
        main_label.Caption = "7"
    Else
        main_label.Caption = main_label.Caption + "7"
    End If
    
End Sub

Private Sub btn_8_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If

    If main_label.Caption = "0" Then
        main_label.Caption = "8"
    Else
        main_label.Caption = main_label.Caption + "8"
    End If
    
End Sub

Private Sub btn_9_Click()

    If operation = "END" Then
        main_label.Caption = "0"
        operation = "WIP"
    End If

    If main_label.Caption = "0" Then
        main_label.Caption = "9"
    Else
        main_label.Caption = main_label.Caption + "9"
    End If
    
End Sub

Private Sub btn_clean_Click()
    main_label.Caption = "0"
    fVal = 0
    tVal = 0
End Sub

Private Sub btn_div_Click()
    tVal = Val(main_label.Caption)
    main_label.Caption = "0"
    operation = "DIV"
End Sub

Private Sub btn_mult_Click()
    tVal = Val(main_label.Caption)
    main_label.Caption = "0"
    operation = "MULT"
End Sub

Private Sub btn_sub_Click()
    tVal = Val(main_label.Caption)
    main_label.Caption = "0"
    operation = "SUB"
End Sub

Private Sub btn_sum_Click()
    tVal = Val(main_label.Caption)
    main_label.Caption = "0"
    operation = "SUM"
End Sub

Private Sub btn_equals_Click()

    Select Case operation
        Case "SUM"
            fVal = tVal + Val(main_label.Caption)
            Equals (fVal)
        Case "SUB"
            fVal = tVal - Val(main_label.Caption)
            Equals (fVal)
        Case "MULT"
            fVal = tVal * Val(main_label.Caption)
            Equals (fVal)
        Case "DIV"
            fVal = tVal / Val(main_label.Caption)
            Equals (fVal)
            
    End Select
    
    operation = "END"
    
End Sub

Private Sub Equals(ByVal fVal As Double)

    main_label.Caption = fVal
    
    tVal = fVal
    fVal = 0
    
End Sub
