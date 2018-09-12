VERSION 5.00
Begin VB.Form MainWindow 
   Caption         =   "MainWindow"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_create_thread_C 
      Caption         =   "Criar Thread C"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btn_create_thread_B 
      Caption         =   "Criar Thread B"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btn_create_thread_A 
      Caption         =   "Criar Thread A"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_create_thread_A_Click()

    Dim logUtil As Util
    Set logUtil = New Util
    
    logUtil.strLog = "Thread A"
    
    Thread logUtil

End Sub

Private Sub btn_create_thread_B_Click()

    Dim logUtil As Util
    Set logUtil = New Util
    
    logUtil.strLog = "Thread B"
    
    Thread logUtil

End Sub

Private Sub btn_create_thread_C_Click()

    Dim logUtil As Util
    Set logUtil = New Util
    
    logUtil.strLog = "Thread C"
    
    Thread logUtil

End Sub
