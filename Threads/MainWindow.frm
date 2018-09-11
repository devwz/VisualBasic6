VERSION 5.00
Begin VB.Form MainWindow 
   Caption         =   "MainWindow"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_create_thread 
      Caption         =   "Criar Thread"
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

Private Sub btn_create_thread_Click()

    Set c = New MyClass
    AsyncThread c

End Sub


Private Sub btn_finalizar_Click()

    ThreadModule.StopThreads

End Sub
