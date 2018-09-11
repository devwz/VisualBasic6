Attribute VB_Name = "ThreadModule"

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Declare Function CreateThread Lib "kernel32.dll" ( _
    ByVal lpThreadAttributes As Long, _
    ByVal dwStackSize As Long, _
    ByVal lpStartAddress As Long, _
    lpParameter As Long, _
    ByVal dwCreationFlags As Long, _
    lpThreadID As Long) _
    As Long
    
' lpThreadAttributes - Atributos de seguran�a, normalmente definido para 0 para atributos padr�es
' dwStackSize - Tamanho da pilha, cada thread tem sua pr�pria pilha
' IpStartAddress - Endere�o de mem�ria onde o encadeamento � iniciado, deve ser um endere�o de uma fun��o em um m�dulo padr�o utilizando o operador AddressOf
' IpParameter - � um par�metro passado para a fun��o que inicia a thread
' dwCreatingFlags - Permite o controle do status do inicio da thread
' IpThreadID - Identificador de encadeamento da thread

' EXEMPLO
' Dim handle&
' handle& = CreateThread(0, 2000, AddressOf AAA, AAA, 0, AAA)

Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long


Public Function AsyncThread(ByVal obj As MyClass)

    Dim threadid As Long
    Dim handle&
    
    Dim threadparam As Long
    threadparam = ObjPtr(obj)
    
    handle = CreateThread(0, 2000, AddressOf BackgroundThread, threadparam, 0, threadid)
    
    If handle = 0 Then
        Exit Function
    End If
    
    CloseHandle handle
    
    AsyncThread = threadid

End Function


Public Sub BackgroundThread()

    For j = 0 To 30
        Debug.Print j
        Sleep 1000
    Next
    
    handle = 0

End Sub

Public Sub StopThreads()

    handle = 0

End Sub
