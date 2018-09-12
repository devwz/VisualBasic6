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

Declare Function TerminateThread Lib "kernel32.dll" (ByVal handle As Long, ByVal dwExitCode As Long) As Long
Declare Function CloseHandle Lib "kernel32.dll" (ByVal handle As Long) As Long

Public Function Thread(ByVal obj As Util)

    Dim threadID As Long
    Dim handle As Long
    
    Dim threadParam As Long
    threadParam = ObjPtr(obj)
    
    handle = CreateThread(0, 0, AddressOf ToDoSomething, threadParam, 0, threadID)
    
    If handle = 0 Then
        Exit Function
    End If
    
    CloseHandle handle
    
    Thread = threadID

End Function

Public Sub ToDoSomething(ByVal param As Long)

        Dim tlog As Util
        Set tlog = New Util
        
        Do While Not tlog.Log(20)
        Loop

End Sub
