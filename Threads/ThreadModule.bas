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
    
' lpThreadAttributes - Atributos de segurança, normalmente definido para 0 para atributos padrões
' dwStackSize - Tamanho da pilha, cada thread tem sua própria pilha
' IpStartAddress - Endereço de memória onde o encadeamento é iniciado, deve ser um endereço de uma função em um módulo padrão utilizando o operador AddressOf
' IpParameter - É um parâmetro passado para a função que inicia a thread
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
