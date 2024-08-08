
'Codigo para realização de login em servidor sap scripting
'Attribute VB_Name = "Login"

Sub Login()

    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object
    Dim SapGuiPath As String
    Dim WshShell As Object
    Dim StartTime As Double
    Dim Timeout As Double
    Dim SapGuiReady As Boolean
    
    On Error Resume Next
    
    'Inicialização de variaveis
    SapGuiPath = "diretorio/NWBC.exe"  'Utilize o diretorio de seu executavel
    Set SapGuiAuto = GetObject("SAPGUISERVER")
    
    'Verifica se sap já está aberto
    If SapGuiAuto Is Nothing Then
    
        'Inicializa SAP // Chr(34) = "
        Set WsShell = CreateObject("WScript.Shell")
        WsShell.Run Chr(34) & SapGuiPath & Chr(34)
        
        'Timer de 60s para esperar abertura //Ajuste para tempo maximo do gargalo do computador
        SapGuiReady = False
        StartTime = Timer
        Timeout = 60

        'Tentar conectar ao sap apos abertura por tempo = timeout
        Do While Not SapGuiReady
            On Error Resume Next
            Set SapGuiAuto = GetObject("SAPGUISERVER")
            On Error GoTo 0
            
            If Not SapGuiAuto Is Nothing Then
                SapGuiReady = True
            ElseIf Timer - StartTime > Timeout Then
                MsgBox "Programa não inicializado"
                Exit Sub
            End If
            DoEvents
        Loop
        
   End If
   
    'Conexao com script
    Set SAPApp = SapGuiAuto.GetScriptingEngine

    Conexao com servidor
    If IsObject(SAPApp) Then
        Set SAPCon = SAPApp.OpenConnection("Seu servidor", True) 'Ajustar para servidor utilizado
    End If

    'Iniciando sessao
    If IsObject(SAPCon) Then
    Set session = SAPCon.Children(0)
    End If
    
    'Descarregando variaveis
    Set SAPApp = Nothing
    Set SAPCon = Nothing
    Set session = Nothing

End Sub

