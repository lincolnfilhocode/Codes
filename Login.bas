Attribute VB_Name = "Login"
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
    SapGuiPath = "C:\Program Files\SAP\NWBC800\NWBC.exe"  'Alterar para localização do executavel de sap
    Set SapGuiAuto = GetObject("SAPGUISERVER")
    
    'Verifica se sap já está aberto
    If SapGuiAuto Is Nothing Then
    
        'Inicializa SAP Chr(34) = "
        Set WsShell = CreateObject("WScript.Shell")
        WsShell.Run Chr(34) & SapGuiPath & Chr(34)
        
        'Timer de 60s para esperar abertura//ajustar para tempo maximo de abertura para pior computador
        SapGuiReady = False
        StartTime = Timer
        Timeout = 60
        
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
   
    'Login no servidor
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    
    If IsObject(SAPApp) Then
        Set SAPCon = SAPApp.OpenConnection("0318 - SA - PB0 - [ERP] (001)", True) 'Ajustar para servidor utilizado
    End If
    
    If IsObject(SAPCon) Then
    Set session = SAPCon.Children(0)
    End If
    
    'Descarregando variaveis
    Set SAPApp = Nothing
    Set SAPCon = Nothing
    Set session = Nothing

End Sub

