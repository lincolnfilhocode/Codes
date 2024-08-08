Attribute VB_Name = "LogOff"
Sub SAPLogoff()

    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object
       
    On Error Resume Next

    ' Usando sapguiserver em vez de sapgui, caso seja necessário
    Set SapGuiAuto = GetObject("SAPGUISERVER")

    ' Verifica se o SAPGUISERVER foi capturado
    If SapGuiAuto Is Nothing Then
        MsgBox "SAPGUISERVER não está disponível ou não pôde ser inicializado."
        Exit Sub
    End If

    ' Inicia o SAP GUI Scripting
    
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    
    If IsObject(SAPApp) Then
        Set SAPCon = SAPApp.Children(0)
    End If
    
    If IsObject(SAPCon) Then
    
        Do While SAPCon.Children.Count > 0
        
        Debug.Print SAPCon.Children.Count
        
        Set session = SAPCon.Children(0)
        
        session.findById("wnd[0]").Close
        
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        
        Loop
    End If
    
    ' Fechar o aplicativo SAP GUI completamente
    ' Isso fecha o processo do SAP GUI
    If Not SAPApp Is Nothing Then
        'Fechar sap pelo shell pelo nome
        Shell "taskkill /F /IM NWBC.exe", vbHide
    End If

        
End Sub

