Função para Deslogar e fechar o SAP

Attribute VB_Name = "LogOff"
Sub SAPLogoff()

    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object
       
    On Error Resume Next

    Captura servidor
    Set SapGuiAuto = GetObject("SAPGUISERVER")

    ' Verifica se o SAPGUISERVER foi capturado
    If SapGuiAuto Is Nothing Then
        Exit Sub
    End If

    ' Inicia o SAP GUI Scripting
    Set SAPApp = SapGuiAuto.GetScriptingEngine

    Obtem conexao
    If IsObject(SAPApp) Then
        Set SAPCon = SAPApp.Children(0)
    End If

    Se tiver conexoes, fecha
    If IsObject(SAPCon) Then
    
        Do While SAPCon.Children.Count > 0
            Set session = SAPCon.Children(0)
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Loop

    End If
    
    ' Fechar o SAP
    If Not SAPApp Is Nothing Then
        'Fechar sap pelo shell pelo nome // matando o processo
        Shell "taskkill /F /IM NWBC.exe", vbHide
    End If
   
End Sub

