' Solicita o número da etiqueta ao usuário
Dim etiqueta
etiqueta = InputBox("Digite o numero da etiqueta:", "Número da Etiqueta")

' Verifica se o usuário inseriu algo
If etiqueta = "" Then
    MsgBox "Você não inseriu nenhum numero. O script será encerrado.", vbExclamation, "Etiqueta não fornecida"
    WScript.Quit
End If

' 1. Adiciona a etiqueta ao Registro do Windows
On Error Resume Next
Dim shell, registryPath
Set shell = CreateObject("WScript.Shell")

registryPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\"
shell.RegWrite registryPath & "SupportPhone", etiqueta, "REG_SZ"

If Err.Number <> 0 Then
    MsgBox "Erro ao escrever no Registro do Windows. Execute o script como Administrador ou consulte o artigo no moviedesk", vbCritical, "Erro no Registro"
    WScript.Quit
Else
    MsgBox "Número da etiqueta registrado com sucesso no Registro do Windows.", vbInformation, "Registro atualizado"
End If

On Error GoTo 0

' 2. Instala o GLPI Agent a partir de um arquivo local
Dim installerPath, objShell
Set objShell = CreateObject("WScript.Shell")

' Caminho do instalador local - MODIFIQUE para o caminho correto
installerPath = "GLPI-Agent.msi"

If Not CreateObject("Scripting.FileSystemObject").FileExists(installerPath) Then
    MsgBox "O arquivo de instalação do GLPI Agent não foi encontrado no caminho especificado: " & installerPath, vbCritical, "Erro no Instalador"
    WScript.Quit
End If

' Executa a instalação silenciosa
On Error Resume Next
objShell.Run "msiexec /i " & Chr(34) & installerPath & Chr(34) & " /quiet /norestart", 0, True

If Err.Number <> 0 Then
    MsgBox "Erro ao instalar o GLPI Agent. Execute o script como Administrador.", vbCritical, "Erro na Instalação"
    WScript.Quit
Else
    MsgBox "GLPI Agent instalado com sucesso.", vbInformation, "Instalação Completa"
End If
On Error GoTo 0

' 3. Configura o GLPI Agent para apontar para o servidor e preencher otherserial
Dim configFilePath, objFSO, configFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

configFilePath = "C:\Program Files\GLPI-Agent\etc\agent.cfg"

If objFSO.FileExists(configFilePath) Then
    Set configFile = objFSO.OpenTextFile(configFilePath, 1)
    Dim fileContent
    fileContent = configFile.ReadAll
    configFile.Close

    ' Adiciona ou atualiza as configurações
    If InStr(fileContent, "server=") > 0 Then
        fileContent = Replace(fileContent, "server=", "server=https://nvirtual.with18.glpi-network.cloud/front/inventory.php")
    Else
        fileContent = fileContent & vbCrLf & "server=https://nvirtual.with18.glpi-network.cloud/front/inventory.php"
    End If

    If InStr(fileContent, "otherserial") > 0 Then
        fileContent = Replace(fileContent, "otherserial=", "otherserial=" & etiqueta)
    Else
        fileContent = fileContent & vbCrLf & "otherserial=" & etiqueta
    End If

    ' Salva as alterações no arquivo
    Set configFile = objFSO.OpenTextFile(configFilePath, 2)
    configFile.Write fileContent
    configFile.Close

    MsgBox "Configurações do GLPI Agent atualizadas com sucesso.", vbInformation, "Configuração Completa"

    ' Reinicia o serviço do GLPI Agent
    objShell.Run "net stop glpi-agent && net start glpi-agent", 0, True
Else
    MsgBox "Arquivo de configuração do GLPI Agent não encontrado. Verifique a instalação.", vbCritical, "Erro na Configuração"
End If
