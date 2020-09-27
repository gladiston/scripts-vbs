' Windows Script Host - VBScript
'-----------------------------------------------------------------
' Nome: backup_dev.vbs
' Proposito : Realização de Backup da estação de Desenvolvimento
'   ou de pastas/arquivos considerados muito importantes
' Trata-se dum backup diferencial usando 7zip, o script quando 
'   executado pela primeira vez fará um backup completo, porém 
'   na vez seguinte criará um novo arquivo igual ao original apenas
'   acrescido um sufixo indicando timestamp com data e hora.
' Se for necessario a restauração, basta restaurar o completo e 
'   depois o ultimo diferencial realizado.
' Este script é muito melhor do que muitas ferramentas de backup com
'   o objetivo de backup diferencial porque é facil, rapido e confiável.
'   Existem programas de backup por aqui que não conseguem estrapolar o
'   limite de path do Windows e este script faz isso.
' By: Gladiston Santana (gladiston.santana em gmail.com)
' Copyright: (c) Jun 2011, Todos os direitos reservados!
'-----------------------------------------------------------------
' ATENÇÃO
' Acrescente os usuários que executarão este script nessas policies, caso contrário, 
' eles não terão permissão de realizar backup e script falhará:
' Vá no aplicativo Administrativos\Diretiva de segurança Local
' Diretivas locais \ Atribuição de direitos de usuário \ Fazer backup de arquivos e direcórios
' Diretivas locais \ Atribuição de direitos de usuário \ Restaurar arquivos e direcórios
'
' E nessas duas opções acrescente membros e grupos que poderão realizar esta tarefas, 
'    administradores e operadores de cópia já estão incluídos
' Para rodar este script em servidores normalmente protegidos pelo UAC, não é possivel 
' executar com duplo clique, neste caso roda com:
' c:\windows\system32\wscript.exe backup_dev.vbs //B //T:3600
Option Explicit
Dim oFS, oWS, oWN
Dim grouplistD,ADSPath,userPath,listGroup
Dim sResultado, iSemParar, sInicio, sFim
Dim sTempVar
Dim sLogDir
Dim sLogFile
Dim sCMD
Dim sDest_Drive
Dim sDest_Drive_Leiame
Dim sDest_logs 
Dim sMensagem
Dim sBackup_Lista
Dim sPasta_Origem, sPasta_Destino
Dim sDomain, sUsername, sPassword, RespostaSimNao
Dim sBaseName
Dim sMembro_de
Dim bMembro_de_SN
Dim objFile
Dim p7z
Dim sIgnore_Files
Dim q
Dim iPos
Dim sDelim

Set oWS = WScript.CreateObject("WScript.Shell")
Set oWN = WScript.CreateObject("WScript.Network")
Set oFS = WScript.CreateObject("Scripting.FileSystemObject")

'
' definicoes do usuario, modifique a vontade
'

' Copia com mensagens parando para dar "OK" use "0", se pretende automatizar 
' jogando no agendador de tarefas  use "1"
iSemParar=1

' Pasta de log
'sLogDir=Wscript.ScriptFullName  ' mesmo local que o script ou ...
sLogDir="c:\WinSrv\Scripts"

' Drive para onde vai o Backup, pode ser tambem o ponto de montagem
sDest_Drive="\\arca1\desenv\backup-desenv\src"

' Localizacao do 7zip
p7z="C:\Program Files\7-Zip\7z.exe"

' Verifica se é necessário que o usuaario que executa é membro de algum grupo
'  se não desejar esta verificação deixe em branco, ie. sMembro_de=""
sMembro_de="Administradores"

' Aspas
q=Chr(34)

' Delimitador usado para separar origem e destino, use > a menos que tenham algum problema
sDelim=">"

' Data atual no formato AAAAMMDDHHMMSS
sInicio = DataEHora(Now(),True)

' Nome fixo desse script sem a extensao, será reutilizado depois
' Este nome base tambem definirá o nome dos arquivos de log e Leiame
sBaseName=GetFilenameWithoutExtension(Wscript.ScriptName)

' Arquivo de Manifesto, se ele não existir em sDest_Drive, o backup nao 
' se inicia. O proposito disso é poder usar drive externos/rede e evitar 
' que uma cópia vá para o drive externo/rede errado
sDest_Drive_Leiame="Leiame_" & sBaseName & ".txt"

If Not oFS.FolderExists(sDest_Drive) Then
  oFS.CreateFolder sDest_Drive
End If

' Este Backup vai manter backups de dias anteriores
sDest_logs = sDest_Drive & "\"
If Not oFS.FolderExists(sDest_logs) Then 
  oFS.CreateFolder sDest_logs
  If Not oFS.FolderExists(sDest_logs) Then   
    Call LimpezaESair
  End If  
End If

' localização inicial do arquivo de log, mas quando gera arquivos de 
' backup, gera um log especifico para essa operacao no mesmo diretorio
'sLogFile = EnvString("temp") & "\backup.log"
sLogFile = sLogDir & "\" & sBaseName & ".log"

' Criando um arquivo contendo lista de extensoes indesejadas
' Atenção: use mascara apenas para nome de arquivos
'  NÃO USE mascaras ou espaços em branco para nome de pastas
sIgnore_Files = "c:\WinSrv\Scripts\" & sBaseName & "-ignore_files.txt"
If Not oFS.FileExists(sIgnore_Files) Then
  AddToLog "*.~*", sIgnore_Files
  AddToLog "~*.*", sIgnore_Files
  AddToLog "*.??~", sIgnore_Files
  AddToLog "*.$*", sIgnore_Files
  AddToLog "*.??$", sIgnore_Files
  AddToLog "*.dcu", sIgnore_Files
  AddToLog "*.log", sIgnore_Files
  AddToLog "*.dcu", sIgnore_Files
  AddToLog "*.bak", sIgnore_Files
  AddToLog "*.fbk", sIgnore_Files    
  AddToLog "*.fdb", sIgnore_Files   
  AddToLog "*.fdb.reverter", sIgnore_Files   
  AddToLog "*.tmp", sIgnore_Files
  AddToLog "*.reverso", sIgnore_Files
  AddToLog "Superado", sIgnore_Files
  AddToLog "lixo", sIgnore_Files
  AddToLog "_lixo", sIgnore_Files
  AddToLog "__history", sIgnore_Files
  AddToLog "__recovery", sIgnore_Files
  AddToLog "Snapshots", sIgnore_Files 
  AddToLog "_RESTORE", sIgnore_Files
  AddToLog "MSOCache", sIgnore_Files
  AddToLog "Downloads", sIgnore_Files
  AddToLog "Temp", sIgnore_Files
  AddToLog "Tmp", sIgnore_Files
  AddToLog "Debug", sIgnore_Files
  AddToLog "Recycled", sIgnore_Files
  AddToLog "RECYCLER", sIgnore_Files
  AddToLog "Thumbs.db", sIgnore_Files
  AddToLog "desktop.ini", sIgnore_Files 
End If

' Drive/Pastas a serem copiados seguido do delimitador e o 
' local de destino que for informado após o delimitador.
' Se for o local de destino padrao, definido por "sDest_Drive" 
'   então use "def" ou "default".
sBackup_Lista=Array( _
    "C:\Fontes" & sDelim & "def", _
	  "C:\Vidy15" & sDelim & "def", _
	  "C:\Vidy11" & sDelim & "def", _
	  "C:\SESMT" & sDelim & "def", _
    "C:\Develop" & sDelim & "def" _
	)
  
' O usuario que esta executando este script tem permissão de backup?
bMembro_de_SN=0
if sMembro_de<>"" Then
  If isMember(sMembro_de) Then
     bMembro_de_SN=1
  End If
End If

'
' Daqui em diante é melhor não mexer se não souber o que está fazendo
'

'--------------------
' Inicio do programa
'--------------------

' Aviso importante
sMensagem = _  
  "Pronto para iniciar o procedimento de Backup :" & vbCrlf &_
  "Inicio :" & sInicio & vbCrlf &_
  "Destino da cópia : " & sDest_Drive & vbCrlf &_
  "Nome base : " & sBaseName & vbCrlf &_
  "Log : " & sLogFile & vbCrlf &_
  "Lista de arquivos ignorados : " & sIgnore_Files & vbCrlf &_
  "Arquivo-manifesto : " & sDest_Drive & "\" & sDest_Drive_Leiame & vbCrlf 

If bMembro_de_SN>0 Then
  sMensagem =  sMensagem & "Membro de '" & sMembro_de & "' : Sim"
Else
  sMensagem =  sMensagem & "Membro de '" & sMembro_de & "' : Não"
End If

AddToLog sMensagem, sLogFile 

If  ( iSemParar<=0 ) Then 
  RespostaSimNao = MsgBox (sMensagem, vbYesNo, "Confirmação:")
  If RespostaSimNao =  vbNo Then
     Call LimpezaESair
  End IF
End If

If Not oFS.FileExists(sDest_Drive & "\" & sDest_Drive_Leiame) Then
  sMensagem = "Cadê o arquivo: " & vbCrlf &_
    sDest_Drive & "\" & sDest_Drive_Leiame & "?" & vbCrlf &_
    "Sem este arquivo de manifestor na unidade de destino não poderei prosseguir."
  AddToLog sMensagem, sLogFile 
  if  ( iSemParar<=0 ) Then 
    WScript.Echo( sMensagem )
  End If 
  Call LimpezaESair
End If

' Pastas que foram especificadas no inicio deste script serão transferidos
' para seu local informado de destino
For Each sTempVar in sBackup_Lista
  iPos = InStr(1, sTempVar, sDelim)
  sMensagem = ""
  If iPos > 0 Then  'Possui o delimitador
    sPasta_Origem  = Trim(Left(sTempVar, iPos-1))
    sPasta_Destino = Trim(Right(sTempVar, Len(sTempVar)-iPos))
    if ((LCase(sPasta_Destino)="def") or (LCase(sPasta_Destino)="default")) Then
      sPasta_Destino=sDest_Drive
    End If
    If not oFS.FolderExists(sPasta_Origem) Then
      sMensagem = sMensagem & "Pasta de origem não existe: " & sPasta_Origem & vbCrlf
    End If
    If not oFS.FolderExists(sPasta_Destino) Then
      sMensagem = sMensagem & "Pasta de destino não existe: " & sPasta_Origem & vbCrlf
    End If

    if sMensagem="" Then
      'WScript.Echo( sPasta_Origem & vbCrlf & sPasta_Destino )
      Call DoBackup(sPasta_Origem, sPasta_Destino)
    Else
      AddToLog sMensagem, sLogFile 
    End If
  End If  
Next

' Aproveitando o momento para fazer o rebuild de icones
DoRebuild_Icon_Cache

' Aviso concludente
sFim = DataEHora(Now(), True)  
sMensagem =  vbCrlf & "Backup finalizado :" & vbCrlf &_
             "Inicio :" & sInicio & vbCrlf &_
             "Termino :" & sFim 
AddToLog sMensagem, sLogFile 
 
If  ( iSemParar<=0 ) Then 
  WScript.Echo( sMensagem & vbCrlf & "Clique em [OK] para prosseguir."  )
End If

'--------------------
' Finaliza o programa
'--------------------
Call LimpezaESair


'-----------------------------------------------------------------
' SubRotinas
'-----------------------------------------------------------------
Sub LimpezaESair()
  Set oWS = Nothing
  Set oWN = Nothing
  Set oFS = Nothing
  WScript.Quit
End Sub

Sub DoBackup(AOrigem, APastaDestino)
  Dim sDrive, sOrigem2
  Dim iPos, sCmd
  Dim sCopiarPara
  Dim sCopiarParaDif
  Dim sNovoNome
  Dim bToUpdate
  Dim sValidChars
  Dim sChar
  Dim sTempVar
  bToUpdate=0
 ' Caracteres validos para nome de arquivos 
  sValidChars="abcdefghijklmnopqrstuvwxyz"
  sValidChars=sValidChars & UCase(sValidChars)
  sValidChars=sValidChars & "01234567890_+"
  
  ' Novo Nome
  ' 1. Remove a letra de drive
  sTempVar=AOrigem
  iPos = InStr(1, sTempVar, ":")
  If iPos > 0 Then  'Drive existe
      sTempVar = Right(sTempVar, Len(sTempVar)-iPos)
  End If
  '2. Trocar barra por underline
  sTempVar=Replace(sTempVar, "\\", "_") 
  sTempVar=Replace(sTempVar, "\", "_") 
  sTempVar=Replace(sTempVar, "  ", "_") 
  sTempVar=Replace(sTempVar, " ", "_") 
  sTempVar=Replace(sTempVar, "__", "_") 
  '3. Nao começa com underline
  Do While (Left(sTempVar, 1) = "_")
    sTempVar=Right(sTempVar, Len(sTempVar)-1) 
  Loop
  '4. Remove espacos e caracteres considerados invalidos
  sNovoNome=""
  For iPos = 1 To Len(sTempVar)
    sChar=Mid(sTempVar, iPos, 1)
    if InStr(sValidChars, sChar)>0 Then
      sNovoNome=sNovoNome & sChar
    End If
  Next
  
  ' Deve terminar com barra "\"
  sCopiarPara=APastaDestino
  if (Right(sCopiarPara,1) <> "\") Then
      sCopiarPara = sCopiarPara & "\"  
  End If
  sCopiarPara = sCopiarPara & "\" & sNovoNome & "\"
  ' Se a pasta não existir entao crio uma
  If Not oFS.FolderExists(sCopiarPara) Then 
    oFS.CreateFolder sCopiarPara
    If Not oFS.FolderExists(sCopiarPara) Then   
       AddToLog "Não foi possivel criar a pasta: " & sCopiarPara, sLogFile 
      Call LimpezaESair
    End If  
  End If  
  
  sCopiarPara=sCopiarPara & sNovoNome & ".7z"
  sLogFile=sCopiarPara & ".log"
  sCopiarParaDif=""
  If oFS.FileExists(sCopiarPara) Then 
     sCopiarParaDif=sCopiarPara & sNovoNome & "-" & sInicio & ".7z"
     sLogFile=sCopiarParaDif & ".log"
     bToUpdate=1
  End If


  if bToUpdate Then  
    sCmd = q & p7z & q & " u " & q & sCopiarPara & q & " " & q & AOrigem & q &_
      " -u- -up0q3r2x2y2z0w2!" & sCopiarParaDif 
  Else
    sCmd = q & p7z & q & " u " & q & sCopiarPara & q & " " & q & AOrigem & q 
  End If
  
  If oFS.FileExists(sIgnore_Files) Then
    sCmd = sCmd & " -xr@" & sIgnore_Files 
  End If
  
  AddToLog sCmd, sLogFile 
  oWS.run sCmd, 1, True
  If Err.Number <> 0 Then
	  sMensagem = vbTab & "  Falhou :" & sCmd & vbCrlf & _
	    "Código do Erro: " & Err.Number & vbCrlf & _
	    "Código do Erro (Hex): " & Hex(Err.Number) & vbCrlf & _
	    "Fonte: " &  Err.Source & vbCrlf & _
	    "Descrição do Erro: " &  Err.Description
	  AddToLog sMensagem, sLogFile 
    if  ( iSemParar<=0 ) Then 
	    WScript.Echo( sMensagem )
    End If 
    if  ( iSemParar>0 ) Then 	
	    sMensagem = vbTab & "Sucesso : " & sCmd
	    AddToLog sMensagem, sLogFile
    End If	
	End If    
  Err.Clear   
  
End Sub

Function AddToLog(AText, ALogFile)
  Const ForAppending = 8
  Dim sUseThisLog
  Dim sUseThisText
  Dim sCheck_Folder
  AddToLog=0  
  sUseThisText= AText
  sUseThisLog = ALogFile
  if sUseThisLog = "" Then
    sUseThisLog = sLogFile
  End If
  'sCheck_Folder=oFS.GetFolder(sUseThisLog).Name
  'WScript.Echo( sCheck_Folder & ": " & ALogFile )
  if sUseThisLog <> "" Then
    'If oFS.FolderExists(sCheck_Folder) Then 
      set objFile = oFS.OpenTextFile(sUseThisLog, ForAppending, True)
      objFile.WriteLine(AText)
      objFile.Close
      AddToLog=1
    'End If
  End If  
End Function

Function DataEHora(sData, sExibeHoras)
Dim Resultado
  Resultado = Year(DateValue(sData)) & "-" &_
	TwoDigits(Month(DateValue(sData))) & "-" &_
	TwoDigits(Day(DateValue(sData)))
  if sExibeHoras=True Then Resultado=Resultado & "+" &_	
	TwoDigits(Hour(sData)) & "h" & _
	TwoDigits(Minute(sData)) & "m" & _
	TwoDigits(Second(sData)) & "s"
	
  DataEHora = Resultado 
End Function

Function TwoDigits(num)
    If(Len(num)=1) Then
        TwoDigits="0"&num
    Else
        TwoDigits=num
    End If
End Function

'This function checks to see if the passed group name contains the current
' user as a member. Returns True or False
Function IsMember(groupName)
    If IsEmpty(groupListD) then
        Set groupListD = CreateObject("Scripting.Dictionary")
        groupListD.CompareMode = 1
        ADSPath = EnvString("userdomain") & "/" & EnvString("username")
        Set userPath = GetObject("WinNT://" & ADSPath & ",user")
        For Each listGroup in userPath.Groups
            groupListD.Add listGroup.Name, "-"
        Next
    End if
    IsMember = CBool(groupListD.Exists(groupName))
End Function

'This function returns a particular environment variable's value.
' for example, if you use EnvString("username"), it would return
' the value of %username%.
Function EnvString(variable)
    variable = "%" & variable & "%"
    EnvString = oWS.ExpandEnvironmentStrings(variable)
End Function

' GetFilenameWithoutExtension - Essa extensão é apenas para remover a extensao de um arquivo
'  para ficar apenas o nome base
Function GetFilenameWithoutExtension(ByVal FileName)
  Dim Result, i
  Result = FileName
  i = InStrRev(FileName, ".")
  If ( i > 0 ) Then
    Result = Mid(FileName, 1, i - 1)
  End If
  GetFilenameWithoutExtension = Result
End Function

Sub DoRebuild_Icon_Cache()
  Dim sCmd
  Dim localappdata  
  Dim iconcache
  Dim iconcache_x
  Dim sPathToDelete
  Dim oFolder
  On error resume next

  localappdata=EnvString("localappdata")
  iconcache=localappdata & "\IconCache.db"
  iconcache_x=localappdata & "\Microsoft\Windows\Explorer"

  ' Temporariamente matamos o explorer porque ele mantem o cache de icones aberto
  if oWS.FileExists(localappdata) Then
	  sCmd="taskkill /f /im explorer.exe"
	  AddToLog sCmd, sLogFile 
	  oWS.run sCmd, 1, True
	  ' Apagamos o database de icones 
      oFS.DeleteFile iconcache, True

	  ' Apagamos o cache de icones aberto  
	  set oFolder = oFS.getFolder(iconcache_x)   
	  if oFolder.Files.Count <> 0 then 
		  oFS.DeleteFile iconcache_x & "\iconcache*.*", True
		  If Err.Number <> 0 Then
			  sMensagem = vbTab & "  Falhou :" & sCmd & vbCrlf & _
				"Código do Erro: " & Err.Number & vbCrlf & _
				"Código do Erro (Hex): " & Hex(Err.Number) & vbCrlf & _
				"Fonte: " &  Err.Source & vbCrlf & _
				"Descrição do Erro: " &  Err.Description
			  AddToLog sMensagem, sLogFile 
			if  ( iSemParar<=0 ) Then 
				WScript.Echo( sMensagem )
			End If 
		  End If    
		  Err.Clear
	  End If
	  ' Recarregamos o explorer que pela falta dos arquivos que foram apagados irá reconstituí-los
	  sCmd="c:\Windows\explorer.exe"
	  AddToLog sCmd, sLogFile 
	  oWS.run sCmd, 1, True
  End If  
End Sub