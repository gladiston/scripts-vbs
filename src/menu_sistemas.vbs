' Windows Script Host - VBScript
'-----------------------------------------------------------------
' Nome: menu_sistemas.vbs
' Proposito : Menu para exibir os programas do dia a dia do usuário
'   e evitar que o usuario use o explorer para explorar estas pastas
' Ideal para ser usados em ambientes onde as policies de segurança
'   impedem o usuário comuns de ter acesso a certos programas que
'   requeiram permissões especiais. Ou então evitar que os usuários
'   tenham que acesso ao explorer e possam ver/copiar pastas sensiveis
'   de sistemas e por ultimo tambem evitar que usem o cmd do windows
'   para rodar comandos. Para estes casos é bem melhor usar este 
'   script e só permitir o que desejamos.
' By: Gladiston Santana (gladiston.santana em gmail.com)
' Copyright: (c) Jun 2011, Todos os direitos reservados!
'-----------------------------------------------------------------
Option Explicit

Dim oFS, oWS, oWN
Set oWS = WScript.CreateObject("WScript.Shell")
Set oWN = WScript.CreateObject("WScript.Network")
Set oFS = WScript.CreateObject("Scripting.FileSystemObject")

Dim sBin, sRunAsPath, sRunAsOption, sRunAs, sUserName, sPassword, sDomain


' Aplicativo padrão
sBin = ""

' Dominio, se for o computador local use %computername%
sDomain = "meudominio"
' Nome do usuario com permissao de execução
sUsername = "sysdba"
' Senha, o til representa o CR+LF 
sPassword = "masterkey~"
' RunAsCmd = sintaxe do comando runas a ser executado
sRunAsPath = "runas"
'sRunAsOption = " /user:" & Chr(34) & sDomain & "\" & sUsername & Chr(34) & " /password:" & Chr(34) & sPassword & Chr(34) 
sRunAsOption = " /user:" & sDomain & "\" & sUsername 
	

'---------------
' Menu Principal
'---------------
do while sBin <> "sair"
	Select Case InputBox ( _
	 "Todos os programas alistados aqui" & vbCrlf & _
	 "estão protegidos contra execução externa." & vbCrlf &  vbCrlf & _
	 "Digite uma opção e então [OK] :" & vbCrlf & vbCrlf & _
	 " [1] PHD" & vbCrlf & _
	 " [2] Recursos Humanos" & vbCrlf & _
	 " [3] Contabil" & vbCrlf & _
	 " [4] DP" & vbCrlf & _
	 " [5] Fiscal" & vbCrlf & _
	 " [6] Ativo" & vbCrlf & _
	 " [7] NFe" & vbCrlf & _
	 " [8] Ciap" & vbCrlf & _
	 " [9] Guias" & vbCrlf & _
	 " [10] INSS" & vbCrlf & _
	 " [11] Médico" & vbCrlf & _
	 " [12] EDI" & vbCrlf & _
	 " [13] Alterdata Mail Configurador" & vbCrlf & _
	 " [14] Brasil Informatica (Legado)" & vbCrlf & _
	 " [15] Candidato" & vbCrlf & _     
	 " [98] Impressoras" & vbCrlf & _   
	 " --------------------------------------------------------------------------" & vbCrlf & _
	 " Qualquer outra opção finaliza este menu.", _ 
	 "Menu Principal")
	 Case "1"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Phd\wphd.exe"
	 Case "2"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Rh\Wrh.exe"
	 Case "3"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Contabil\Wcont.exe"
	 Case "4"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Dp\Wdp.exe"
	 Case "5"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Fiscal\Wfiscal.exe"
	 Case "6"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Ativo\WAtivo.exe"
	 Case "7"
	  sBin = "C:\scripts\nfe.cmd"
	 Case "8"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Ciap\WCIAP.exe"
	 Case "9"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Guias\Wguias.exe"
	 Case "10"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\INSS\winss.exe"
	 Case "11"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Rh\Medico.exe" 
	 Case "12"
	  sBin = "C:\Program Files (x86)\Alterdata\EDI\WEDI.EXE"
	 Case "13"
	  sBin = "C:\Program Files (x86)\Alterdata\EDI\MailCfg.exe"
	 Case "14"
	  sBin = "C:\scripts\brinfo.cmd"
	 Case "15"
	  sBin = "C:\Program Files (x86)\Alterdata\Pack\Rh\WCand.exe"
	 Case "98"
	  sBin = "C:\scripts\impressoras.cmd"
	 Case Else
	  sBin = "sair"
	End Select

	' Verificando se o aplicativo existe realmente
	if not oFS.FileExists( sBin ) and ( sBin <> "" ) and ( sBin <> "sair" ) Then
	  WScript.Echo "Aplicativo : " & Chr(34) & sBin  & Chr(34) & " não existe !" 
	Else
	  '----------
	  ' Executando o aplicativo escolhido
	  '----------
      sRunAs = sRunAsPath & " " & sRunAsOption & " " & Chr(34) & sBin & Chr(34)
	  ' Apenas Debug :
	  'WScript.Echo "Executando : " & vbCrlf & sRunAs
	  oWS.run( sRunAs )
	  WScript.Sleep 100
      'oWS.AppActivate "RunAs"
	  oWS.Sendkeys sPassword	  
	End If
loop

'-------------------
' Limpeza e Saida
'-------------------
Call LimpezaESair

'-----------------------------------------------------------------
' SubRotinas
'-----------------------------------------------------------------

'---------------------
Sub LimpezaESair()
  Set oWS = Nothing
  Set oWN = Nothing
  Set oFS = Nothing
  WScript.Quit
End Sub
