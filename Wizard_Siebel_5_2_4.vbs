'-----------------------------------------------------------
'|  Wizard Siebel 5.2									    |
'|															|
'| - - - ChangeLog - - -									|
'| - Nova Analise de Credito	 *							|
'| - XML Portabilidade		*								|
'| - Gerar Arquivo Associacao Pedido Venda *				|
'| - SVN			*										|
'| - Gerar XML - Contigencia #P - Tratamento Erro 	     	|
'| - Código Oi Vende
'|															|
'|  DESENVOLVEDOR: JOSE ARTEIRO TEIXEIRA (JOSE.CAVALCANTI)	|
'------------------------------------------------------------

VAR_WIZARD_SIEBEL_VERSION = "5.2.5"
VAR_GLOBAL_SIEBEL_DEBUG   = "FALSE"		'TRUE OR FALSE

FUNCTION LoginSenha(ByVal VAR_GLOBAL_SIEBEL, ByVal VAR_GLOBAL_PARAM)
	'VAR_GLOBAL_SIEBEL, VAR_GLOB_SBl_LOG, VAR_GLOB_SBl_PAS, VAR_GLOBAL_ARBOR
	ConecArray = Array( _
		Array("SBPCSTI8" , "LOGIN" , "SENHA"        , "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.58.193.81)(PORT=1553))(CONNECT_DATA=(SERVICE_NAME=fxct2ti)))", "TI8"), _
		Array("SBPCSTI4" , "LOGIN" , "SENHA"        , ""                                                                                                         , "TI3"), _
		Array("SBPCSTI3" , "LOGIN" , "SENHA"        , ""                                                                                                         , "TS"), _
		Array("SBPCSTI6" , "LOGIN" , "SENHA"        , ""                                                                                                         , "TS2"), _
		Array("SBPCSTI1" , "LOGIN" , "SENHA"        , ""                                                                                                         , "TI1"), _
		Array("SBPCSDV8" , "LOGIN" , "SENHA"        , ""                                                                                                         , "DV8"), _
		Array("SBPCSTI4" , "LOGIN" , "SENHA"        , ""                                                                                                         , "DV4"), _
		Array("SBDV10"   , "LOGIN" , "SENHA"        , ""                                                                                                         , "DV5"), _
		Array("SBPCSTI2" , "LOGIN" , "SENHA"        , ""                                                                                                         , "B1"), _
		Array("SBPCSDV7" , "LOGIN" , "SENHA"        , ""                                                                                                         , "B2"), _
		Array("SBPCSDV"  , "LOGIN" , "SENHA"        , ""                                                                                                         , "B10"), _
		Array("SBPCSDV3" , "LOGIN" , "SENHA"        , ""                                                                                                         , "TS_"), _
		Array("SBPCSINT" , "LOGIN" , "SENHA"      , ""                                                                                                         , "BCV INT"), _
		Array("SBPCSINF" , "LOGIN" , "SENHA" , ""                                                                                                         , "BCV INF"), _
		Array("SBPCSPRD" , "LOGIN" , "SENHA"     , ""                                                                                                         , "PRD"), _
		Array("SBPCSHML" , "LOGIN" , "SENHA"        , ""                                                                                                         , "HML"), _
		Array("SBDTGUARD", "LOGIN" , "SENHA"      , ""                                                                                                         , "DTGUARD"), _
		Array("SIEBELLOCAL", "LOGIN"  , "SENHA"        , ""                                                                                                         , "LOCAL") _
	)
	For Each Connect In ConecArray
		If Connect(0) = VAR_GLOBAL_SIEBEL Then
			Select Case VAR_GLOBAL_PARAM
				Case "VAR_GLOB_SBl_LOG"
					LoginSenha = Connect(1)
				Case "VAR_GLOB_SBl_PAS"
					LoginSenha = Connect(2)
				Case "VAR_GLOBAL_ARBOR"
					LoginSenha = Connect(3)
				Case Else
					LoginSenha = ""
			End Select
		End If
		
		If VAR_GLOBAL_SIEBEL = "ALL" Then
			If Len(TextLine) > 0 Then
				TextLine = TextLine & ";"
			End If
			
			TextLine = TextLine & Connect(0)
			For Aux = Len(Connect(0)) to 10
				TextLine = TextLine & " "
			Next
			TextLine = TextLine & " - " & Connect(4)
		End If
	Next
	If VAR_GLOBAL_SIEBEL = "ALL" Then
		LoginSenha = TextLine
	End If
END FUNCTION

FUNCTION MyOIChipDV(ByVal ASSET_NUM)	
			VAR_ASSET_NUM_1  = MID(ASSET_NUM, 18, 1)
			VAR_ASSET_NUM_2  = MID(ASSET_NUM, 17, 1)
			VAR_ASSET_NUM_3  = MID(ASSET_NUM, 16, 1)
			VAR_ASSET_NUM_4  = MID(ASSET_NUM, 15, 1)
			VAR_ASSET_NUM_5  = MID(ASSET_NUM, 14, 1)
			VAR_ASSET_NUM_6  = MID(ASSET_NUM, 13, 1)
			VAR_ASSET_NUM_7  = MID(ASSET_NUM, 12, 1)
			VAR_ASSET_NUM_8  = MID(ASSET_NUM, 11, 1)
			VAR_ASSET_NUM_9  = MID(ASSET_NUM, 10, 1)
			VAR_ASSET_NUM_10 = MID(ASSET_NUM, 9, 1)
			VAR_ASSET_NUM_11 = MID(ASSET_NUM, 8, 1)
			VAR_ASSET_NUM_12 = MID(ASSET_NUM, 7, 1)
			VAR_ASSET_NUM_13 = MID(ASSET_NUM, 6, 1)
			VAR_ASSET_NUM_14 = MID(ASSET_NUM, 5, 1)
			VAR_ASSET_NUM_15 = MID(ASSET_NUM, 4, 1)
			VAR_ASSET_NUM_16 = MID(ASSET_NUM, 3, 1)
			VAR_ASSET_NUM_17 = MID(ASSET_NUM, 2, 1)
			VAR_ASSET_NUM_18 = MID(ASSET_NUM, 1, 1)
		
			IF (VAR_ASSET_NUM_1*2) >= 10 THEN CAL_ASSET_NUM_1   = (1+((VAR_ASSET_NUM_1*2)-10)) ELSE CAL_ASSET_NUM_1   = (VAR_ASSET_NUM_1*2) END IF
			CAL_ASSET_NUM_2  = VAR_ASSET_NUM_2
			IF (VAR_ASSET_NUM_3*2) >= 10 THEN CAL_ASSET_NUM_3   = (1+((VAR_ASSET_NUM_3*2)-10)) ELSE CAL_ASSET_NUM_3   = (VAR_ASSET_NUM_3*2) END IF
			CAL_ASSET_NUM_4  = VAR_ASSET_NUM_4
			IF (VAR_ASSET_NUM_5*2) >= 10 THEN CAL_ASSET_NUM_5   = (1+((VAR_ASSET_NUM_5*2)-10)) ELSE CAL_ASSET_NUM_5   = (VAR_ASSET_NUM_5*2) END IF
			CAL_ASSET_NUM_6  = VAR_ASSET_NUM_6
			IF (VAR_ASSET_NUM_7*2) >= 10 THEN CAL_ASSET_NUM_7   = (1+((VAR_ASSET_NUM_7*2)-10)) ELSE CAL_ASSET_NUM_7   = (VAR_ASSET_NUM_7*2) END IF
			CAL_ASSET_NUM_8  = VAR_ASSET_NUM_8
			IF (VAR_ASSET_NUM_9*2) >= 10 THEN CAL_ASSET_NUM_9   = (1+((VAR_ASSET_NUM_9*2)-10)) ELSE CAL_ASSET_NUM_9   = (VAR_ASSET_NUM_9*2) END IF
			CAL_ASSET_NUM_10 = VAR_ASSET_NUM_10
			IF (VAR_ASSET_NUM_11*2) >= 10 THEN CAL_ASSET_NUM_11 = (1+((VAR_ASSET_NUM_11*2)-10)) ELSE CAL_ASSET_NUM_11 = (VAR_ASSET_NUM_11*2) END IF
			CAL_ASSET_NUM_12 = VAR_ASSET_NUM_12
			IF (VAR_ASSET_NUM_13*2) >= 10 THEN CAL_ASSET_NUM_13 = (1+((VAR_ASSET_NUM_13*2)-10)) ELSE CAL_ASSET_NUM_13 = (VAR_ASSET_NUM_13*2) END IF
			CAL_ASSET_NUM_14 = VAR_ASSET_NUM_14
			IF (VAR_ASSET_NUM_15*2) >= 10 THEN CAL_ASSET_NUM_15 = (1+((VAR_ASSET_NUM_15*2)-10)) ELSE CAL_ASSET_NUM_15 = (VAR_ASSET_NUM_15*2) END IF
			CAL_ASSET_NUM_16 = VAR_ASSET_NUM_16
			IF (VAR_ASSET_NUM_17*2) >= 10 THEN CAL_ASSET_NUM_17 = (1+((VAR_ASSET_NUM_17*2)-10)) ELSE CAL_ASSET_NUM_17 = (VAR_ASSET_NUM_17*2) END IF
			CAL_ASSET_NUM_18 = VAR_ASSET_NUM_18
			
			SOMA_CAL_ASSET_NUM = CAL_ASSET_NUM_1 + CAL_ASSET_NUM_2 + CAL_ASSET_NUM_3 + CAL_ASSET_NUM_4 + CAL_ASSET_NUM_5 
			SOMA_CAL_ASSET_NUM = SOMA_CAL_ASSET_NUM + CAL_ASSET_NUM_6 + CAL_ASSET_NUM_7 + CAL_ASSET_NUM_8 + CAL_ASSET_NUM_9 
			SOMA_CAL_ASSET_NUM = SOMA_CAL_ASSET_NUM + CAL_ASSET_NUM_10 + CAL_ASSET_NUM_11 + CAL_ASSET_NUM_12 + CAL_ASSET_NUM_13 
			SOMA_CAL_ASSET_NUM = SOMA_CAL_ASSET_NUM + CAL_ASSET_NUM_14 + CAL_ASSET_NUM_15 + CAL_ASSET_NUM_16 + CAL_ASSET_NUM_17 + CAL_ASSET_NUM_18
			
			LAST_SOMA_CAL_ASSET_NUM = RIGHT(SOMA_CAL_ASSET_NUM, 1)
			
			IF LAST_SOMA_CAL_ASSET_NUM = 0 THEN DV_ASSET_NUM = 0 ELSE DV_ASSET_NUM = (10-LAST_SOMA_CAL_ASSET_NUM)
			
			MyOIChipDV = DV_ASSET_NUM			
END FUNCTION

FUNCTION XML_OiTotal (ByVal VAR_STR_MENU, ByVal VAR_STR_MSISDN, ByVal VAR_STR_ROWID)
	TextLine = "WF: PCS #P - Receber Notificacao Status" & VbCrLf & VbCrLf & VbCrLf
		
	If VAR_STR_MENU <> "" Then
		For varTMP = 1 to 2
			If varTMP = 1 Then
				TextLine = TextLine & "### Fixo ###" & VbCrLf
			Elseif varTMP = 2 Then
				TextLine = TextLine & "### Velox ###" & VbCrLf
			End If

			TextLine = TextLine & _
				"<?xml version=""1.0"" encoding=""UTF-8""?>" & VbCrLf & _
				"<SiebelMessage MessageId="""" MessageType=""Integration Object"" IntObjectName=""PCS Notificacao Status - In"" xmlns:xmm=""http://namespace.softwareag.com/entirex/xml/mapping"" xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & VbCrLf & _
				" <ListOfNotifStatusIn>" & VbCrLf & _
				"	<DadosNotif>" & VbCrLf
			
			If varTMP = 1 Then
				TextLine = TextLine & "  		<Sistema>STC</Sistema>" & VbCrLf
			Elseif varTMP = 2 Then
				TextLine = TextLine & "  		<Sistema>STCDADOS</Sistema>" & VbCrLf
			End If
				
			TextLine = TextLine & _
				"  		<NumIdentTvFixo>" & VAR_STR_MSISDN(0) & "</NumIdentTvFixo>" & VbCrLf & _
				"  		<IdVenda>" & VAR_STR_ROWID &"</IdVenda>" & VbCrLf & _
				"  		<DtStatus>" & Date & " " & Time & "</DtStatus>" & VbCrLf & _
				"  		<Status>" & VAR_STR_MENU & "</Status>" & VbCrLf & _
				"  		<PermiteAgendam>N</PermiteAgendam>" & VbCrLf & _
				"  		<VelocidadeADSL></VelocidadeADSL>" & VbCrLf & _
				"  		<AgendamObrigatorio>N</AgendamObrigatorio>" & VbCrLf & _
				"  		<IdFibra></IdFibra>" & VbCrLf & _
				"  		<Detail></Detail>" & VbCrLf & _
				"  		<Reason></Reason>" & VbCrLf & _
				"	</DadosNotif>" & VbCrLf & _
				" </ListOfNotifStatusIn>" & VbCrLf & _
				"</SiebelMessage>"
									
			TextLine = TextLine & VbCrLf & VbCrLf
		Next
			
		VAR_STR_PROD_NAME = QUERY_INPUT("SELECT LISTAGG(S_PROD_INT.NAME, '|') WITHIN GROUP (ORDER BY S_PROD_INT.X_CLASSE) FROM SIEBEL.cx_shared_bundle JOIN SIEBEL.S_QUOTE_ITEM ON CX_SHARED_BUNDLE.ROW_ID = S_QUOTE_ITEM.SD_ID JOIN SIEBEL.S_PROD_INT ON S_PROD_INT.ROW_ID = S_QUOTE_ITEM.PROD_ID WHERE CX_SHARED_BUNDLE.ROW_ID = '" & VAR_STR_ROWID & "' AND S_PROD_INT.X_CLASSE = 'Oi TV' AND S_QUOTE_ITEM.STAT_CD <> 'Inativo' GROUP BY S_PROD_INT.X_CLASSE;", 1)				
		
		If VAR_STR_PROD_NAME <> "" Then 
		
			VAR_STR_MENU2 = VAR_STR_MENU
		
			If VAR_STR_MENU = "ABV" Then VAR_STR_MENU = "OS Aberta"
			If VAR_STR_MENU = "ATV" Then VAR_STR_MENU = "OS Finalizada"
			If VAR_STR_MENU = "CAA" Then VAR_STR_MENU = "Impossivel Instalar"
			If VAR_STR_MENU = "CAN" Then VAR_STR_MENU = "Impossivel Instalar"
			If VAR_STR_MENU = "DTV" Then VAR_STR_MENU = "OS Finalizada"
		
			If VAR_STR_MENU = "OS Aberta" Then VAR_STR_MENU2 = "Em Provisionamento" 
			If VAR_STR_MENU = "Impossivel Instalar" Then VAR_STR_MENU2 = "Erro"
			If VAR_STR_MENU = "OS Finalizada" And VAR_STR_MENU2 = "ATV" Then VAR_STR_MENU2 = "Ativo"
			If VAR_STR_MENU = "OS Finalizada" And VAR_STR_MENU2 = "DTV" Then VAR_STR_MENU2 = "Inativo"
		
			VAR_STR_PROD_NAME = Split(VAR_STR_PROD_NAME,"|")
			
			TextLine = TextLine & "### TV ###" & VbCrLf
			TextLine = TextLine & _
				"<?xml version=""1.0"" encoding=""UTF-8""?>" & VbCrLf & _
				"<SiebelMessage MessageId="""" MessageType=""Integration Object"" IntObjectName=""PCS Notificacao Status - In"" xmlns:xmm=""http://namespace.softwareag.com/entirex/xml/mapping"" xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & VbCrLf & _
				" <ListOfNotifStatusIn>" & VbCrLf & _
				"	<DadosNotif>" & VbCrLf & _
				"  		<Sistema>sinn</Sistema>" & VbCrLf & _
				"  		<NumIdentTvFixo>" & VAR_STR_MSISDN(0) & "</NumIdentTvFixo>" & VbCrLf & _
				"  		<IdVenda>" & VAR_STR_ROWID &"</IdVenda>" & VbCrLf & _
				"  		<DtStatus>" & Date & " " & Time & "</DtStatus>" & VbCrLf & _
				"  		<Status>" & VAR_STR_MENU & "</Status>" & VbCrLf & _
				"  		<PermiteAgendam></PermiteAgendam>" & VbCrLf & _
				"  		<AgendamObrigatorio></AgendamObrigatorio>" & VbCrLf & _
				"  		<IdFibra></IdFibra>" & VbCrLf & _
				"  		<Detail></Detail>" & VbCrLf & _
				"  		<Reason></Reason>" & VbCrLf & _
				"  		<VelocidadeADSL />" & VbCrLf & _
				"  		<ListOfProdutosTV>" & VbCrLf
				
			VAR_STR_TV_SIEBEL = "PRODUTOS:" & VbCrLf
				
			For each PROD_NOME in VAR_STR_PROD_NAME
			
				PROD_SIEBEL = QUERY_INPUT("SELECT X_NOME_SIEBEL FROM SIEBEL.S_PROD_INT WHERE NAME = '" & PROD_NOME & "';", 1)
				
				VAR_STR_TV_SIEBEL = VAR_STR_TV_SIEBEL & "- " & PROD_NOME & " -> " & PROD_SIEBEL & VbCrLf
			
				TextLine = TextLine & _
					"  			<ProdutosTV>" & VbCrLf & _
					"  				<CodProduto>" &	PROD_NOME & "</CodProduto>" & VbCrLf & _
					"  				<StatusProduto>" & VAR_STR_MENU2 & "</StatusProduto>" & VbCrLf & _
					"  			</ProdutosTV>" & VbCrLf
			Next
				
			TextLine = TextLine & _
				"  		</ListOfProdutosTV>" & VbCrLf & _
				"	</DadosNotif>" & VbCrLf & _
				" </ListOfNotifStatusIn>" & VbCrLf & _
				"</SiebelMessage>"

			TextLine = TextLine & VbCrLf & VbCrLf & VAR_STR_TV_SIEBEL
		End If
		XML_OiTotal =  TextLine & VbCrLf & VbCrLf & VbCrLf
	End If
END FUNCTION

FUNCTION XML_OCT(ByVal VAR_STR_MENU, ByVal VAR_STR_MSISDN, ByVal VAR_STR_ROWID)
	TextLine = "WF: PCS Recebe Fixo STC" & VbCrLf & VbCrLf & VbCrLf
	TextLine = TextLine & "<SiebelMessage MessageId="""" MessageType=""Integration Object"" IntObjectName=""PCS Recebe Fixo STC"">" & VbCrLf
    TextLine = TextLine & "<ListOfPcsRecFxSTC>" & VbCrLf
    TextLine = TextLine & "  <PCSRecFxSTC>" & VbCrLf
    TextLine = TextLine & "    <Id/>" & VbCrLf
    TextLine = TextLine & "    <Created/>" & VbCrLf
    TextLine = TextLine & "    <Updated/>" & VbCrLf
    TextLine = TextLine & "    <ConflictId/>" & VbCrLf
    TextLine = TextLine & "    <Acao>" & VAR_STR_MENU & "</Acao>" & VbCrLf
    TextLine = TextLine & "    <DDD>" & Left(VAR_STR_MSISDN(0), 2) &"</DDD>" & VbCrLf
    TextLine = TextLine & "    <DDDNovoNumFixo/>" & VbCrLf
    TextLine = TextLine & "    <DataOperacao>" & Date & " " & Time & "</DataOperacao>" & VbCrLf
    TextLine = TextLine & "    <NovoNumFixo/>" & VbCrLf
    TextLine = TextLine & "    <NumFixo>" & Mid(VAR_STR_MSISDN(0), 3) & "</NumFixo>" & VbCrLf
    TextLine = TextLine & "  </PCSRecFxSTC>" & VbCrLf
    TextLine = TextLine & "</ListOfPcsRecFxSTC>" & VbCrLf
    TextLine = TextLine & "</SiebelMessage>"
	XML_OCT = TextLine
END FUNCTION

FUNCTION XML_Legado()
	VAR_STR_ROWID=InputBox("Informe o ROW_ID do Bundle:","Siebel 6.3 - XML Notificacao Status Legado", "2-")
	VAR_STR_MSISDN = QUERY_INPUT("SELECT X_NUM_FIXO_CONVERGENTE || ';' || s_prod_int.PREF_CARRIER_CD || ';' || PD2.PREF_CARRIER_CD FROM siebel.cx_shared_bundle JOIN siebel.s_prod_int ON S_PROD_INT.X_CHAVE_PRODUTO = CX_SHARED_BUNDLE.X_PLANO LEFT JOIN siebel.s_prod_int pd2 ON pd2.X_CHAVE_PRODUTO = CX_SHARED_BUNDLE.X_NOVO_PLANO WHERE cx_shared_bundle.row_id = '" & VAR_STR_ROWID & "';", 1)
	
	If VAR_STR_ROWID <> "" And VAR_STR_MSISDN <> "" Then
	
		VAR_STR_MSISDN = Split(VAR_STR_MSISDN, ";")
		
		If Len(VAR_STR_MSISDN(0)) = 0 Then
			VAR_STR_MSISDN(0)=InputBox("Informe um FIXO para o Bundle:","Siebel 6.3 - XML Notificacao Status Legado", "21")
		End If
	
		VAR_STR_MENU = 	"Selecione um Status:" & VbCrLf &_
					"" & VbCrLf &_
					"-- #P --" & VbCrLf &_
					"ABV - OS Aberta" & VbCrLf &_
					"ATV - Em Ativação -> Ativo" & VbCrLf &_
					"DTV - Confirmação de Desativação ou Desassociação do STC" & VbCrLf &_
					"RET - Confirmação de Retirada de Fixo" & VbCrLf &_
					"CAN - ERRO FIXO"   & VbCrLf &_
					"CAA - ERRO VELOX"  & VbCrLf &_
					"DOA - PORTABILIDADE DOADORA DO FIXO" & VbCrLf &_
					"CHG - MUDAR NUMERO DO FIXO (XML COM O NOVO NUMERO)" & VbCrLf &_
					"ENA - MUDANÇA VELOX SEM RETORNO DE VELOCIDADE MENOR" & VbCrLf &_
					"ETV - MUDANÇA DE VELOX COM RETORNO DE VELOCIDADE MENOR" & VbCrLf &_
					"" & VbCrLf &_
					"-- OCT --" & VbCrLf &_
					"RET - Confirmação de Retirada de Fixo" & VbCrLf &_
					"DBQ - Desbloqueio do STC" & VbCrLf &_
					"" & VbCrLf &_
					"Especifico:"
 
		VAR_STR_MENU=InputBox(VAR_STR_MENU,"Siebel 6.3 - XML Notificacao Status Legado","ABV")
		
		If VAR_STR_MSISDN(2) = "1P" Or VAR_STR_MSISDN(2) = "2P" Or VAR_STR_MSISDN(2) = "3P" Or VAR_STR_MSISDN(2) = "4P" Then
			TextLine = XML_OiTotal(VAR_STR_MENU, VAR_STR_MSISDN, VAR_STR_ROWID)
		
		ElseIf VAR_STR_MSISDN(2) = "OCT" Then
			TextLine = TextLine & XML_OCT(VAR_STR_MENU, VAR_STR_MSISDN, VAR_STR_ROWID)
			
		Elseif VAR_STR_MSISDN(1) = "1P" Or VAR_STR_MSISDN(1) = "2P" Or VAR_STR_MSISDN(1) = "3P" Or VAR_STR_MSISDN(1) = "4P" Then
			TextLine = XML_OiTotal(VAR_STR_MENU, VAR_STR_MSISDN, VAR_STR_ROWID)
			
		Else
			TextLine = TextLine & XML_OCT(VAR_STR_MENU, VAR_STR_MSISDN, VAR_STR_ROWID)
		End If
				
		Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
			Set OTF = FSO.OpenTextFile("XML_DV_OUTPUT.txt", 8, True)
				OTF.WriteLine TextLine
				OTF.Close
			Set OTF = Nothing
		Set FSO = Nothing
			
		Set objShell = CreateObject("WScript.Shell")
			Set objExecObject = objShell.Exec("notepad.exe XML_DV_OUTPUT.txt")
				'varTMP = objExecObject.StdOut.ReadAll
				'DEBUG(varTMP)
			Set objExecObject = Nothing
		Set objShell = Nothing
			
		WScript.Sleep 5000
		REMOVER_INPUT("XML_DV_OUTPUT.txt")
		
	Else
		MsgBox("ROW_ID Invalido")
	End If
END FUNCTION

FUNCTION Viabilidade_Velox()

	VAR_STR_MENU = 	"Selecione um Velocidade:" & VbCrLf &_
					"" & VbCrLf &_
					"01. 300KB" & VbCrLf &_
					"02. 600KB" & VbCrLf &_
					"03. 1MB"   & VbCrLf &_
					"04. 10MB"  & VbCrLf &_
					"05. 15MB"  & VbCrLf &_
					"06. 20MB"  & VbCrLf &_
					"07. 25MB"  & VbCrLf &_
					"08. 30MB"  & VbCrLf &_
					"09. 35MB"  & VbCrLf &_
					"10. 40MB"  & VbCrLf &_
					"11. 50MB"  & VbCrLf &_
					"12. 60MB"  & VbCrLf &_
					"13. 70MB"
 
	VAR_STR_MENU=InputBox(VAR_STR_MENU,"Siebel 6.3 - Viabilidade Velox",4)
	
	'VAR_STR_MENU2 = "Configurar Velocidade Velox?" & VbCrLf & "" & VbCrLf & "01. SIM" & VbCrLf & "" & VbCrLf & "02. NAO"
 
	'VAR_STR_MENU2=InputBox(VAR_STR_MENU2,"Siebel 6.3 - Viabilidade Velox",1)
	
	VAR_STR_MENU2 = MsgBox("Configurar Velocidade Velox?", "36", "Viabilidade Velox")
	
	If IsNumeric(VAR_STR_MENU) And IsNumeric(VAR_STR_MENU2) Then
	
		VAR_STR_SQL = "UPDATE siebel.cx_shared_bundle SET " 
		
		VAR_STR_ROWID=InputBox("Informe o ROW_ID do Bundle:","Siebel 6.3 - Viabilidade Velox", "2-")
		TextLine = QUERY_INPUT("SELECT COUNT(*) AS TOTAL FROM siebel.cx_shared_bundle WHERE row_id = '" & VAR_STR_ROWID & "';", 1)
		
		If TextLine >= 1 And VAR_STR_ROWID <> "" And VAR_STR_ROWID <> "2-" Then
			Select Case VAR_STR_MENU
				Case 1 
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '300KB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '11' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 2
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '600KB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '31' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 3
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '1MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '51' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 4
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '10MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '111' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 5
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '15MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '131' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 6
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '20MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '151' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 7
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '25MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '171' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 8
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '30MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '191' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 9
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '35MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '211' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 10
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '40MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '231' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 11
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '50MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '251' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 12
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '60MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '271' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 13
					If VAR_STR_MENU2 <> 7 Then
						VAR_STR_SQL = VAR_STR_SQL & "x_velocidade = '70MB', "
					End If
					TextLine = QUERY_INPUT(VAR_STR_SQL & "x_viabilidade = '291' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case Else
					MsgBox("Opcao Invalida.")
			End Select
		Else
			MsgBox("ROW_ID Invalido")
		End If
	Else
		MsgBox("Opcao Invalida. Somente numeros.")
	End If
	
END FUNCTION

FUNCTION Analise_Credito()
	VAR_STR_MENU = 	"Selecione um Tipo de Plano:" & VbCrLf &_
					"" & VbCrLf &_
					"1. Pos"  & VbCrLf &_
					"2. 2P"  & VbCrLf &_
					"3. 3P"  & VbCrLf &_
					"4. 4P"  & VbCrLf &_
					"5. 4P + TV"  & VbCrLf &_
					"6. Portabilidade"
 
	VAR_STR_MENU=InputBox(VAR_STR_MENU,"Siebel 6.3 - Analise de Credito",5)
 
	VAR_STR_MENU2=InputBox("Linhas Solicitadas:","Siebel 6.3 - Analise de Credito",10)
	
	If IsNumeric(VAR_STR_MENU) Then
		
		VAR_STR_SQL = "UPDATE siebel.s_finan_prof SET " & _
					  "x_dt_validade_credito = SYSDATE + 60, " & _ 
					  "x_qte_max_msisdn = " & VAR_STR_MENU2 & "," & _
					  "x_qte_msisdn_solicitados = '" & VAR_STR_MENU2 & "'," & _
					  "credit_score = 0, " & _
					  "x_vlr_limite_credito = 5000, " & _
					  "x_cod_classe_risco = 'A', " & _
					  "x_descr_decisao = 'Pós-Pago', " & _
					  "type_cd = '2 - Análise 2ª Etapa', " & _
					  "x_motivo_decisao = 'Pós-Pago', " & _
					  "x_cod_plano_preco = 'Z', " & _
					  "x_num_proposta = '" & Day(date) & Month(date) & Year(date) & "'"
					  
					  '"X_QTE_MSISDN_ATIVOS = '10', " & _
					  '"X_COD_PDV_SAP = '0001001622', " & _
					  '"X_IND_ANALISE_COMPLETA = 'Y', " & _
					  '"X_QTD_MAX_AVULSO = '10', " & _
					  '"X_QTD_AVULSO_PEDIDOS = '10', " & _
					  '"X_ANALISE_PAGGO = 'N', " & _
					  '"X_CONTRATO_TV = '', " & _
					  '"X_NUM_FIXO = '', " & _
					  '"X_EXPIRADA = 'N', " & _
					  '"X_FIXO = 'Novo', " & _
					  '"X_MOVEL = 'Novo', " & _
					  '"X_TV = 'Novo', " & _
					  '"X_VELOX = 'Novo'"
					  
		
		VAR_STR_ROWID=InputBox("Informe o ROW_ID da Analise:","Siebel 6.3 - Analise de Credito", "2-")
		TextLine = QUERY_INPUT("SELECT COUNT(*) AS TOTAL  FROM siebel.s_finan_prof WHERE row_id = '" & VAR_STR_ROWID & "';", 1)
		
		If TextLine >= 1 And VAR_STR_ROWID <> "" And VAR_STR_ROWID <> "2-" Then
			Select Case VAR_STR_MENU
				Case 1 
						TextLine = QUERY_INPUT(VAR_STR_SQL & ", x_tipo_analise_aparelho = 'Ativação com Aparelho Oi', COLL_STAGE_CD = 'N', X_PORTABILIDADE = 'N' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 2
						TextLine = QUERY_INPUT(VAR_STR_SQL & ", x_tipo_analise_aparelho = 'Ativação sem Aparelho Oi', COLL_STAGE_CD = 'Y', X_PORTABILIDADE = 'N' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 3
						TextLine = QUERY_INPUT(VAR_STR_SQL & ", x_tipo_analise_aparelho = 'Ativação sem Aparelho Oi', COLL_STAGE_CD = 'Y', X_PORTABILIDADE = 'N' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 4
						TextLine = QUERY_INPUT(VAR_STR_SQL & ", x_tipo_analise_aparelho = 'Ativação com Aparelho Oi', COLL_STAGE_CD = 'N', X_PORTABILIDADE = 'N' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 5
						TextLine = QUERY_INPUT(VAR_STR_SQL & ", x_tipo_analise_aparelho = 'Ativação com Aparelho Oi', COLL_STAGE_CD = 'Y', X_PORTABILIDADE = 'N' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case 6
						TextLine = QUERY_INPUT(VAR_STR_SQL & ", x_tipo_analise_aparelho = 'Ativação com Aparelho Oi', COLL_STAGE_CD = 'N', X_PORTABILIDADE = 'Y' where row_id = '" & VAR_STR_ROWID & "';", 1)
				Case Else
					MsgBox("Opcao Invalida.")
			End Select
		Else
			MsgBox("ROW_ID Invalido")
		End If
	Else
		MsgBox("Opcao Invalida. Somente numeros.")
	End If

END FUNCTION

FUNCTION DV_Generator()
	ASSET_NUM = InputBox("Qual o numero do seu Oi Chip:", "Oi Chip")
		If IsNumeric(ASSET_NUM) Then
			If Len(ASSET_NUM) <> 18 then
				MsgBox("Numero de caractere invalido!")
										
			else
				DV_ASSET_NUM = MyOIChipDV(ASSET_NUM)
				MsgBox("DV (Oi Chip): " & DV_ASSET_NUM)
			end If
		else
			MsgBox("Opcao Invalida. Somente numeros.")
		end if
END FUNCTION

FUNCTION MyDVCPF(Byval VAR_CPF)
	VAR_CPF_1  = MID(VAR_CPF, 1, 1) * 10
	VAR_CPF_2  = MID(VAR_CPF, 2, 1) * 9
	VAR_CPF_3  = MID(VAR_CPF, 3, 1) * 8
	VAR_CPF_4  = MID(VAR_CPF, 4, 1) * 7
	VAR_CPF_5  = MID(VAR_CPF, 5, 1) * 6
	VAR_CPF_6  = MID(VAR_CPF, 6, 1) * 5
	VAR_CPF_7  = MID(VAR_CPF, 7, 1) * 4
	VAR_CPF_8  = MID(VAR_CPF, 8, 1) * 3
	VAR_CPF_9  = MID(VAR_CPF, 9, 1) * 2
	
	VAR_SOMA   = VAR_CPF_1 + VAR_CPF_2 + VAR_CPF_3 + VAR_CPF_4 + VAR_CPF_5 + VAR_CPF_6 + VAR_CPF_7 + VAR_CPF_8 + VAR_CPF_9
	VAR_SOMA   = VAR_SOMA MOD 11
	
	IF VAR_SOMA < 2 THEN 
		VAR_SOMA = 0 
	ELSE
		VAR_SOMA = 11 - VAR_SOMA
	END IF
	
	VAR_CPF_1  = MID(VAR_CPF, 1, 1) * 11
	VAR_CPF_2  = MID(VAR_CPF, 2, 1) * 10
	VAR_CPF_3  = MID(VAR_CPF, 3, 1) * 9
	VAR_CPF_4  = MID(VAR_CPF, 4, 1) * 8
	VAR_CPF_5  = MID(VAR_CPF, 5, 1) * 7
	VAR_CPF_6  = MID(VAR_CPF, 6, 1) * 6
	VAR_CPF_7  = MID(VAR_CPF, 7, 1) * 5
	VAR_CPF_8  = MID(VAR_CPF, 8, 1) * 4
	VAR_CPF_9  = MID(VAR_CPF, 9, 1) * 3
	VAR_CPF_10 = VAR_SOMA * 2
	VAR_CPF    = VAR_SOMA
	
	VAR_SOMA   = VAR_CPF_1 + VAR_CPF_2 + VAR_CPF_3 + VAR_CPF_4 + VAR_CPF_5 + VAR_CPF_6 + VAR_CPF_7 + VAR_CPF_8 + VAR_CPF_9 + VAR_CPF_10
	VAR_SOMA   = VAR_SOMA MOD 11
	
	IF VAR_SOMA < 2 THEN 
		VAR_SOMA = 0 
	ELSE
		VAR_SOMA = 11 - VAR_SOMA
	END IF
	
	VAR_CPF = VAR_CPF & VAR_SOMA
	MyDVCPF = VAR_CPF
END FUNCTION

FUNCTION MyCPF(VAR_STR_MENU)
	Randomize
	VAR_CPF = Int((9-1+1)*Rnd+1)
			
	FOR VAR_COUNT = 2 TO 8
		Randomize
		VAR_CPF = VAR_CPF & Int((9-0+1)*Rnd+0)
	NEXT
			
	VAR_DV_CPF = MyDVCPF(VAR_CPF & VAR_STR_MENU)
			
	IF Len(VAR_DV_CPF) < 2 Then 
		VAR_DV_CPF = String(2 - Len(VAR_DV_CPF), "0") & VAR_DV_CPF
	END IF
			
	MyCPF = VAR_CPF & VAR_STR_MENU & VAR_DV_CPF	
END FUNCTION

FUNCTION VAL_MyCPF(VAR_CPF)
	VAR_CONTACT = QUERY_INPUT("SELECT COUNT(*) AS VALOR FROM SIEBEL.S_CONTACT WHERE SOC_SECURITY_NUM = '" & VAR_CPF & "';", 1)
	If VAR_CONTACT >= "1" Then VAL_MyCPF = TRUE Else VAL_MyCPF = FALSE
END FUNCTION

FUNCTION CPF_Generator()
	VAR_STR_MENU = 	"Selecione um Estado:" & VbCrLf &_
					"" & VbCrLf &_
					"1. DF, GO, MS e TO"  & VbCrLf &_
					"2. PA, AM, AC, AP, RO e RR"  & VbCrLf &_
					"3. CE, MA e PI"  & VbCrLf &_
					"4. PE, RN, PB e AL"  & VbCrLf &_
					"5. BA e SE"  & VbCrLf &_
					"6. MG"  & VbCrLf &_
					"7. RJ e ES"  & VbCrLf &_
					"8. SP"  & VbCrLf &_
					"9. PR e SC"  & VbCrLf &_
					"0. RS"
 
	VAR_STR_MENU=InputBox(VAR_STR_MENU,"Siebel 6.3 - Oi Menu",7)
	
	If IsNumeric(VAR_STR_MENU) Then
		If VAR_STR_MENU >= 0 AND VAR_STR_MENU <=9 THEN
			Do
				VAR_CPF = MyCPF(VAR_STR_MENU)

			Loop While VAL_MyCPF(VAR_CPF) = TRUE
			
			VAR_CPF = InputBox("CPF", "Siebel 6.3 - Gera CPF",VAR_CPF)
		Else
			MsgBox("Estado Invalido.")
		End If
	Else
		MsgBox("Opcao Invalida. Somente numeros.")
	End If
END FUNCTION

FUNCTION MyDVCNPJ(Byval VAR_CNPJ)	
	VAR_CNPJ_1  = MID(VAR_CNPJ, 1, 1)  * 5
	VAR_CNPJ_2  = MID(VAR_CNPJ, 2, 1)  * 4
	VAR_CNPJ_3  = MID(VAR_CNPJ, 3, 1)  * 3
	VAR_CNPJ_4  = MID(VAR_CNPJ, 4, 1)  * 2
	VAR_CNPJ_5  = MID(VAR_CNPJ, 5, 1)  * 9
	VAR_CNPJ_6  = MID(VAR_CNPJ, 6, 1)  * 8
	VAR_CNPJ_7  = MID(VAR_CNPJ, 7, 1)  * 7
	VAR_CNPJ_8  = MID(VAR_CNPJ, 8, 1)  * 6
	VAR_CNPJ_9  = MID(VAR_CNPJ, 9, 1)  * 5
	VAR_CNPJ_10 = MID(VAR_CNPJ, 10, 1) * 4
	VAR_CNPJ_11 = MID(VAR_CNPJ, 11, 1) * 3
	VAR_CNPJ_12 = MID(VAR_CNPJ, 12, 1) * 2
	
	VAR_SOMA   = VAR_CNPJ_1 + VAR_CNPJ_2 + VAR_CNPJ_3 + VAR_CNPJ_4 + VAR_CNPJ_5 + VAR_CNPJ_6 + VAR_CNPJ_7 + VAR_CNPJ_8 + VAR_CNPJ_9 + VAR_CNPJ_10 + VAR_CNPJ_11 + VAR_CNPJ_12
	VAR_SOMA   = VAR_SOMA MOD 11
	
	IF VAR_SOMA < 2 THEN 
		VAR_SOMA = 0 
	ELSE
		VAR_SOMA = 11 - VAR_SOMA
	END IF
	
	VAR_CNPJ_1  = MID(VAR_CNPJ, 1, 1)  * 6
	VAR_CNPJ_2  = MID(VAR_CNPJ, 2, 1)  * 5
	VAR_CNPJ_3  = MID(VAR_CNPJ, 3, 1)  * 4
	VAR_CNPJ_4  = MID(VAR_CNPJ, 4, 1)  * 3
	VAR_CNPJ_5  = MID(VAR_CNPJ, 5, 1)  * 2
	VAR_CNPJ_6  = MID(VAR_CNPJ, 6, 1)  * 9
	VAR_CNPJ_7  = MID(VAR_CNPJ, 7, 1)  * 8
	VAR_CNPJ_8  = MID(VAR_CNPJ, 8, 1)  * 7
	VAR_CNPJ_9  = MID(VAR_CNPJ, 9, 1)  * 6
	VAR_CNPJ_10 = MID(VAR_CNPJ, 10, 1) * 5
	VAR_CNPJ_11 = MID(VAR_CNPJ, 11, 1) * 4
	VAR_CNPJ_12 = MID(VAR_CNPJ, 12, 1) * 3
	VAR_CNPJ_13 = VAR_SOMA * 2
	VAR_CNPJ    = VAR_SOMA
	
	VAR_SOMA   = VAR_CNPJ_1 + VAR_CNPJ_2 + VAR_CNPJ_3 + VAR_CNPJ_4 + VAR_CNPJ_5 + VAR_CNPJ_6 + VAR_CNPJ_7 + VAR_CNPJ_8 + VAR_CNPJ_9 + VAR_CNPJ_10 + VAR_CNPJ_11 + VAR_CNPJ_12 + VAR_CNPJ_13
	VAR_SOMA   = VAR_SOMA MOD 11
	
	IF VAR_SOMA < 2 THEN 
		VAR_SOMA = 0 
	ELSE
		VAR_SOMA = 11 - VAR_SOMA
	END IF
	
	VAR_CNPJ = VAR_CNPJ & VAR_SOMA
	MyDVCNPJ  = VAR_CNPJ
END FUNCTION

FUNCTION CNPJ_Generator()
	VAR_STR_FILIAL = "0001"
	Randomize
	VAR_CNPJ = Int((9-1+1)*Rnd+1)
			
	FOR VAR_COUNT = 2 TO 8
		Randomize
		VAR_CNPJ = VAR_CNPJ & Int((9-0+1)*Rnd+0)
	NEXT
			
	VAR_DV_CNPJ = MyDVCNPJ(VAR_CNPJ & VAR_STR_FILIAL)
			
	IF Len(VAR_DV_CNPJ) < 2 Then 
		VAR_DV_CNPJ = String(2 - Len(VAR_DV_CNPJ), "0") & VAR_DV_CNPJ
	END IF
			
		VAR_CNPJ = VAR_CNPJ & VAR_STR_FILIAL & VAR_DV_CNPJ
		VAR_CNPJ = InputBox("CNPJ", "Siebel 6.3 - Gera CNPJ",VAR_CNPJ)
END FUNCTION

FUNCTION REMOVER_INPUT(VAR_STRING)
	Set FSO = WScript.CreateObject("Scripting.FileSystemObject") 
		If (FSO.FileExists(VAR_STRING)) Then
			FSO.DeleteFile(VAR_STRING)
		End If
	Set FSO = Nothing
END FUNCTION

FUNCTION QUERY_INPUT(VAR_STRING, VAR_TYPE)
	vFile1 = VAR_GLOBAIS("IN")
	vFile2 = VAR_GLOBAIS("OUT")
	Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
		REMOVER_INPUT(vFile1)
		REMOVER_INPUT(vFile2)

		Set OTF = FSO.OpenTextFile(vFile1, 8, True)
			OTF.WriteLine "SET HEADING OFF FEEDBACK OFF ECHO OFF PAGESIZE 0 LINESIZE 530"
			OTF.WriteLine "spool '" & vFile2 & "'"
			OTF.WriteLine VAR_STRING
			OTF.WriteLine "SPOOL OFF"
			OTF.WriteLine "EXIT;"
			OTF.Close
		Set OTF = Nothing
	Set FSO = Nothing
	
	Set objShell = CreateObject("WScript.Shell")
		Set objExecObject = objShell.Exec("cmd /c sqlplus " & VAR_GLOBAIS("SBL_LOG") & "/" & VAR_GLOBAIS("SBL_PAS") & "@" & VAR_GLOBAIS("SIEBEL") & " @" & vFile1)
			varTMP = objExecObject.StdOut.ReadAll
			DEBUG(varTMP) 
		Set objExecObject = Nothing
	Set objShell = Nothing
	
	If VAR_TYPE = 0 Then
		Randomize
		OUT_vFile2 = Int((50-1+1)*Rnd+1)
	Else
		OUT_vFile2 = VAR_TYPE
	End IF
	
	Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
		If (FSO.FileExists(vFile2)) Then
			Set OTF = FSO.OpenTextFile(vFile2, 1)
				For varTMP = 1 to OUT_vFile2
					If OTF.AtEndOfStream <> True Then
						TextLine = OTF.ReadLine
					End If
				Next
				OTF.Close
			Set OTF = Nothing
			
			REMOVER_INPUT(vFile1)
			REMOVER_INPUT(vFile2)
	
			TextLine = RTrim(TextLine)
			QUERY_INPUT = TextLine
		Else
			WScript.Echo "Arquivo " & vFile2 & " Inexistente"
			QUERY_INPUT = ""
		End If
	Set FSO = Nothing
		
	
END FUNCTION

FUNCTION VARIFICAR_ARBOR(X_NUM_IMSI)
	vFile1 = VAR_GLOBAIS("IN")
	vFile2 = VAR_GLOBAIS("OUT")
	Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
		REMOVER_INPUT(vFile1)
		REMOVER_INPUT(vFile2)

		Set OTF = FSO.OpenTextFile(vFile1, 8, True)
			OTF.WriteLine "SET HEADING OFF FEEDBACK OFF ECHO OFF PAGESIZE 0 LINESIZE 530"
			OTF.WriteLine "spool '" & vFile2 & "'"
			OTF.WriteLine "SELECT COUNT(*) FROM external_id_equip_map WHERE external_id = '" & X_NUM_IMSI & "';"
			OTF.WriteLine "SPOOL OFF"
			OTF.WriteLine "EXIT;"
			OTF.Close
		Set OTF = Nothing
	Set FSO = Nothing
	
	Set objShell = CreateObject("WScript.Shell")
		Set objExecObject = objShell.Exec("cmd /c sqlplus arbor/upgrade#84178@" & VAR_GLOBAIS("ARBOR") & " @" & vFile1)
			varTMP = objExecObject.StdOut.ReadAll
			DEBUG(varTMP)
		Set objExecObject = Nothing
	Set objShell = Nothing
	
	Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
		Set OTF = FSO.OpenTextFile(vFile2, 1)
			If OTF.AtEndOfStream <> True Then
				TextLine = OTF.ReadLine
			End If
			OTF.Close
		Set OTF = Nothing
	Set FSO = Nothing
		
	REMOVER_INPUT(vFile1)
	REMOVER_INPUT(vFile2)
	
	TextLine = RTrim(TextLine)
	TextLine = LTrim(TextLine)
	VARIFICAR_ARBOR = TextLine
END FUNCTION

FUNCTION IMEI_SQL()
	TextLine = InputBox("Informe o Modelo do Aparelho:","Siebel 6.3 - IMEI DE APARELHOS", "Nokia 3310")
	TextLine = QUERY_INPUT("SELECT ASSET_NUM FROM SIEBEL.S_ASSET A WHERE TYPE_CD = 'IMEI' AND ASSET_NUM LIKE '%' AND X_MODELO = '" & TextLine & "' AND STATUS_CD LIKE 'Dispon%vel' AND ROWNUM < 50;", 0)
	If TextLine <> "" Then
		TextLine = InputBox("IMEI", "Siebel 6.3 - Recuperar IMEI",TextLine)
	Else
		Msgbox "APARELHO NÃO ENCONTRADO."
	End If
END FUNCTION

FUNCTION SIMCARD_SQL()
	HLR_NUM = InputBox("Informe o HLR:", "Oi Chip")
		If IsNumeric(HLR_NUM) Then
			If Len(HLR_NUM) <> 2 then
				MsgBox("Numero de caractere invalido!")						
			else
				TextLine = "SELECT ASSET_NUM || '|' || X_NUM_IMSI || '|' || X_NUM_HLR || '|' || VERSION || '|' || X_TIPO_DESIG as VSTRING FROM SIEBEL.S_ASSET A WHERE STATUS_CD LIKE 'Dispon%vel' AND TYPE_CD = 'IMSI' AND ROWNUM < 50 AND X_RESERVA_ID IS NULL AND X_TIPO_DESIG <> 'Substituição' AND X_NUM_HLR = '" & HLR_NUM & "'"
				If Right(HLR_NUM, 1) = 9 Then
					TextLine = TextLine & " AND X_NUM_IMSI like '724%'"
					'TextLine = TextLine & " AND VERSION = 'uSIM'"
				End If
				TextLine = TextLine & ";"
				TextLine = QUERY_INPUT(TextLine, 0)
				If TextLine = "" Then
					MsgBox "Retorno Vazio para HLR: " & HLR_NUM
				Else
					TextArra = Split(TextLine, "|")
					DV_ASSET_NUM = MyOIChipDV(TextArra(0))
					TextLine = InputBox("Oi Chip + DV (" & TextArra(3) & ")", "Siebel 6.3 - Recuperar SimCard",TextArra(0) & "-" & DV_ASSET_NUM)
					
					If VAR_GLOBAL_ARBOR <> "" Then
						TextLine = MsgBox("Verificar X_NUM_IMSI: " & TextArra(1) & " no Arbor?", vbYesNo, "Verificar SIMCARD (" & TextArra(0) & ") no Arbor")
						If TextLine = vbYes Then
							TextLine = VARIFICAR_ARBOR(TextArra(1))
							If TextLine > 0 Then
								MsgBox "ERRO(" & TextLine & "): X_NUM_IMSI ja em uso no Arbor."
							Else
								MsgBox "OK(" & TextLine & "): X_NUM_IMSI sem uso no Arbor."
							End IF
						End If
					End IF
				End If
			end If
		else
			MsgBox("Opcao Invalida. Somente numeros.")
		end if
END FUNCTION

FUNCTION VAR_GLOBAIS(VAR_STR_GLOBAL)
	Select Case VAR_STR_GLOBAL
		Case "IN" 
			VAR_GLOBAIS = "Query_DV_INPUT.sql"
		Case "OUT"
			VAR_GLOBAIS = "OUT_DV_INPUT.sql"
		Case "SHOW"
			MsgBox "DESENVOLVIDO: JOSE ARTEIRO (jose.cavalcanti)" & VbCrLf & VbCrLf & _
				   "jose.cavalcanti@accenture.com" & VbCrLf & _
			       "SIEBEL 6.3 AM N3" & VbCrLf & _
			       "Versão: " & VAR_WIZARD_SIEBEL_VERSION
		Case "ARBOR"
			vFile = VAR_GLOBAIS("CONEXAO")
			Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
				Set OTF = FSO.OpenTextFile(vFile, 1)
					For varTMP = 1 to 2
						If OTF.AtEndOfStream <> True Then
							VAR_GLOBAIS = OTF.ReadLine
						End If
					Next
					OTF.Close
				Set OTF = Nothing
			Set FSO = Nothing
		Case "SIEBEL"
			vFile = VAR_GLOBAIS("CONEXAO")
			Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
				Set OTF = FSO.OpenTextFile(vFile, 1)
					For varTMP = 1 to 1
						If OTF.AtEndOfStream <> True Then
							VAR_GLOBAIS = OTF.ReadLine
						End If
					Next
					OTF.Close
				Set OTF = Nothing
			Set FSO = Nothing
		Case "SBL_LOG"
			vFile = VAR_GLOBAIS("CONEXAO")
			Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
				Set OTF = FSO.OpenTextFile(vFile, 1)
					For varTMP = 1 to 4
						If OTF.AtEndOfStream <> True Then
							VAR_GLOBAIS = OTF.ReadLine
						End If
					Next
					OTF.Close
				Set OTF = Nothing
			Set FSO = Nothing
		Case "SBL_PAS"
			vFile = VAR_GLOBAIS("CONEXAO")
			Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
				Set OTF = FSO.OpenTextFile(vFile, 1)
					For varTMP = 1 to 5
						If OTF.AtEndOfStream <> True Then
							VAR_GLOBAIS = OTF.ReadLine
						End If
					Next
					OTF.Close
				Set OTF = Nothing
			Set FSO = Nothing
		Case "CONEXAO"
			VAR_GLOBAIS = "Conexao_DV_INPUT.tmp"
		Case Else
			MsgBox("Opcao Invalida.")
	End Select
END FUNCTION

FUNCTION DEBUG(vParam)
	vFile = VAR_GLOBAIS("CONEXAO")
	Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
		Set OTF = FSO.OpenTextFile(vFile, 1)
			For varTMP = 1 to 3
				If OTF.AtEndOfStream <> True Then
					vFile = OTF.ReadLine
				End If
			Next
			OTF.Close
		Set OTF = Nothing
	Set FSO = Nothing
	If vFile = "DEBUG|TRUE" Then
		MsgBox vParam
	End If
END FUNCTION

FUNCTION CONEXAO()

	VAR_ARRAY_MENU = LoginSenha("ALL", "")
	VAR_ARRAY_MENU = Split(VAR_ARRAY_MENU, ";")
	VAR_COUNT_MENU = 0
	
	VAR_STR_MENU = 	"Selecione Conexao Banco de Dados:" & VbCrLf & VbCrLf
	
	For Each ConnectMenu In VAR_ARRAY_MENU
		VAR_COUNT_MENU = VAR_COUNT_MENU + 1
		If VAR_COUNT_MENU < 10 Then
			VAR_STR_MENU = VAR_STR_MENU & "0" & VAR_COUNT_MENU & ". " & ConnectMenu & VbCrLf
		Else
			VAR_STR_MENU = VAR_STR_MENU & VAR_COUNT_MENU & ". " & ConnectMenu & VbCrLf
		End If
	Next
	
	VAR_STR_MENU = VAR_STR_MENU & ""  & VbCrLf & VAR_COUNT_MENU + 1 & ". Sair"
	 
	VAR_STR_MENU=InputBox(VAR_STR_MENU,"Siebel 6.3 - Oi Menu",1)
			
	If IsNumeric(VAR_STR_MENU) And VAR_STR_MENU <> "" Then	
				
		If CInt(VAR_STR_MENU) <= VAR_COUNT_MENU Then
			VAR_ARRAY_MENU2 = Split(VAR_ARRAY_MENU(VAR_STR_MENU - 1), "-")
			VAR_GLOBAL_SIEBEL = Trim(VAR_ARRAY_MENU2(0))
			
		ElseIf CInt(VAR_STR_MENU) = VAR_COUNT_MENU + 1 Then
			CONEXAO = "7"
			
		Else
			MsgBox("Opcao Invalida.")
			CONEXAO = "7"
		End If
		
		If Len(VAR_GLOBAL_SIEBEL) > 0 Then
			VAR_GLOBAL_ARBOR  = LoginSenha(VAR_GLOBAL_SIEBEL, "VAR_GLOBAL_ARBOR")
			VAR_GLOB_SBl_LOG  = LoginSenha(VAR_GLOBAL_SIEBEL, "VAR_GLOB_SBl_LOG")
			VAR_GLOB_SBl_PAS  = LoginSenha(VAR_GLOBAL_SIEBEL, "VAR_GLOB_SBl_PAS")
			
			vFile1 = VAR_GLOBAIS("CONEXAO")
			REMOVER_INPUT(vFile1)
			
			Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
				Set OTF = FSO.OpenTextFile(vFile1, 8, True)
					OTF.WriteLine VAR_GLOBAL_SIEBEL
					OTF.WriteLine VAR_GLOBAL_ARBOR
					OTF.WriteLine "DEBUG|" & VAR_GLOBAL_SIEBEL_DEBUG
					OTF.WriteLine VAR_GLOB_SBl_LOG
					OTF.WriteLine VAR_GLOB_SBl_PAS
					OTF.Close
				Set OTF = Nothing
			Set FSO = Nothing
			DEBUG(VAR_GLOBAL_SIEBEL)
		End If
	Else
		MsgBox("Opcao Invalida. Somente numeros.")
		CONEXAO = "7"
	End If	
END FUNCTION

FUNCTION Desb_BCV()
	VAR_STR_MENU = 	"Função que limpa o campo SIEBEL.S_EMPLOYEE.nick_name" & VbCrLf & "" & VbCrLf & "Informe o Login:"
	VAR_STR_MENU=InputBox(VAR_STR_MENU,"Siebel 6.3 - Desb. Login BCV Client","TR")
	
	VAR_STR_QUERY = QUERY_INPUT("SELECT ROWID || ';' || NICK_NAME FROM SIEBEL.S_EMPLOYEE WHERE login LIKE '" & VAR_STR_MENU & "';", 1)				
	
	If VAR_STR_QUERY <> "" Then
	
		VAR_STR_QUERY = Split(VAR_STR_QUERY, ";")
		
		If Len(VAR_STR_QUERY(1)) > 0 Then
			TextLine = QUERY_INPUT("UPDATE SIEBEL.S_EMPLOYEE SET nick_name = NULL, upg_comp_id = NULL WHERE login LIKE '" & VAR_STR_MENU & "';", 1)
			TextLine = QUERY_INPUT("UPDATE SIEBEL.S_LST_OF_VAL SET active_flg = 'Y' WHERE type = 'PCS_LOGIN_CONT_STC' AND val LIKE '" & VAR_STR_MENU & "';", 1)
			TextLine = QUERY_INPUT("UPDATE SIEBEL.S_LST_OF_VAL SET HIGH = 'SUPRIME' WHERE TYPE = 'PCS_MONITORA_ACESSO_DEFAULT'AND VAL = 'SRF';", 1)
			TextLine = QUERY_INPUT("ALTER USER TR158689 QUOTA UNLIMITED ON TBS_S_DATA;", 1)
			MsgBox("Usuário Liberado.")
		Else
			MsgBox("Login Já se Encontra Desbloqueado.")
		End If
	Else
		MsgBox("Login não Encontrado.")
	End IF
END FUNCTION

'-------------- VOID MAIN --------------------

vFile1 = VAR_GLOBAIS("CONEXAO")
REMOVER_INPUT(vFile1)

vFile1 = VAR_GLOBAIS("IN")
REMOVER_INPUT(vFile1)

vFile1 = VAR_GLOBAIS("OUT")
REMOVER_INPUT(vFile1)

VAR_STOP = CONEXAO()

If VAR_STOP <> "7" Then
	Do
		VAR_STR_MENU = 	"Selecione uma opcao:" & VbCrLf &_
						"" & VbCrLf &_
						"01. Gerador DV - Oi Chip" & VbCrLf &_
						"02. Gerador CPF" & VbCrLf &_
						"03. Gerador CNPJ" & VbCrLf &_
						"04. Recuperar IMEI" & VbCrLf &_
						"05. Recuperar SimCard" & VbCrLf &_
						"06. Analise de Credito" & VbCrLf &_
						"07. Viabilidade Velox" & VbCrLf &_
						"08. XML Notificacao Status Legado" & VbCrLf &_
						"09. Desb. Login BCV Client" & VbCrLf &_
						"" & VbCrLf &_
						"10. Mudar Conexao Banco de Dados" & VbCrLf &_
						"11. Sobre" & VbCrLf &_
						"12. Sair"
	 
		VAR_STR_MENU=InputBox(VAR_STR_MENU,"Siebel 6.3 - Oi Menu",1)
		
		If IsNumeric(VAR_STR_MENU) Then
			Select Case VAR_STR_MENU
				Case 1 
					DV_Generator()
				Case 2
					CPF_Generator()
				Case 3
					CNPJ_Generator()
				Case 4
					IMEI_SQL()
				Case 5
					SIMCARD_SQL()
				Case 6
					Analise_Credito()
				Case 7
					Viabilidade_Velox()
				Case 8
					XML_Legado()
				Case 9
					Desb_BCV()
				Case 10
					CONEXAO()
				Case 11
					VAR_GLOBAIS("SHOW")
				Case 12
					VAR_STOP = "7"
				Case Else
					MsgBox("Opcao Invalida.")
			End Select
		Else
			MsgBox("Opcao Invalida. Somente numeros.")
		End If	
		
		If VAR_STOP <> "7" Then
			VAR_STOP = MsgBox("Realizar Nova Pesquisa?", "36", "Wizard Siebel " & VAR_WIZARD_SIEBEL_VERSION)
		End If
		
	Loop until VAR_STOP = "7"
End If

vFile1 = VAR_GLOBAIS("CONEXAO")
REMOVER_INPUT(vFile1)

vFile1 = VAR_GLOBAIS("IN")
REMOVER_INPUT(vFile1)

vFile1 = VAR_GLOBAIS("OUT")
REMOVER_INPUT(vFile1)