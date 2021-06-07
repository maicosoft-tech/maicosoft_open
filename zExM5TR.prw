#include 'protheus.ch'
#include 'totvs.ch'
#include 'parmtype.ch'
#include 'TopConn.CH'

#Define STR_PULA Chr(13)+ Chr(10)
#Define STR_NAME 'SQLEXEC'

/*/{Protheus.doc} zExM5TR
	Função que executa a montagem da sentença SQL em uma tela de tabulação de dados.
	Permite exportar os dados em .xml compatível com Excel
	@author Maicon Macedo
	@since 27/04/2020
	@version 1.0
/*/
USER Function zExM5TR(cSentenca)
	Local aArea			:= GetArea()
	Local oAreaDados
	//Botões
	Local nTamBtn		:= 60
	Local oBtExcel
	Local oBtFecha
	Local oBtEMail
	//Default
	DEFAULT cSentenca	:= ""
	//Private
	Private cConsSQL	:= cSentenca
	Private aEstrutr	:= {} // Estrutura da Grid - baseada na consulta QRY_TMP
	Private cQryNm		:= ""
	Private cDirArq 	:= ""
	//MsDialog - por inteiro
	Private oDltTela
	Private nTelaLarg	:= 1000
	Private nTelaAltu	:= 0600
	//Area dos Dados
	Private oMSNovo	
	Private aHeader	:= {}
	Private aColunas	:= {}
	
	//Chama a montagem da Estrutura do MSDialog na Area de Dados
	FWMsgRun(, {|oSay| MnTEstr(oSay)},"Montando a estrutura da Consulta", "Iniciando...")
	
	DEFINE MSDIALOG oDlgTela TITLE "Consulta de Dados - "+STR_NAME FROM 000, 000 TO nTelaAltu, nTelaLarg PIXEL STYLE DS_MODALFRAME
		@ 003,003 GROUP oAreaDados TO (nTelaAltu/2)-3, (nTelaLarg/2)-3 PROMPT "  Resultado da sentença SQL:  " OF oDlgTela COLOR 0, 16777215 PIXEL
			oMSNovo := MsNewGetDados():New(	010, 009,;
											(nTelaAltu/2)-30,;
											(nTelaLarg/2)-6,;
											GD_INSERT+GD_DELETE+GD_UPDATE,;
											"AllwaysTrue()",;
											,"", , , ; 	//cTudoOk , Inici Campos, Alteracao, Congelamento
											999,;
											, , , ; 	// Campo Ok, Super Del, Delete
											oDlgTela,;
											aHeader,;	// Array do Cabeçalho
											aColunas)	// Array das Colunas
			oMSNovo:lActive := .F.
			
		FWMsgRun(, {|oSay| PopDados(oSay)},"Executando consulta", "Iniciando...")
		@ (nTelaAltu/2)-22, (nTelaLarg/2)-((nTamBtn*1)+06) BUTTON oBtFecha PROMPT "Fechar"		SIZE nTamBtn, 013	OF oDlgTela ACTION (btFecha()) PIXEL	
		@ (nTelaAltu/2)-22, (nTelaLarg/2)-((nTamBtn*2)+12) BUTTON oBtExcel PROMPT "Exportar"	SIZE nTamBtn, 013	OF oDlgTela ACTION (FWMsgRun(, {|oSay| GeraExcel(1,oSay)},"Exportando os dados", "Iniciando...") ) 	PIXEL
		@ (nTelaAltu/2)-22, (nTelaLarg/2)-((nTamBtn*3)+18) BUTTON oBtEMail PROMPT "E-Mail"		SIZE nTamBtn, 013	OF oDlgTela ACTION (FWMsgRun(, {|oSay| zDoMail(oSay)},"Enviando os dados via E-Mail", "Iniciando...") ) 		PIXEL
		@ (nTelaAltu/2)-22, (nTelaLarg/2)-((nTamBtn*4)+24) BUTTON oBtExcel PROMPT "Gera .xlsx"	SIZE nTamBtn, 013	OF oDlgTela ACTION (FWMsgRun(, {|oSay| u_zGeraXlsx(STR_NAME,oSay)},"Exportando os dados", "Iniciando...") ) 	PIXEL
	oMSNovo:oBrowse:SetFocus()
		
	ACTIVATE MSDIALOG oDlgTela CENTERED
	
	RestArea(aArea)
			
Return

/*---------------------------------------------------------------------------------------------------------*
*	MnTEstr - Montagem da tela de dados
*---------------------------------------------------------------------------------------------------------*/
Static Function MnTEstr(oSay)
	Local aAreaX3	:= SX3->(GetArea())
	Local cQuery	:= ""
	Local nAtual	:= 0

	oSay:SetText("Montando estrutura...")

	//Definir como zero os valores das Colunas e do Cabeçalho
	aHeader 	:= {}
	aColunas	:= {}
	
	cQuery := cConsSQL
	
	If Select ("QRY_TMP") <> 0
		DbSelectArea("QRY_TMP")
		DbCloseArea()
	EndIf
	
	TCQuery cQuery New Alias "QRY_TMP"
	
	aEstrutr := QRY_TMP->(DbStruct())
	
	QRY_TMP->(DbCloseArea())
	
	DbSelectArea("SX3")
	SX3->(DbSetOrder(2))
	SX3->(DbGoTop())
	
	For nAtual := 1 To Len(aEstrutr)
		cCampoAtu := aEstrutr[nAtual][1]
		
		If SX3->(DbSeek(cCampoAtu))
			aAdd(aHeader, { X3Titulo(),;
												cCampoAtu,;
												PesqPict(SX3->X3_ARQUIVO, cCampoAtu),;
												SX3->X3_TAMANHO,;
												SX3->X3_DECIMAL,;
												".F.",;
												".F.",;
												SX3->X3_TIPO,;
												"",;
												"" 	}	)
		Else
			aAdd(aHeader, { Capital(StrTran(cCampoAtu, '_', ' ')),;
												cCampoAtu,;
												"",;
												aEstrutr[nAtual][3],;
												aEstrutr[nAtual][4],;
												".F.",;
												".F.",;
												aEstrutr[nAtual][2],;
												"",;
												"" 	}	)
		EndIf
	
	Next
	
	RestArea(aAreaX3)

Return

/*---------------------------------------------------------------------------------------------------------*
*	Função PodDados - para popular os dados da grid da tela												   *
*----------------------------------------------------------------------------------------------------------*/
Static Function PopDados(oSay)
	Local cQuery	:= ""
	Local nAtual	:= 0
	Local nCampAux	:= 1
	Private cQrySQL	:= GetNextAlias()

	oSay:SetText("Buscando os dados na tabela...")
	
	cQryNm := cQrySQL
	
	aColunas	:= {}
	
	cQuery := cConsSQL
	cQuery := ChangeQuery(cQuery)
	
	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cQrySQL,.T.,.T.)
	
	//Tratamento da exibição dos campos de Data e Números
	For nAtual := 1 to Len(aHeader)
		If aHeader[nAtual][8] == "D"
			TcSetField(cQrySQL, aHeader[nAtual][2],'D')
		ElseIf aHeader[nAtual][8] == "N"
			TcSetField(cQrySQL, aHeader[nAtual][2],'N', aHeader[nAtual][4],aHeader[nAtual][5])
		EndIf
	Next
	
	While (cQrySQL)->(!EoF())
		nCampAux := 1
		aAux	:= {}
		
		For nAtual := 1 To Len(aEstrutr)
			cCampoAtu := aEstrutr[nAtual][1]
			
			If aEstrutr[nAtual][2] $ "N;D"
				aAdd(aAux, &( (cQrySQL)->(cCampoAtu) ) )
			Else
				aAdd(aAux, cValToChar( &( (cQrySQL)->(cCampoAtu)  ) ) )
			EndIf		
		Next
		aAdd(aAux, .F. )
		
		aAdd(aColunas, aClone(aAux) )
		
		(cQrySQL)->(DbSkip() )
	
	EndDo
	
	// Caso a sentença não retorne dado nenhum = passar linhas em branco
	If Len(aColunas) == 0
		aAux := {}
		
		For nAtual := 1 To Len(aEstrutr)
			aAdd(aAux, '')
		Next
		
		aAdd(aAux, .F. )
		
		aAdd(aColunas, aClone(aAux) )
		
	EndIf
	
	oMSNovo:SetArray(aColunas)
	oMSNovo:oBrowse:Refresh()
Return

/*---------------------------------------------------------------------------------------------------------*
*	Função GeraExcel - mecanismo utilizado para exportar os dados em .xml compatível com Excel
*	@param nOper, Operação 1 = Exportar arquivo e abrir - 2 = Exportar arquivo para e-mail
*----------------------------------------------------------------------------------------------------------*/
Static Function GeraExcel(nOper,oSay)
	Local oExcel	:= FWMSEXCEL():New()
	Local lOk		:= .F.
	Local cArq		:= ""
	Local cDirTmp	:= "C:\temp\"
	Local nAtual	:= 0
	Local aCampos	:= {}
	Local lDirOk	:= .F.
	Default nOper	:= 1
	
	cDirArq	:= "\anexos\"

	oSay:SetText("Montando o arquivo...")
	
	dbSelectArea(cQryNm)
	(cQryNm)->(dbGoTop())
	
	//Atribuindo fortamação ao Excel
	//https://html-color.codes
	oExcel:SetTitleSizeFont(12)
	oExcel:SetTitleBold(.T.)
	oExcel:SetTitleFrColor("#2200fc")
	oExcel:SetTitleBgColor("#adadad")
	oExcel:SetFontSize(11)
	oExcel:SetFont("Calibri")
	
	oExcel:AddWorkSheet(STR_NAME)
	oExcel:AddTable(STR_NAME,STR_NAME)
	
	FOR nAtual := 1 To Len(aHeader)
		cCampoAtu := aHeader[nAtual][1]
		
		oExcel:AddColumn(STR_NAME,STR_NAME,cCampoAtu,1,1)
	NEXT
	
	While (cQryNm)->(!EoF())
		aCampos := {}
		
		For nAtual := 1 To Len(aHeader)
			cCampoAtu := aHeader[nAtual][2]
			Aadd( aCampos, & ( (cQryNm)->(cCampoAtu)  ) ) 
		Next
		
		oExcel:AddRow(STR_NAME,STR_NAME,aCampos)
		
		lOk := .T.
		
		(cQryNm)->(DbSkip())
	
	EndDo
	
	oExcel:Activate()
	
	cArq := STR_NAME + "_" + CriaTrab(NIL, .F.) + ".xml"
	
	oExcel:GetXMLFile(cArq)
	
	oSay:SetText("Exportando o arquivo na pasta...")

	If nOper == 1 //Exportar o arquivo e abrir o Excel
	//Verificar a existência do diretório indicado na cDirTmp
		If ExistDir(cDirTmp,nil, .F.) == .F.
				If MsgYesNo("Diretorio - "+cDirTmp+" - nao encontrado, deseja cria-lo?" ) 
					If  MakeDir(cDirTmp) <> 0
						MsgInfo("Falha ao criar diretório " + cDirTmp + " ! Erro: " + cValToChar( FError() )  , "Diretório")
					Else
						MsgInfo("Diretório " + cDirTmp + " criado com sucesso!" , "Diretório")
						lDirOk = .T.
					EndIf
				Else
					MsgInfo("O diretório " + cDirTmp + " não foi criado!" , "Diretório")
				EndIf
		Else
			lDirOk = .T.
		EndIF
	Else //nOper == 2 //Exportar o arquivo e enviar um e-mail
		//Verificar a existência do diretório indicado na cDirArq
		If ExistDir(cDirArq,nil, .F.) == .F.
				If MsgYesNo("Diretorio - "+cDirArq+" - nao encontrado, deseja cria-lo?" ) 
					If  MakeDir(cDirArq) <> 0
						MsgInfo("Falha ao criar diretório " + cDirArq + " ! Erro: " + cValToChar( FError() )  , "Diretório")
					Else
						MsgInfo("Diretório " + cDirArq + " criado com sucesso!" , "Diretório")
						lDirOk = .T.
					EndIf
				Else
					MsgInfo("O diretório " + cDirArq + " não foi criado!" , "Diretório")
				EndIf
		Else
			lDirOk = .T.
		EndIF
	EndIf

	If lDirOk
		If nOper == 1 //Exportar e abrir o arquivo - Diretorio cDirTmp
			If __CopyFile(cArq,cDirTmp + cArq)
				If lOk
					oExcelApp := MSExcel():New()
					oExcelApp:WorkBooks:Open(cDirTmp + cArq)
					oExcelApp:SetVisible(.T.)
					oExcelApp:Destroy()
					
					MsgInfo("<h2><font color='#0000FF'>O arquivo Excel foi gerado no dirtério: " + cDirTmp + cArq + "</font></h2>","Gera Excel")
				Else
					MsgAlert("Erro ao abrir o arquivo!")
				EndIf
			Else
				MsgAlert("Erro ao gerar o arquivo!")
			EndIf
		Else //ElseIf nOper == 2 //Exportar e enviar via e-mail  - Diretorio cDirArq
			If __CopyFile(cArq,cDirArq + cArq)
				cDirArq += cArq
				MsgInfo("<h2><font color='#0000FF'>O arquivo gerado no dirtério: " + cDirArq + "</font></h2>","E-Mail")
			Else
				MsgAlert("Erro ao gerar o arquivo!")
			EndIf
		EndIf
	Else
		MsgAlert("Erro ao gerar o arquivo! O Diretório não existe.")
	EndIf

Return Nil

/*---------------------------------------------------------------------------------------------------------*
*	Função GeraExcel - mecanismo utilizado para exportar os dados em .xml compatível com Excel
*----------------------------------------------------------------------------------------------------------*/
Static Function zDoMail(oSay)
    Local aArea        := GetArea()
    Local lRet         := .T.
    Local oMsg         := Nil //Objeto da Classe TMailMessage
    Local oSrv         := Nil //Objeto da Classe tMailManager
    Local nRet         := 0
    /* Variáveis para receber os parâmetros SX6 */
    Local cFrom        := Alltrim(GetMV("MV_RELACNT"))
    Local cPass        := Alltrim(GetMV("MV_RELPSW"))
    Local cSrvFull     := Alltrim(GetMV("MV_RELSERV"))
    Local nTimeOut     := GetMV("MV_RELTIME")
    /* Variáveis para conexão com a classe tMailManager */
    Local cUser        := SubStr(cFrom, 1, At('@', cFrom)-1)
    Local cServer      := Iif(':' $ cSrvFull, SubStr(cSrvFull, 1, At(':', cSrvFull)-1), cSrvFull)
    Local nPort        := Iif(':' $ cSrvFull, Val(SubStr(cSrvFull, At(':', cSrvFull)+1, Len(cSrvFull))), 587)
    Local lUsaTLS      := .T.
    /* Variável para apresentar os Logs */
    Local cLog         := ""
    /* Variáveis para compor o e-mail*/
    Local cAssunto     := ""
    Local cCorpo       := "" //Corpo do e-Mail (com suporte à html)
    Local cHtml        := ""
    Local cPara        := ""
    Local cAnexo       := ""
	Local lRmail	   := .T.

	oSay:SetText("Preparando...")

	FWMsgRun(, {|oSay| GeraExcel(2,oSay)},"Obtendo o arquivo", "Iniciando...")

	cAnexo := cDirArq 

	While lRmail
		cPara := FWInputBox("Informe o e-mail", "")
		If IsEmail(cPara)
			lRmail := .F.
		Else
			MsgAlert("Por favor, insira um e-mail válido!")
		EndIf
	EndDo
	
    cAssunto := "Relatório :: "+STR_NAME+" :: {"+dToC(Date())+"}"

    cHtml    += '<html xmlns="http://www.w3.org/1999/xhtml">' 
    cHtml    += '<head><title></title><meta charset="iso-8859-1"></head>'
    cHtml    += '<body>'
    cHtml    += '<p align=left><font face="Lucida Sans Unicode" size="4">Relatório '+STR_NAME+' enviado em '+dToC(Date())+'.</font></p>' 
    cHtml    += '</body></html>'

    cCorpo   := cHtml

	oSay:SetText("Preparando...")

    If lRet
        //Cria a nova mensagem
        oMsg := TMailMessage():New()
        oMsg:Clear()
 
        //Define os atributos da mensagem
        oMsg:cFrom    := cFrom
        oMsg:cTo      := cPara
        oMsg:cSubject := cAssunto
        oMsg:cBody    := cCorpo
 

        If File(cAnexo)
            //Anexa o arquivo na mensagem de e-Mail
            nRet := oMsg:AttachFile(cAnexo)
            If nRet < 0
                cLog += "01 - Nao foi possivel anexar o arquivo '"+cAnexo+"'!" + CRLF
            EndIf
        //Senao, acrescenta no log
        Else
            cLog += "02 - Arquivo '"+cAnexo+"' nao encontrado!" + CRLF
        EndIf
 
    //Cria servidor para disparo do e-Mail
        oSrv := tMailManager():New()
 
    //Define se irá utilizar o TLS
        If lUsaTLS
            oSrv:SetUseTLS(.T.)
        EndIf
 
    //Inicializa conexão com o servidor de e-mail
        nRet := oSrv:Init("", cServer, cUser, cPass, 0, nPort)
        If nRet != 0
            cLog += "03 - Nao foi possivel inicializar o servidor SMTP: " + oSrv:GetErrorString(nRet) + CRLF
            lRet := .F.
        EndIf
 
        If lRet
        //Define o time out
            nRet := oSrv:SetSMTPTimeout(nTimeOut)
            If nRet != 0
                cLog += "04 - Nao foi possivel definir o TimeOut '"+cValToChar(nTimeOut)+"'" + CRLF
            EndIf
 
        //Conecta no servidor
            nRet := oSrv:SMTPConnect()
            If nRet <> 0
                cLog += "05 - Nao foi possivel conectar no servidor SMTP: " + oSrv:GetErrorString(nRet) + CRLF
                lRet := .F.
            EndIf
 
            If lRet
            //Realiza a autenticação do usuário e senha
                nRet := oSrv:SmtpAuth(cFrom, cPass)
                If nRet <> 0
                    cLog += "06 - Nao foi possivel autenticar no servidor SMTP: " + oSrv:GetErrorString(nRet) + CRLF
                    lRet := .F.
                EndIf
 
                If lRet
                //Envia a mensagem
                    nRet := oMsg:Send(oSrv)
                    If nRet <> 0
                        cLog += "07 - Nao foi possivel enviar a mensagem: " + oSrv:GetErrorString(nRet) + CRLF
                        lRet := .F.
					Else 
						oSay:SetText("Enviando...")
                    EndIf
                EndIf
 
            //Disconecta do servidor
                nRet := oSrv:SMTPDisconnect()
                If nRet <> 0
                    cLog += "08 - Nao foi possivel disconectar do servidor SMTP: " + oSrv:GetErrorString(nRet) + CRLF
                EndIf
            EndIf
        EndIf
    EndIf
 
	//Resultado 
		/* Primeiro = Caso houver algum erro será apresentado o log correspondente, desde que a rotina não esteja sendo executada em segundo plano */
	If !IsBlind() 
		If !Empty(cLog)
			cLog := "+======================= zDoMail =======================+" + CRLF + CRLF +;
					"zDoMail      - " + dToC(Date()) + " " + Time() + CRLF + ;
					"Rotina       - " + FunName() + CRLF + ;
					"Destinatário - " + cPara + CRLF + ;
					"Assunto      - " + cAssunto + CRLF + CRLF +;
					"Qual o problema ocorrido: "+ CRLF+;
					cLog + CRLF +;
					"+=======================================================+"

			Aviso("Ops! Ocorreu algum problema", cLog, {"Ok"}, 2)
		Else
			cLog := "+======================= zDoMail =======================+" + CRLF + CRLF +;
					"zDoMail      - " + dToC(Date()) + " " + Time() + CRLF + ;
					"Rotina       - " + FunName() + CRLF + ;
					"Destinatário - " + cPara + CRLF + ;
					"Assunto      - " + cAssunto + CRLF + CRLF +;
					"« Mensagem enviada com sucesso! »" + CRLF+CRLF +;
					"+=======================================================+"

			Aviso("Mensagem enviada com sucesso", cLog, {"Ok"}, 2)
		EndIf
	EndIf
 
    RestArea(aArea)
Return lRet

/*---------------------------------------------------------------------------------------------------------*
*	Função btFecha - Encerra a sessão e fecha a tela
*---------------------------------------------------------------------------------------------------------*/
Static Function btFecha()

	(cQryNm)->(DbCloseArea())
	
	oDlgTela:End()
Return
