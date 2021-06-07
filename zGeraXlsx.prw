#include 'Totvs.ch'
#include 'Protheus.ch'

/*/{Protheus.doc} zGeraXlsx
    Fun��o para consulta gen�rica - cria tela de exibi��o com pesquisa. Permite copiar c�lulas e exportar os dados para arquivo .xml
	@author Maicon Macedo
	@since 16 nov. 2020
	@version 1.0
/*/

User Function zGeraXlsx(cConsulta, oSay)
	Local oXlsx			:= FWMsExcelXlsx():New()
	Local lOk			:= .F.
	Local cArq			:= ""
	Local cDirTmp		:= GetTempPath()
	Local nAtual		:= 0
	Local aCampos		:= {}
	Local lDirOk		:= .F.
	DEFAULT cConsulta	:= "SQLEXEC_"

	oSay:SetText("Montando o arquivo...")
	
	dbSelectArea(cQryNm)
	(cQryNm)->(dbGoTop())
	
	oXlsx:SetFontSize(12)
	oXlsx:SetFont("Calibri")
	
	oXlsx:AddWorkSheet(cConsulta)
	oXlsx:AddTable(cConsulta,cConsulta)
	
	FOR nAtual := 1 To Len(aHeader)
		cCampoAtu := aHeader[nAtual][1]
		
		oXlsx:AddColumn(cConsulta,cConsulta,cCampoAtu,1,1)
	NEXT
	
	While (cQryNm)->(!EoF())
		aCampos := {}
		
		For nAtual := 1 To Len(aHeader)
			cCampoAtu := aHeader[nAtual][2]
			Aadd( aCampos, & ( (cQryNm)->(cCampoAtu)  ) ) 
		Next
		
		oXlsx:AddRow(cConsulta,cConsulta,aCampos)
		
		lOk := .T.
		
		(cQryNm)->(DbSkip())
	
	EndDo
	
	oXlsx:Activate()
	
	cArq := cConsulta + "_" + CriaTrab(NIL, .F.) + ".xml"
	
	oXlsx:GetXMLFile(cArq)

	oSay:SetText("Exportando o arquivo na pasta...")
	
		If ExistDir(cDirTmp,nil, .F.) == .F.
				If MsgYesNo("Diretorio - "+cDirTmp+" - nao encontrado, deseja cria-lo?" ) 
					If  MakeDir(cDirTmp) <> 0
						MsgInfo("Falha ao criar diret�rio " + cDirTmp + " ! Erro: " + cValToChar( FError() )  , "Diret�rio")
					Else
						MsgInfo("Diret�rio " + cDirTmp + " criado com sucesso!" , "Diret�rio")
						lDirOk = .T.
					EndIf
				Else
					MsgInfo("O diret�rio " + cDirTmp + " n�o foi criado!" , "Diret�rio")
				EndIf
		Else
			lDirOk = .T.
		EndIF

	If lDirOk
		If __CopyFile(cArq,cDirTmp + cArq)
			If lOk
				oXlsxApp := MSExcel():New()
				oXlsxApp:WorkBooks:Open(cDirTmp + cArq)
				oXlsxApp:SetVisible(.T.)
				oXlsxApp:Destroy()
				
			MsgInfo("<h2><font color='#0000FF'>O arquivo Excel foi gerado no dirt�rio: " + cDirTmp + cArq + "</font></h2>","Gera Excel")
			EndIf
		Else
			MsgAlert("Erro ao copiar o arquivo!")
		EndIf
	Else
		MsgAlert("Erro ao copiar o arquivo! O Diret�rio n�o existe.")
	EndIf


Return Nil
