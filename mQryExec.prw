#include 'protheus.ch'
#include 'parmtype.ch'
#include 'TopConn.CH'

/*/{Protheus.doc} mQryExec
	Função que permite colar uma sentença SQL e executa no fonte U_zExM5TR, o qual irá 
		executar a montagem da sentença SQL em uma tela de tabulação de dados
	@author Maicon Macedo
	@since 27/04/2020
	@version 1.0
/*/

User Function mQryExec()
	Local oDlg
	Local oGrp
	Local oMGet
	Local cMGet := "Cole aqui"
	Local oBtn
/*/
  DEFINE MSDIALOG oDlg TITLE "Executar sentenças SQL" FROM 000, 000  TO 500, 500 COLORS 0, 16777215 PIXEL
    @ 003, 006 GROUP oGrp TO 247, 246 PROMPT "  Cole suas sentenças SQL aqui e clique em 'Executar'  " OF oDlg COLOR 0, 16777215 PIXEL
    @ 012, 010 GET oMGet VAR cMGet OF oDlg MULTILINE SIZE 232, 210 COLORS 0, 16777215 HSCROLL PIXEL
    @ 225, 090 BUTTON oBtn PROMPT ":: Executar ::" SIZE 068, 016 OF oDlg ACTION( Processa({||U_zExM5TR(cMGet)},,"Processando consulta ...") ) PIXEL

  ACTIVATE MSDIALOG oDlg CENTERED
/*/
	/* Construtor MsDialog - https://tdn.totvs.com.br/display/tec/Construtor+MsDialog%3ANew*/
  	oDlg  := MSDialog():New(  000,000,500,500,"Executar sentenças SQL",/*6*/,/*7*/,.F.,/*9*/,CLR_BLACK,CLR_WHITE,/*12*/,/*oWnd*/,.T. )
	/* Construtor TGroup - https://tdn.totvs.com.br/display/tec/Construtor+TGroup%3ANew */
	oGrp  := TGroup():New(    003,006,247,246,"  Cole suas sentenças SQL aqui e clique em 'Executar'  ",oDlg,CLR_BLACK,CLR_WHITE,.T.,.F. )
	/* Construtor TMultiGet - https://tdn.totvs.com.br/display/tec/TMultiGet%3ANew */
	oMGet := TMultiGet():New( 012,010,{|u| If(PCount()>0,cMGet:=u,cMGet)},oGrp,232,210,/*oFont*/,/*8*/,/*9*/,/*10*/,/*11*/,.T.,"",/*14*/,/*bWhen*/,.F.,.F.,.F.,/*bValid*/,/*20*/,.F.,/*lNoBorder*/,.T.,/*cLabelText*/,/*nLabelPos*/,/*oLabelFont*/,/*nLabelColor*/  )
	/* Construtor TButton - https://tdn.totvs.com.br/display/tec/Construtor+TButton%3ANew */
	oBtn  := TButton():New(   225,090,":: Executar ::",oGrp,/*bAction*/,068,016,/*8*/,/*oFont*/,/*10*/,.T.,/*12*/,"",/*14*/,/*bWhen*/,/*16*/,.F. )

	oBtn:bAction := {||U_zExM5TR(cMGet)}

	/* TDialog:Activate - https://tdn.totvs.com.br/display/tec/Activate */
	oDlg:Activate(,,,.T.) //Inicia a exibição da Tela com esta centralizada
Return
