#INCLUDE "PROTHEUS.CH"
/*/{Protheus.doc} WSBBC
Função que busca moedas do banco central usando json
@type function
@version 1.0.2
@author Rodrigo Araujo
@since 5/20/2023
/*/
User Function WSBBC()
	Local cModulo 	:= 'SIGAFAT'
	MsApp():New(cModulo)
	oApp:cInternet := NIL
	oApp:CreateEnv()
	PtSetTheme("OCEAN")
	oApp:cStartProg    	:= 'U_WSBBCM'
	oApp:lMessageBar	:= .T.
	oApp:cModDesc		:= cModulo
	__lInternet 		:= .T.
	lMsFinalAuto 		:= .F.
	oApp:lMessageBar	:= .T.
	oApp:Activate()
Return

User Function WSBBCM()
	Local oBmpSim    := LoadBitmap(GetResources(), "LBOK")
	Local oBmpNao    := LoadBitmap(GetResources(), "LBNO")
	Local aBotoes    := {}
	Local oPesquisar := Nil
	Local oProcurar  := Nil
	Local oImportar  := Nil
	Local oLayer     := Nil
	Local oTela1     := Nil
	Local oTela2     := Nil
	Local oPanel0    := Nil
	Local oScroll    := Nil
	Local oGetMoeda  := Nil
	Local cGetMoeda  := Space(3)
	Local cButton1   := "QPushButton { background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgb(232, 232, 232), stop:1 rgb(253,253,253)); border: 1px solid #C0C0C0;outline:0; border-radius: 4px; font: normal 12px Arial; padding: 6px;color: #000000;} QPushButton:pressed {background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgb(253,253,253), stop:1 rgb(232, 232, 232));border-style: inset; border-color: #000000; color: #000000;}"
	Local cButton2   := "QPushButton { background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgb(59,170,200), stop:1 rgb(26,141,175)); font-weight: bold; border: 2px solid #096A82;outline:0; border-radius: 4px; font: normal 12px Arial; padding: 6px;color: #ffffff;} QPushButton:pressed {background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgb(26,141,175), stop:1 rgb(59,170,200));border-style: inset; border-color: #35ACCA; color: #ffffff; }"
	Local cButton3   := "QPushButton { background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgb(59,170,200), stop:1 rgb(26,141,175)); font-weight: bold; border: 2px solid #096A82;outline:0; border-radius: 4px; font: normal 12px Arial; padding: 6px;color: #ffffff;} QPushButton:pressed {background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgb(26,141,175), stop:1 rgb(59,170,200));border-style: inset; border-color: #35ACCA; color: #ffffff; }"
	
	Private cMvMoeda := Alltrim(SUPERGETMV( "MV_MOEDACM", .F., ""))
	Private oMoeSis  := Nil
	Private cMoeSis  := "" //Moedas do sistema
	Private dMoeSis  := ""
	Private aCurrency:= {{2,"USD"}, {3,"EUR"}, {4,"DKK"}}
	Private cPasta   := GetTempPath()
	Private cArquivo := ""
	Private oTipo    := Nil
	Private cTipo    := ""
	Private oMoeda   := Nil
	Private aMoedas  := {}
	Private oData    := Nil
	Private dData    := CTOD("//")

	Static oDlg

	AbreParam()

	DEFINE MSDIALOG oDlg TITLE "Cotação de moeda" FROM 000, 000  TO 600, 500 COLORS 0, 16777215 PIXEL
	dData := dDatabase

	oLayer := FWLayer():New()
	oLayer:Init( oDlg, .F. )

	oLayer:AddLine( "LINE01", 74 )
	oLayer:AddLine( "LINE02", 19 )
	
	oLayer:AddCollumn( "BOX_1" ,100,,"LINE01" )
	oLayer:AddCollumn( "BOX_2" ,100,,"LINE02" )
	
	oLayer:AddWindow( "BOX_1" , "PANEL01","Menu"	,100, .F.,,,"LINE01" )
	oLayer:AddWindow( "BOX_2" , "PANEL02","Moedas" 	,100, .F.,,,"LINE02" )

	oTela1 := oLayer:GetWinPanel("BOX_1" ,"PANEL01","LINE01")
	oTela2 := oLayer:GetWinPanel("BOX_2" ,"PANEL02","LINE02")

	oPanel0:= tPanel():New(0,0,"",oTela1,,,,,,oTela1:nWidth,30)
	oPanel0:Align := CONTROL_ALIGN_TOP

	oGetMoeda:= TGet():New(001,001,{|u| if(PCount()>0,cGetMoeda:=u,cGetMoeda)} ,oPanel0,030, 014,"@!",{|| },0,,,.F.,,.T.,,.F.,{|| .T. },.F.,.F.,{||  },.F.,.F.,"","cGetMoeda",,,,.F.,.T.,,"Procurar",1,,0 )
	oGetMoeda:cToolTip   := "Digite a moeda que deseja buscar na lista abaixo"
	oGetMoeda:bHelp := {|| ShowHelpCpo( "cGetMoeda",{ "Digite a moeda que deseja buscar na lista abaixo" }, 2, { "" }, 2) }

	oProcurar	:= tButton():New(008,035,'Procurar' ,oPanel0, {|| Procurar(cGetMoeda)  },50,16,,,,.T.)
	oProcurar:SetCss(cButton2)

	oImportar:= tButton():New(008,090,'Importar' ,oPanel0, {|| MsAguarde({|| IMPORTAR("") },"Moedas BBC","Aguarde...")   },50,16,,,,.T.)
	oImportar:cToolTip   := "O arquivo deverá estar em CSV com 2 colunas apenas (codigo da moeda e descrição) e separadas por ponto-e-virgula."
	oImportar:SetCss(cButton1)

	oScroll:= TScrollBox():New(oPanel0,-10,0,oPanel0:nHeight-20,80,.T.,.F.,.T.)
	oScroll:Align := CONTROL_ALIGN_RIGHT

	oMoeSis:= tSay():New( 000, 000,{|| "Moedas do Sistema<br>" + cMoeSis },oScroll,,,,,,.T.,,,100,300,,,,,,.T.)
	GetMoeda()

	oMoeda:= TCBrowse():New( 0,0, 400, 300, , , , oTela1,,,,,,,,,,,,.F.,,.T.,,,,.T.,.T.)
	oMoeda:AddColumn(TCColumn():New(""              , {|| If(aMoedas[oMoeda:nAt,01],oBmpSim,oBmpNao)     },,,,,,.T.,.F.,,,,.F., ) )
	oMoeda:AddColumn(TCColumn():New("Moeda"		    , {|| aMoedas[oMoeda:nAt,02]},"@!"         ,,,"CENTER", 035,.F.,.F.,,{|| .F. },,.F., ) )
	oMoeda:AddColumn(TCColumn():New("Descrição"		, {|| aMoedas[oMoeda:nAt,03]},"@!"         ,,,"LEFT"  , 100,.F.,.F.,,{|| .F. },,.F., ) )
	oMoeda:AddColumn(TCColumn():New("Valor" 		, {|| aMoedas[oMoeda:nAt,04]},"@E 999.9999",,,"RIGHT" , 040,.F.,.F.,,{|| .F. },,.F., ) )
	oMoeda:SetArray(aMoedas)
	oMoeda:bWhen        :={|| Len(aMoedas) > 0 }
	oMoeda:bLDblClick   :={|| aMoedas[oMoeda:nAt,1] := !aMoedas[oMoeda:nAt,1]}
	oMoeda:bHeaderClick :={|| SelecTudo() }
	oMoeda:lHScroll     := .F.
	oMoeda:Align        := CONTROL_ALIGN_ALLCLIENT
	oMoeda:cToolTip     := "Você pode selecionar as moedas que utiliza no sistema para ser gravado na tabela de moedas SM2"
	oMoeda:SetOrder(3)
	oMoeda:Refresh()

	oData:= TGet():New(001,005,{|u| if(PCount()>0,dData:=u,dData)} ,oTela2,50, 014,"@D",{|| GetMoeda() },0,,,.F.,,.T.,,.F.,{|| .T. },.F.,.F.,{||  },.F.,.F.,"","dData",,,,.F.,.T.,,"Data",1,,0 )
	oData:cToolTip   := "Informe a data que deseja pesquisar a cotação da moeda"
	oData:bHelp := {|| ShowHelpCpo( "dData",{ "Informe a data que deseja pesquisar a cotação da moeda" }, 2, { "" }, 2) }

	oTipo:= TComboBox():New(001,060,{|u|if(PCount()>0,cTipo:=u,cTipo)},{"1=Cotação Venda","2=Cotação Compra"},065,17,oTela2,,,,,,.T.,,,,,,,,,"cTipo","Pesquisar em:",1)
	oTipo:cToolTip   := "Você pode cotar o valor de venda ou de compra"
	oTipo:bHelp := {|| ShowHelpCpo( "cTipo",{ "Você pode cotar o valor de venda ou de compra" }, 2, { "" }, 2) }

	oPesquisar:= tButton():New(008,130,'Buscar valor da moeda no BBC' ,oTela2, {|| MsAguarde({|| Pesquisa() },"Banco Central do Brasil","Aguarde...")  },((oTela2:nWidth/2) - 135),16,,,,.T.)
	oPesquisar:SetCss(cButton3)

	if File(cArquivo)
		MsAguarde({|| IMPORTAR(cArquivo) },"Moedas BBC","Aguarde...")
	Endif

	ACTIVATE MSDIALOG oDlg ON INIT EnchoiceBar(oDlg, {|| MsAguarde({|| ATUALIZAR() },"Aguarde...") }, {|| oDlg:End()},,aBotoes) CENTERED

	SalvaParam()
Return

Static Function GetMoeda()
	Local _cDescr1  := ""
	Local _cDescr2  := ""
	Local _cMoeda   := ""
	Local _aMoedas  := {}
	//Local cMvMoeda := Alltrim(SUPERGETMV( "MV_MOEDACM", .F., ""))
	Local i
	
	DBSELECTAREA( "SM2" )
	SM2->(dbSetOrder(1))
	SM2->(dbGoTop())
	If SM2->(DbSeek(DTOS(dData)))

		For i := 1 To Len(cMvMoeda)
			_cMoeda := Alltrim(Substr(cMvMoeda,i,1))
			aadd(_aMoedas, { DTOC(SM2->M2_DATA), _cMoeda, GETMV("MV_MOEDA" + _cMoeda) , &("SM2->M2_MOEDA" + _cMoeda) })
		Next

		
		For i := 1 To Len(_aMoedas)
			_cDescr1 := "<span style='color: #DC143C;'>" + _aMoedas[i,1] + "</span>"
			_cDescr2 += PADR(_aMoedas[i,3], 15) + " = <b><span style='color: #DC143C;'>" + Alltrim(Transform(_aMoedas[i,4], PesqPict("SM2","M2_MOEDA2") ) ) + "</span></b><br>"
		Next
		
		cMoeSis:= _cDescr1 + "<br>" + _cDescr2
		oMoeSis:Refresh()
	Endif
Return

Static Function Procurar(cMoeda)
	Local i

	For i := 1 To Len(aMoedas)
		If cMoeda == aMoedas[i,2]
			oMoeda:GoPosition(i)
			oMoeda:Refresh()
			Exit
		Endif
	Next

Return

Static Function SelecTudo()
	Local i

	For i := 1 To Len(aMoedas)
		aMoedas[i,1] := !aMoedas[i,1]
	Next
	oMoeda:Refresh()
Return

Static Function Importar(_cArquivo)
	Local _aMoedas := {}
	Local i
	Local nHdl

	If !File(_cArquivo)
		cArquivo := tFileDialog( "Arquivo CSV (*.csv)", "Selecione o Arquivo....",, cPasta, .F.,   )

		If !File(cArquivo)
			MSGSTOP( "Arquivo inválido ou inexistente!", "Erro" )
			Return
		ENDIF
	Else
		cArquivo := _cArquivo
	Endif

	MsProcTxt("Lendo arquivo...")

	nHdl := fOpen(cArquivo)
	If nHdl == -1
		IF FERROR()== 516
			ALERT("Feche o arquivo.")
			Return
		EndIF
	EndIf

	If nHdl == -1
		MsgAlert("O arquivo de nome "+cArquivo+" nao pode ser aberto! Verifique os parametros.","Atencao!")
		Return
	Endif

	FSEEK(nHdl,0,0)
	fClose(nHdl)

	FT_FUse(cArquivo)  //abre o arquivo
	FT_FGOTOP()         //posiciona na primeira linha do arquivo

	_aMoedas:={}
	While !FT_FEOF()
		if !empty(FT_FREADLN())
			MsProcTxt("Importando Moeda...")
			AADD(_aMoedas,Separa(FT_FREADLN(),";",.T.))
		endif

		FT_FSKIP()
	EndDo

	FT_FUse()
	fClose(nHdl)

	aMoedas:={}
	MsProcTxt("Atualizando lista de Moedas...")
	For i := 1 To Len(_aMoedas)
		aAdd(aMoedas,{.F., _aMoedas[i,1], _aMoedas[i,2], 0})
	Next
	if Len(_aMoedas) > 1
		aSort(aMoedas,,, {|x,y| x[2] < y[2]})
		oMoeda:SetArray(aMoedas)
		oMoeda:Refresh()
	Endif
Return

Static Function Pesquisa()
	Local cUrl        := "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoMoedaPeriodoFechamento(codigoMoeda=@codigoMoeda,dataInicialCotacao=@dataInicialCotacao,dataFinalCotacao=@dataFinalCotacao)?"
	Local cParam      := ""
	Local cRetorno    := ""
	Local cData       := ""
	Local nValor      := 0
	Local i           := 0
	Local oJson       := Nil
	local oRestClient := Nil
	Local nConta      := 0
	Local dNovaData   := CTOD("")

	For i := 1 To Len(aMoedas)
		IF aMoedas[i,1]
			nConta ++
		Endif
	Next

	If nConta == 0
		MSGSTOP( "Selecione ao menos 1 moeda para atualizar", "Atenção" )
		Return
	Endif

	dNovaData := dData

	If Empty(dNovaData)
		dNovaData := dDatabase
	Endif

	If dow(dNovaData)== 1 //domingo
		dNovaData -= 2
		MSGINFO( "A data informada é de um Domingo, iremos buscar a moeda referente a data " + DTOC(dNovaData), "Domingo" )
	ElseIf dow(dNovaData)== 7 //sabado
		dNovaData -= 1
		MSGINFO( "A data informada é de um Sábado, iremos buscar a moeda referente a data " + DTOC(dNovaData), "Sábado" )
	EndIf

	For i := 1 To Len(aMoedas)
		IF aMoedas[i,1]
			
			aMoedas[i,1] := .F.
			
			MsProcTxt("Buscando moeda "+ Upper(aMoedas[i,3]) +"...")

			cData  := StrZero(Month(dNovaData),2) + "-" + StrZero(Day(dNovaData),2) + "-" + Alltrim(Str(Year(dNovaData)))
			cParam := "@codigoMoeda='"+aMoedas[i,2]+"'"
			cParam += "&@dataInicialCotacao='"+cData+"'"
			cParam += "&@dataFinalCotacao='"+cData+"'"
			cParam += "&$format=json"
			cParam += "&$select=" + IIf(Left(cTipo,1)=="1","cotacaoVenda","cotacaoCompra")

			oRestClient := FWRest():New(cUrl)
			oRestClient:setPath(cParam)
			if oRestClient:Get()
				cRetorno := oRestClient:GetResult()
				FWJsonDeserialize(cRetorno, @oJson)

				If VALTYPE(oJson) == "O"
					if VALTYPE(oJson:VALUE) == "A"
						if Len(oJson:VALUE) > 0
							if VALTYPE( oJson:value[1]:cotacaovenda ) == "N"
								MsProcTxt("Atualizando moeda "+ Upper(aMoedas[i,3]) +"...")
								nValor := oJson:value[1]:cotacaovenda
								aMoedas[i,4] := nValor
								aMoedas[i,1] := .T.
								oMoeda:GoPosition(i)
							Endif
						Endif
					Endif
				Endif
			endif
		Endif		
		oMoeda:Refresh()
	Next
Return

Static Function ATUALIZAR()
	Local i         := 0
	Local y         := 0
	Local nConta    := 0	
	Local _aMoedas  := {}

	For i := 1 To Len(aMoedas)
		IF aMoedas[i,1]
			nConta ++
		Endif
	Next

	If nConta == 0
		MSGSTOP( "Selecione ao menos 1 moeda para atualizar", "Atenção" )
		Return
	Endif

	If !MsgYesNo("Deseja atualizar as moedas na tabela SM2 ?")
		Return
	Endif

	For y := 1 To Len(aCurrency)
		For i := 1 To Len(aMoedas)
			If aMoedas[i,1]
				If aCurrency[y,2] == aMoedas[i,2]	
					aadd(_aMoedas,{ aCurrency[y,1],aCurrency[y,2],aMoedas[i,4] })
					aMoedas[i,1] := .F.
					oMoeda:Refresh()
				Endif
			Endif
		Next
	Next

	DBSELECTAREA( "SM2" )
	SM2->(dbSetOrder(1))
	SM2->(dbGoTop())
	If SM2->(DbSeek(DTOS(dData)))
		//Alterar
		MsProcTxt("Atualizando moeda...")
		RecLock("SM2", .F.)
		For i := 1 To Len(_aMoedas)
			&("SM2->M2_MOEDA" + StrZero(_aMoedas[i,1],1)) := _aMoedas[i,3]
		NExt
		SM2->(MsUnlock())		
	Else
		//Incluir
		MsProcTxt("Incluindo moeda...")
		RecLock("SM2", .T.)
		SM2->M2_INFORM := "S"
		SM2->M2_DATA   := dDatabase
		For i := 1 To Len(_aMoedas)
			&("SM2->M2_MOEDA" + StrZero(_aMoedas[i,1],1)) := _aMoedas[i,3] // &("nMoeda"+_aMoedas[i,1])
		NExt
		SM2->(MsUnlock())
	Endif

	If Len(_aMoedas) > 0
		MSGINFO( "Atualização de moedas", "Aviso" )
	Endif

	GetMoeda()
Return

Static Function SalvaParam()
	Local _cFile  := GetTempPath() + "\wsbbc.ini"
	Local oParam := Nil
	Local cDrive := cDir:= cNome:= cExt := ""

	SplitPath( cArquivo, @cDrive, @cDir, @cNome, @cExt )
	
	cPasta := cDrive + cDir

	oParam := FCreate(_cFile,0)
	FWrite(oParam, ";************************************" + CRLF)
	FWrite(oParam, ";Descrição: Buscar Moeda BBC" + CRLF)
	FWrite(oParam, ";Autor....: Rodrigo Araujo" + CRLF)
	FWrite(oParam, ";Programa.: " + FunName() + CRLF)
	FWrite(oParam, ";************************************" + CRLF)
	FWrite(oParam, "[Parametros]" + CRLF)
	FWrite(oParam, "cPasta   = "  + Alltrim(cPasta) + CRLF)
	FWrite(oParam, "cArquivo = "  + Alltrim(cArquivo) + CRLF)
	FWrite(oParam, "cTipo    = "  + Alltrim(cTipo) + CRLF)
	FClose(oParam)
Return

Static Function AbreParam()
	Local _cFile  	:= GetTempPath() + "\wsbbc.ini"
	cPasta   := GetPvProfString("parametros", "cPasta", GetTempPath() , _cFile)
	cArquivo := GetPvProfString("parametros", "cArquivo", "", _cFile)
	cTipo 	 := GetPvProfString("parametros", "cTipo", "", _cFile)
Return
