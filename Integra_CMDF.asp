<%
Server.ScriptTimeOut = 50000

response.Buffer = true

Retorno = request("retorno")

Data_Local = data_integra_sinacor


Set Tab_Data_Cmdf	= conn.execute("SELECT Data_cmdf_aux From CMDF_AUX WHERE Data_cmdf_aux = "&fixdate(Data_Local))
if Tab_Data_Cmdf.eof then
%>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td width="100%"><p align="center"><b><font face="Verdana" color="#800000" size="2">A T E N Ç Ã O!</font></b></td></tr>
	<tr><td width="100%"><p align="center"><font size="2" face="Verdana" color="#800000">&nbsp;
	Para essa integração ser realizada é necessário que os arquivos XML BVBG.020.01 e BVBG021.02 (referente ao antigo CMDF) do dia <%=cdate(Data_Local)%> estejam integrados!<BR></font></p></td></tr>
	<tr><td width="100%">&nbsp;&nbsp;</font></td></tr>
	</table>	
<%	INTEGRA_CMDF_XML = "N"
ELSE

	data_pregao_arquivo = data_integra_sinacor
	
	if Apenas_Clubes = "" then
		Apenas_Clubes = Trim(Request("Apenas_Clubes"))
		if Apenas_Clubes = "" then
			Apenas_Clubes = "S"
		end if	
	end if
	
	Set Tab_Flag_cliente	= conn.execute("SELECT * From Flag_cliente WHERE Codigo_Flag_Cliente = 1")
	Codigo_Corretora_Bvsp_local	= Tab_Flag_cliente("Codigo_Corretora_Bvsp_Flag_Cliente")
	
	'vem no request do integra_diaria 
	'Codigo_Inicial
	'Codigo_Final
	If Codigo_Inicial = "" Then
		Codigo_Inicial			= Cdbl(Trim(Request("Codigo_Inicial")))
		If Codigo_Inicial = "" Then
			Codigo_Inicial = 0
		end if	
	End If
	
	If Codigo_Final = "" Then	
		Codigo_Final				= Cdbl(Trim(Request("Codigo_Final")))
		if Codigo_Final = "" then
			Codigo_Final = 99999999
		end if	
	End If	
	
	'********************************* LOG *********************************
	
	'******************************************* LOG ********************************************
	Antes_Modificacao	= Antes_Modificacao & "Data_Arquivo|S|" & Data_Local
	Data_Arquivo_Modificacao = Data_Local
	Data_Integra = Data_Local
	'********************************************************************************************
	Antes_Modificacao		= "Data Inicial = " & Data_Local
	Antes_Modificacao = Antes_Modificacao	& "| Data_Integracao|S|" & Data_Local & "|S|"
	Antes_Modificacao = Antes_Modificacao	& "| Codigo Cliente Inicial = " & codigo_inicial
	Antes_Modificacao = Antes_Modificacao	& "| Codigo Cliente Final = " & codigo_final
	'***********************************************************************
	
	ano_local			= Year(Data_Local)
	mes_local			= Month(Data_Local)
	dia_local			= Day(Data_Local)
	Data_Ativo_Bvsp_Sai = cdate(Data_Local)
	
	'************************** LOG *****************************
	Cadastro_Modificacao	= "Integra_CMDF_Bvsp"
	Acao_Modificacao		= "I"
	strSQL					= "Inicio"
	%>
	<!-- #include file="../../../../../include/Grava_Alteracao.Asp" -->
	<%
	'*********************************************************************
	
	Total_Registros	= 0
	Total_Incluidos	= 0 
	Total_Checados	= 0
	INCLUI_SUB		= 0 
	INCLUI_EMPR		= 0
	Total_Sobra_Subscricao = 0
	
	If Retorno <> "S" then
	    'leio o arquivo CMDF e gravo numa tabela auxiliar os papéis que deram subscrição %>
	    <!-- #include file="Integra_CMDF_Subscricao.asp" -->    
	    <%
	    'mostro tela para entrar com o papel recibo
	    %>
	    <!-- #include file="Inclui_Subscricao.asp" -->
	<%
	'response.end
	end if
	
	If Flag_Fechamento_Clube <> "S" Then%>
	<div align="center"><center>
	<table border="0" width="99%">
	<%
	End If
	
	
	 StrSql = "Select * from Sim_Flag where nome_sim_flag = 'BAIXA_BTC_CMDF'"
	 set tab_sim_flag = conn.execute(StrSql)
	 IF not tab_sim_flag.eof THEN
	  	BAIXA_BTC_CMDF = tab_sim_flag("Parametro_Sim_Flag")
	 else
	 	BAIXA_BTC_CMDF = "N"
	 end if
	
	 
	' Aqui eu Gero a Tabela de Cliente Provisão para a data de abertura com os mesmos registros da Data anterior
	StrSQl = ""
	StrSql = StrSql & " DELETE FROM Boleta_bvsp "
	StrSql = StrSql & " WHERE "
	StrSql = StrSql & " (Data_Boleta_bvsp ="&fixdate(data_local)&")"
	StrSql = StrSql & " And "
	StrSql = StrSql & " (Custodia_Boleta_Bvsp= 'CLC') "
	StrSql = StrSql & " And "
	StrSql = StrSql & " (Codigo_Bvsp_Boleta_Bvsp >= "& cdbl(Codigo_Inicial)&")"
	StrSql = StrSql & " And "
	StrSql = StrSql & " (Codigo_Bvsp_Boleta_Bvsp <= "& cdbl(Codigo_Final)&")"
	'response.write strsql&"<BR>"
	Set Tab_Deleta_Boleta_Bvsp = Conn.Execute(StrSql)
	
	Data_Execucao_CP = Year(Data_Local)&"/"&Month(Data_Local)&"/"&Day(Data_Local)
	Tipo_Filtro = "Ler_Dia_Anterior"
	Variavel_Ordena = 	"" 
	%>
	<!-- #include file="../../../../../include/Filtra_Cliente_Provisao.asp" -->
	<%
	Data_Anterior = Year(Data_Anterior)&"/"&Month(Data_Anterior)&"/"&Day(Data_Anterior)
	StrSQl = ""
	StrSql = StrSql & " Delete From Cliente_Provisao "
	StrSql = StrSql & " WHERE "
	StrSql = StrSql & " Data_Cliente_Provisao = "& Fixdate(Data_Local)
	StrSql = StrSql & " And "
	StrSql = StrSql & " Codigo_Bvsp_Cliente_Provisao >= " & cdbl(Codigo_Inicial)
	StrSql = StrSql & " And "
	StrSql = StrSql & " Codigo_Bvsp_Cliente_Provisao <= " & cdbl(Codigo_Final)
	'response.write strsql&"<BR>"
	Set Deletar_Cliente_Provisao = Conn.Execute(StrSql)
	response.Flush()
	
	StrSQl = ""
	StrSql = StrSql & " SELECT * From Cliente_Provisao "
	StrSql = StrSql & " WHERE "
	StrSql = StrSql & " Data_Cliente_Provisao = "& Fixdate(Data_Anterior)
	StrSql = StrSql & " And "
	StrSql = StrSql & " Codigo_Bvsp_Cliente_Provisao >= " & cdbl(Codigo_Inicial)
	StrSql = StrSql & " And "
	StrSql = StrSql & " Codigo_Bvsp_Cliente_Provisao <= " & cdbl(Codigo_Final)
	'response.write strsql&"<BR>"
	Set Tab_Provisao 	= conn.execute(StrSql)
	response.Flush()
	
	While Not Tab_Provisao.Eof
		Codigo_Bvsp_Cliente_Provisao_local				=	Tab_Provisao("Codigo_Bvsp_Cliente_Provisao") 
		Tipo_Provisao_Cliente_Provisao_local			=	Trim(Tab_Provisao("Tipo_Provisao_Cliente_Provisao"))
		Descricao_Ativo_Bvsp_Cliente_Provisao_local		=	Trim(Tab_Provisao("Descricao_Ativo_Bvsp_Cliente_Provisao"))
		Valor_Cliente_Provisao_local					=	cdbl(Tab_Provisao("Valor_Cliente_Provisao"))
		Papel_Cliente_Provisao_local					=	Trim(Tab_Provisao("Papel_Cliente_Provisao"))
		Data_Provisao_Cliente_Provisao_local			=	Trim(Tab_Provisao("Data_Provisao_Cliente_Provisao"))
		Data_Pagamento_Cliente_Provisao_local			=	Trim(Tab_Provisao("Data_Pagamento_Cliente_Provisao"))
		Quantidade_Papel_Cliente_Provisao_local			=	Trim(Tab_Provisao("Quantidade_Papel_Cliente_Provisao"))
		Percentual_Cliente_Provisao_local				=	Trim(Tab_Provisao("Percentual_Cliente_Provisao"))	
		
		
		StrSql = ""
		StrSql = StrSql & " Insert Into Cliente_Provisao "
		StrSql = StrSql & " Values "
		StrSql = StrSql & " ("
		StrSql = StrSql & FixDate(Data_local)
		StrSql = StrSql & " , " & Codigo_Bvsp_Cliente_Provisao_local
		StrSql = StrSql & " , '" & Tipo_Provisao_Cliente_Provisao_local
		StrSql = StrSql & "' , '" & Descricao_Ativo_Bvsp_Cliente_Provisao_local &"'"
		StrSql = StrSql & " , '" & Papel_Cliente_Provisao_local &"'"
		StrSql = StrSql & " , " & FixDate(Data_Provisao_Cliente_Provisao_local) 
		StrSql = StrSql & " , " & FixDate(Data_Pagamento_Cliente_Provisao_local) 
		StrSql = StrSql & " , " & Troca(Quantidade_Papel_Cliente_Provisao_local) 
		StrSql = StrSql & " , " & Troca(Percentual_Cliente_Provisao_local)
		StrSql = StrSql & " , " & troca(Valor_Cliente_Provisao_Local)
		StrSql = StrSql & " )"
		Set Tab_Insere_Provisao	= Conn.Execute(StrSql)
		'response.write "aqui 1 " & StrSql &"<BR>"
		response.Flush()
		Tab_Provisao.Movenext
	Wend
	
	Cor_de_fundo	= ""
	'Percorrer o arquivo texto
	
	'response.write " BAIXA_BTC_CMDF " & BAIXA_BTC_CMDF &"<BR>"
	
	StrSql = ""
	StrSql = StrSql & " Delete From Cliente_Deposito_Retirada "
	StrSql = StrSql & " Where "
	StrSql = StrSql & " Data_Cliente_Deposito_Retirada = " & FixDate(Data_Local)
	StrSql = StrSql & " And "
	StrSql = StrSql & " (Tipo_Cliente_Deposito_Retirada = 'M' Or (Tipo_Cliente_Deposito_Retirada = 'D' and Descricao_Cliente_Deposito_Retirada not Like 'Dividendos a distribuir aos cotistas%')  "
	if BAIXA_BTC_CMDF = "S" then
		StrSql = StrSql & " or (Tipo_Cliente_Deposito_Retirada = 'P' and Descricao_Cliente_Deposito_Retirada like '%Juros sobre Emprestimo%') " 
		StrSql = StrSql & " or (Tipo_Cliente_Deposito_Retirada = 'P' and Descricao_Cliente_Deposito_Retirada  like 'Credito Repasse TX LIQ BTC%') " 
	end if	
	StrSql = StrSql & " )"
	StrSql = StrSql & " And "
	StrSql = StrSql & " Codigo_Bvsp_Cliente_Deposito_Retirada >= " & cdbl(Codigo_Inicial)
	StrSql = StrSql & " And "
	StrSql = StrSql & " Codigo_Bvsp_Cliente_Deposito_Retirada <= " & cdbl(Codigo_Final)
	'response.write StrSql &"<BR>"
	Set Deletar_Cliente_Deposito_Retirada = Conn.Execute(StrSql)
	response.Flush()
	
	Tipo_Liq_Local			= "9" 'campo para identificar se veio desse programa	
	StrSQl = ""
	StrSQl = StrSQl & " Delete From Clube_Aplicacao_Resgate "
	StrSQl = StrSQl & " Where "
	StrSQl = StrSQl & " DATA_CLUBE_APLICACAO_RESGATE =  " & FixDate(data_Local)
	StrSQl = StrSQl & " And "
	StrSQl = StrSQl & " TIPO_LIQ =  "& Tipo_Liq_Local
	StrSql = StrSql & " And "
	StrSql = StrSql & " Codigo_Bvsp_Clube_Aplicacao_Resgate >= " & cdbl(Codigo_Inicial)
	StrSql = StrSql & " And "
	StrSql = StrSql & " Codigo_Bvsp_Clube_Aplicacao_Resgate <= " & cdbl(Codigo_Final)
	set Tab_Deleta_Apl_Resg = conn.execute(StrSql)
	
	StrSQl = ""
	StrSQl = StrSQl & " Delete From Cliente_Subscricao "
	StrSQl = StrSQl & " Where "
	StrSQl = StrSQl & " DATA_Cliente_Sub =  " & FixDate(data_Local)
	StrSql = StrSql & " And "
	StrSql = StrSql & " Cod_Bvsp_Cliente_Sub >= " & cdbl(Codigo_Inicial)
	StrSql = StrSql & " And "
	StrSql = StrSql & " Cod_Bvsp_Cliente_Sub <= " & cdbl(Codigo_Final)
	set Tab_Deleta_Cliente_Sub = conn.execute(StrSql)
	
				
	if Mostra_Valores_CMDF_Bvsp = "S" then%>
	<tr>
	<td bgcolor="#A0A0A0" align="center"><font face="Verdana" size="2">Cliente</font></td>
	<td bgcolor="#A0A0A0" align="center"><font face="Verdana" size="2">Grupo</font></td>
	<td bgcolor="#A0A0A0"><font face="Verdana" size="2">Cod.Lan</font></td>
	<td bgcolor="#A0A0A0"><font face="Verdana" size="2">Deb/Cred</font></td>
	<td bgcolor="#A0A0A0"><font face="Verdana" size="2">Desc LanFin</font></td>
	<td bgcolor="#A0A0A0"><font face="Verdana" size="2">Desc RefLan</font></td>
	<td bgcolor="#A0A0A0"><font face="Verdana" size="2">Val Lanc</font></td>
	</tr>
	
	<%
	end if
	
	'************************** LOG *****************************
	Cadastro_Modificacao	= "Integra_CMDF_Bvsp"
	Acao_Modificacao		= "I"
	strSQL					= "Inicio da leitura da Tabela Auxiliar"
	%>
	<!-- #include file="../../../../../include/Grava_Alteracao.Asp" -->
	<%
	'************************************************************************
	 StrSql = "Select * from Sim_Flag where nome_sim_flag = 'CMDF_NAO_INTEGRAR' "
	 set tab_sim_flag_cmdf = conn.execute(StrSql)
	 IF not tab_sim_flag_cmdf.eof THEN
	 	TESTA_COD_CMDF = "S"
	  	codigos_nao_Integrar = tab_sim_flag_cmdf("Parametro_Sim_Flag")
	 else
		TESTA_COD_CMDF = "N"
	 end if

	StrSQl = ""
	StrSQl = StrSQl & " SELECT 	"
	StrSQl = StrSQl & " 	Tipo_Reg_cmdf_aux "
	StrSQl = StrSQl & " 	,Num_Ref_cmdf_aux "
	StrSQl = StrSQl & " 	,Data_cmdf_aux "
	StrSQl = StrSQl & " 	,Cod_Cliente_cmdf_aux "
	StrSQl = StrSQl & " 	,cliente.codigo_bvsp_cliente "
	StrSQl = StrSQl & " 	,cliente.codigo_Conta_Mae_Bvsp_cliente "
	StrSQl = StrSQl & " 	,cliente.Nome_Espec_Cliente "
	StrSQl = StrSQl & " 	,cliente.Administra_Cota_Carteira_Cliente "
	StrSQl = StrSQl & " 	,Cod_Lanca_cmdf_aux "
	StrSQl = StrSQl & " 	,Descr_Lanca_cmdf_aux "
	StrSQl = StrSQl & " 	,DebCred_cmdf_aux "
	StrSQl = StrSQl & " 	,Cod_Grupo_cmdf_aux "
	StrSQl = StrSQl & " 	,Descr_Grupo_cmdf_aux "
	StrSQl = StrSQl & " 	,Descr_Ref_Lanca_cmdf_aux "
	StrSQl = StrSQl & " 	,Sum(Valor_cmdf_aux) as Valor_cmdf_aux "
	StrSQl = StrSQl & " 	,Sum(Valor_Emprestimo_cmdf_aux) AS Valor_Emprestimo_cmdf_aux "
	StrSQl = StrSQl & " 	,Sum(Qtd_cmdf_aux) as Qtd_cmdf_aux"
	StrSQl = StrSQl & " 	,Papel_cmdf_aux "
	StrSQl = StrSQl & " 	,Isin_cmdf_aux "
	StrSQl = StrSQl & " FROM Cmdf_Aux inner join cliente on Cod_Cliente_cmdf_aux = Cliente.Codigo_Bvsp_Cliente "
	StrSQl = StrSQl & " WHERE Data_cmdf_aux = "&fixdate(data_Local)
	StrSQl = StrSQl & " 	  AND Cod_Cliente_cmdf_aux >=  "&Codigo_Inicial &" AND Cod_Cliente_cmdf_aux <=  "&Codigo_Final
	StrSQl = StrSQl & "       AND controla_carteira_cliente = 'S' and Codigo_carteira_cliente = '00' "
	
	'PEGANDO SÓ O QUE INTERESSA PARA A NOSSA INTEGRAÇÃO
	'CMDF
	'StrSQl = StrSQl & "       AND Cod_Lanca_cmdf_aux IN (878, 880, 890 ,958 ,959 ,962 ,984 ,1773 ,1774 ,1852 ,1853 ,1856 ,1957 ,1959 ,1960 ,1962 ,1965 ,2158 ,2164 ,2176 ,2177 ,2180 ,2181 ,2184 ,2185 ,2188 ,2189 ,2277 ,2278 ,2279 ,2280) "
	
	'BVBG
	StrSQl = StrSQl & "    AND Cod_Lanca_cmdf_aux IN (10166, 10168, 10188, 10192 ,10194, 10196, 11000 ,11048, 11050, 11144, 11145, 11222, 11223 ,15084 ,15062 ,15082 ,15088 ,15076 ,15112 ,15068 ,15092 ,15098 ,15094 ,10162 ,10164 ,10174 ,10176) "
	
	if TESTA_COD_CMDF = "S" then
		StrSQl = StrSQl & "    AND 	Cod_Lanca_cmdf_aux NOT IN ("&codigos_nao_Integrar&")"
	end if	
		
	'StrSQl = StrSQl & "       and administra_cota_carteira_cliente = 'S'"
	'não testar administra cota porque preciso ler só quem controla carteira também para as subscrições.
	StrSQl = StrSQl & " GROUP BY "
	StrSQl = StrSQl & " 	Tipo_Reg_cmdf_aux "
	StrSQl = StrSQl & " 	,Num_Ref_cmdf_aux "
	StrSQl = StrSQl & " 	,Data_cmdf_aux "
	StrSQl = StrSQl & " 	,Cod_Cliente_cmdf_aux "
	StrSQl = StrSQl & " 	,cliente.codigo_bvsp_cliente "
	StrSQl = StrSQl & " 	,cliente.codigo_Conta_Mae_Bvsp_cliente "
	StrSQl = StrSQl & " 	,cliente.Nome_Espec_Cliente "
	StrSQl = StrSQl & " 	,cliente.Administra_Cota_Carteira_Cliente "
	StrSQl = StrSQl & " 	,Cod_Lanca_cmdf_aux "
	StrSQl = StrSQl & " 	,Descr_Lanca_cmdf_aux "
	StrSQl = StrSQl & " 	,DebCred_cmdf_aux "
	StrSQl = StrSQl & " 	,Cod_Grupo_cmdf_aux "
	StrSQl = StrSQl & " 	,Descr_Grupo_cmdf_aux "
	StrSQl = StrSQl & " 	,Descr_Ref_Lanca_cmdf_aux "
	StrSQl = StrSQl & " 	,Papel_cmdf_aux "
	StrSQl = StrSQl & " 	,Isin_cmdf_aux "
	'response.write StrSQl &"<BR>"
	set Tab_Consulta_CMDF_AUX = conn.execute(StrSql)
	
	'While Not TheFile.AtEndOfStream
	while not Tab_Consulta_CMDF_AUX.eof
	
		Total_Registros			= Total_Registros	 + 1
	
		Codigo_Cliente_Cmdf		= Tab_Consulta_CMDF_AUX("Cod_Cliente_cmdf_aux")
	
	 If Codigo_Cliente_Cmdf >= cdbl(Codigo_Inicial) And Codigo_Cliente_Cmdf <= cdbl(Codigo_Final) Then
	
	
		Codigo_Lancamento_Cmdf  = Cdbl(Tab_Consulta_CMDF_AUX("Cod_Lanca_cmdf_aux"))
		DebCred_Cmdf			= Tab_Consulta_CMDF_AUX("DebCred_cmdf_aux")	
		Codigo_Grupo			= Tab_Consulta_CMDF_AUX("Cod_Grupo_cmdf_aux")
		Codigo_Desc_Grupo		= Codigo_Grupo&Tab_Consulta_CMDF_AUX("Descr_Grupo_cmdf_aux")
		Descricao_LanFin_Cmdf	= Tab_Consulta_CMDF_AUX("Descr_Lanca_cmdf_aux")	
		Descricao_RefLan_Cmdf  	= Tab_Consulta_CMDF_AUX("Descr_Ref_Lanca_cmdf_aux")	
		Valor_Lancamento_Cmdf	= cdbl(Tab_Consulta_CMDF_AUX("Valor_cmdf_aux"))
		Papel_Lancamento_Cmdf	= Tab_Consulta_CMDF_AUX("Papel_cmdf_aux")	
		Cod_Lan_Subscricao		= cdbl(Tab_Consulta_CMDF_AUX("Cod_Lanca_cmdf_aux"))	
		Integra_Registro			= "N"
		
	'*************************	AQUI início do teste da subscricao ****************************************************************
		if Cod_Lan_Subscricao = "15062" then
		
			data_pagamento_cmdf = Tab_Consulta_CMDF_AUX("Data_cmdf_aux")
			
			Subscricao = "S"
			
			Cod_Bvsp_Cliente_Sub 	= Tab_Consulta_CMDF_AUX("Codigo_Bvsp_Cliente")
			Nome_Cliente_Subs		= Tab_Consulta_CMDF_AUX("Nome_Espec_Cliente")
			
			'agora vem o cod RECIBO
			Papel_Recibo_Cliente_Sub = Tab_Consulta_CMDF_AUX("Papel_cmdf_aux")
			
			'if Papel_Direito_Cliente_Sub = "" then
			if Papel_Recibo_Cliente_Sub = "" then
				'response.write " entrei " & "<BR>"
				Isin = trim(Tab_Consulta_CMDF_AUX("Isin_cmdf_aux"))	
				'response.write Isin &"<BR>" 
				Especie_papel = Mid(Isin,10,2)
				'response.write "Especie_papel " & Especie_papel &"<BR>"			
				If Especie_papel = "PR" then
				 	Especie_papel = "2" 
				elseif Especie_papel = "OR" then
				 	Especie_papel = "1" 
				elseif Especie_papel = "PB" then
				 	Especie_papel = "11" 
				else	
				 	Especie_papel = "12" 
				end if
						
				Papel_Recibo_Cliente_Sub = Trim(Mid(Isin,3,4))&Especie_papel
						'response.write " aquiii " & Papel_Direito_Cliente_Sub &"<BR>"
			end if
			
			Data_PAPT_CMDF = proxdiautil(data_local,1)
			Data_PAPT_CMDF_sai = day(Data_PAPT_CMDF) &"/"&month(Data_PAPT_CMDF) &"/"& year(Data_PAPT_CMDF)
	
			Data_Ativo_Teste = proxdiautil(data_local,-1)
			Papel_Recibo_Cliente_Sub_teste = mid(Papel_Recibo_Cliente_Sub, 1, len(Papel_Recibo_Cliente_Sub) - 2) 
			
			StrSQl = ""
			StrSql = StrSql & " SELECT top 1 Ativo_Atual_Ativo_Modificacao as Papel_Direito "
			StrSql = StrSql & " FROM Ativo_Bvsp_Modificacao "
			StrSql = StrSql & " WHERE (Cod_Provento_Ativo_Modificacao = 50 or Cod_Provento_Ativo_Modificacao = 52) AND (Data_Ativo_Modificacao = "&fixdate(Data_Ativo_Teste)&") AND  "
			StrSql = StrSql & "       (Ativo_Atual_Ativo_Modificacao like '"&Papel_Recibo_Cliente_Sub_teste&"%')"
			'response.write strsql&"<BR>"
			'response.end
			set Tab_Direito = conn.execute(StrSql)
	
			if not Tab_Direito.eof then
				Papel_Direito_Cliente_Sub = Tab_Direito("Papel_Direito")
			else
				' se não achar o recibo em papt procuro em subscricao_aux
				StrSQl = ""
				StrSql = StrSql & " SELECT * From Subscricao_Aux "
				StrSql = StrSql & " WHERE "
				StrSql = StrSql & " Data_Sub_Aux = "& Fixdate(Data_Local)
				StrSql = StrSql & " And "
				StrSql = StrSql & " Papel_Recibo_Sub_Aux = '"&Papel_Recibo_Cliente_Sub&"' "
				'response.write StrSql
				'response.end
				set Tab_Sub_AUX = conn.execute(StrSql)			
				if not Tab_Sub_AUX.eof then
					Papel_Direito_Cliente_Sub = 	Tab_Sub_AUX("Papel_Direito_Sub_Aux")
				else	
					Papel_Direito_Cliente_Sub = 	Papel_Recibo_Cliente_Sub
					'Subscricao = "N"
				end if	
	
			end if
	
			Qtd_Cliente_Sub = Tab_Consulta_CMDF_AUX("Qtd_cmdf_aux")
			Valor_Cliente_Sub 	 = Tab_Consulta_CMDF_AUX("Valor_cmdf_aux")		
			
			'procurando ativo_bvsp_modificacao para pegar o papel que gerou o direito e a data do início da subscrição
			StrSQl = ""
			StrSql = StrSql & " SELECT * "
			StrSql = StrSql & " FROM Ativo_Bvsp_Modificacao "
			StrSql = StrSql & " WHERE (Cod_Provento_Ativo_Modificacao = 52) AND (Data_Ativo_Modificacao = "&fixdate(Data_Ativo_Teste)&") AND  "
			StrSql = StrSql & "       (Ativo_Atual_Ativo_Modificacao = '"&Papel_Direito_Cliente_Sub&"')"
			'response.write strsql&"<BR>"
			'response.end
			set Tab_Ativo_bvsp = conn.execute(StrSql)
	
			if not Tab_Ativo_bvsp.eof then
				'response.write " achei " &"<BR>"
				Papel_Cliente_Sub = Tab_Ativo_bvsp("Ativo_Anterior_Ativo_Modificacao")
				Data_Inicio_Cliente_Sub = Tab_Ativo_bvsp("Dt_Pgto_Dividendo_Ativo_Modificacao")
				
				'preciso dar outro select porque só tenho o papel que gerou o direito no dia que houve a subscricao (50)
				'então depois que acho a data inicio da subscrição que procuro novamente
				StrSQl = ""
				StrSql = StrSql & " SELECT * "
				StrSql = StrSql & " FROM Ativo_Bvsp_Modificacao "
				StrSql = StrSql & " WHERE (Cod_Provento_Ativo_Modificacao = 50) AND (Data_Ativo_Modificacao = "&fixdate(Data_Inicio_Cliente_Sub)&") AND  "
				StrSql = StrSql & "       (Ativo_Atual_Ativo_Modificacao = '"&Papel_Direito_Cliente_Sub&"')"
				'response.write strsql&"<BR>"
				'response.end
				set Tab_Ativo_bvsp_aux = conn.execute(StrSql)		
				if not Tab_Ativo_bvsp_aux.eof then
					'response.write " achei 2 " &"<BR>"			
					Papel_Cliente_Sub = Tab_Ativo_bvsp_aux("Ativo_Anterior_Ativo_Modificacao")
					'response.write "Papel_Cliente_Sub " & Papel_Cliente_Sub &"<BR>"
					Data_Inicio_Cliente_Sub = Tab_Ativo_bvsp_aux("Dt_Pgto_Dividendo_Ativo_Modificacao")
				else
					'response.write " entrei" 
					if mid(Papel_Direito_Cliente_Sub, 5,1) = "2" then
					 	Papel_Cliente_Sub = mid(Papel_Direito_Cliente_Sub,1,4)&"4"
					elseif mid(Papel_Direito_Cliente_Sub, 5,1) = "1" then
					 	Papel_Cliente_Sub = mid(Papel_Direito_Cliente_Sub,1,4)&"3"
					elseif mid(Papel_Direito_Cliente_Sub, 5,1) = "11" then
					 	Papel_Cliente_Sub = mid(Papel_Direito_Cliente_Sub,1,4)&"6"
					else
						Papel_Cliente_Sub = Papel_Direito_Cliente_Sub
					end if
						
					Data_Inicio_Cliente_Sub = Data_local		
				end if
	
			else
					'response.write " entrei 2 " 				
					if mid(Papel_Direito_Cliente_Sub, 5,1) = "2" then
					 	Papel_Cliente_Sub = mid(Papel_Direito_Cliente_Sub,1,4)&"4"
					elseif mid(Papel_Direito_Cliente_Sub, 5,1) = "1" then
					 	Papel_Cliente_Sub = mid(Papel_Direito_Cliente_Sub,1,4)&"3"
					elseif mid(Papel_Direito_Cliente_Sub, 5,1) = "11" then
					 	Papel_Cliente_Sub = mid(Papel_Direito_Cliente_Sub,1,4)&"6"
					else
						Papel_Cliente_Sub = Papel_Direito_Cliente_Sub
					end if
					Data_Inicio_Cliente_Sub = Data_local
			end if
			
			if Subscricao <> "N" then	
				strSQL = ""
				strSQL = strSQL & " SELECT  * "
				strSQL = strSQL & " FROM  Cliente_Subscricao "
				strSQL = strSQL & " WHERE   (Data_Cliente_Sub = "&FIXDATE(data_local)& ") AND "
				strSQL = strSQL & " 		(Cod_Bvsp_Cliente_Sub = "&Cod_Bvsp_Cliente_Sub&") AND "
				strSQL = strSQL & " 		(Papel_Recibo_Cliente_Sub	= '"&Papel_Recibo_Cliente_Sub&"')  AND "		
				strSQL = strSQL & " 		(Data_Inicio_Cliente_Sub = "&fixdate(Data_Inicio_Cliente_Sub)&") "
				strSQL = strSQL & " 		order by Papel_Cliente_Sub desc "
				'response.write strSQL
				'response.end
				Set Tab_Cliente_Subs = conn.execute(strSQL)
				
				if Tab_Cliente_Subs.eof then
					
					'verificar se essa mesma subscricao veio em outra data adiantada.
					strSQL = ""
					strSQL = strSQL & " SELECT  * "
					strSQL = strSQL & " FROM  Cliente_Subscricao "
					strSQL = strSQL & " WHERE    "
					strSQL = strSQL & " 		(Cod_Bvsp_Cliente_Sub = "&Cod_Bvsp_Cliente_Sub&") AND "
					strSQL = strSQL & " 		(Papel_Recibo_Cliente_Sub	= '"&Papel_Recibo_Cliente_Sub&"')  AND "		
					strSQL = strSQL & " 		(Valor_Cliente_Sub = "&troca(Valor_Cliente_Sub)&") AND "
					strSQL = strSQL & " 		(Qtd_Cliente_Sub= "& Troca(Qtd_Cliente_Sub)&") "
					strSQL = strSQL & " 		order by Papel_Cliente_Sub desc "
					'response.write strSQL
					'response.end
					Set Tab_Cliente_Subs_teste = conn.execute(strSQL)
					
					if Tab_Cliente_Subs_teste.eof then			
						strSQL = ""
						strSQL = strSQL & " INSERT INTO Cliente_Subscricao (Data_Cliente_Sub,"
						strSQL = strSQL &                    "Cod_Bvsp_Cliente_Sub,"
						strSQL = strSQL &                    "Papel_Cliente_Sub, "
						strSQL = strSQL & 					 "Data_Inicio_Cliente_Sub, "				
						strSQL = strSQL & 					 "Qtd_Cliente_Sub, "		
						strSQL = strSQL & 					 "Valor_Cliente_Sub, "							
						strSQL = strSQL & 					 "Papel_Direito_Cliente_Sub, "							
						strSQL = strSQL & 					 "Papel_Recibo_Cliente_Sub "									
						strSQL = strSQL &                         " ) "
						strSQL = strSQL & "VALUES (" & fixdate(data_Local)& ", "
						strSQL = strSQL &              Cod_Bvsp_Cliente_Sub& ",'"
						strSQL = strSQL &              Papel_Cliente_Sub&"' , "
						strSQL = strSQL &              fixdate(Data_Inicio_Cliente_Sub)&" , "
						strSQL = strSQL &              Troca(Qtd_Cliente_Sub)&" , "
						strSQL = strSQL &              Troca(Valor_Cliente_Sub)&" , '"
						strSQL = strSQL &              Papel_Direito_Cliente_Sub&"' ,'"
						strSQL = strSQL &              Papel_Recibo_Cliente_Sub&"'"		
						strSQL = strSQL &              ")"
						'response.write strsql &"<BR>" &"<BR>"
						'response.end				
						Set Tab_Inclui_Cli_Subs = conn.execute(strSQL)
						INCLUI_SUB = INCLUI_SUB + 1
					end if	
				else
	
						'dar um update somando o valor que tá vindo se não for igual
						Valor_Cliente_Sub_Base = Tab_Cliente_Subs("Valor_Cliente_Sub")
						Valor_Cliente_Sub_Total = Valor_Cliente_Sub_Base + Valor_Cliente_Sub
						Qtd_Cliente_Sub_Base = Tab_Cliente_Subs("Qtd_Cliente_Sub")
						Qtd_Cliente_Sub_Total = Qtd_Cliente_Sub_Base + Qtd_Cliente_Sub
						
						'response.write " entrando no update " & "<BR>"
						'response.write "Valor_Cliente_Sub_Base " &  Valor_Cliente_Sub_Base &"<BR>"
						'response.write "Valor_Cliente_Sub " & Valor_Cliente_Sub &"<BR>"				
						'response.write "Valor_Cliente_Total " & Valor_Cliente_Sub_Total &"<BR>"&"<BR>"								
						
						'response.write "Qtd_Cliente_Sub_Base " &  Qtd_Cliente_Sub_Base &"<BR>"
						'response.write "Qtd_Cliente_Sub " & Qtd_Cliente_Sub &"<BR>"				
						'response.write "Qtd_Cliente_Sub_Total " & Qtd_Cliente_Sub_Total &"<BR>"&"<BR>"		 
						
						StrSql = ""
						StrSQl = StrSQl & " Update Cliente_Subscricao "
						StrSQl = StrSQl & " Set "
						StrSQl = StrSQl & " Valor_Cliente_Sub =  " & Troca(Valor_Cliente_Sub_Total)&" ,"
						StrSQl = StrSQl & " Qtd_Cliente_Sub =  " & Troca(Qtd_Cliente_Sub_Total)
						strSQL = strSQL & " WHERE   (Data_Cliente_Sub = "&FIXDATE(data_local)& ") AND "
						strSQL = strSQL & " 		(Cod_Bvsp_Cliente_Sub = "&Cod_Bvsp_Cliente_Sub&") AND "
						strSQL = strSQL & " 		(Papel_Recibo_Cliente_Sub = '"&Papel_Recibo_Cliente_Sub&"')  AND "
						strSQL = strSQL & " 		(Data_Inicio_Cliente_Sub = "&fixdate(Data_Inicio_Cliente_Sub)&") "
						'response.write strSQL &"<BR>"
						Set Tab_Altera_Cli_Subs = conn.execute(StrSql)				
					end if		
			
				Cor_Da_Letra = "#000000"
				If Cor_de_fundo <> "#800000" Then
					If Cor_de_fundo = "#DEE0D6" Then
						Cor_de_fundo 			= ""
					Else
						Cor_de_fundo			= "#DEE0D6"
					End If
				Else
					Cor_Da_Letra = "#FFFFFF"
				End If
				
				if Mostra_Valores_CMDF_Bvsp = "S" then%>
				
				<tr>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Cod_Bvsp_Cliente_Sub& " - " & Nome_Cliente_Subs%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Cod_Lan_Subscricao%> - SUBSCRIÇÃO</font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="center"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>">Papel: <%=Papel_Cliente_Sub%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="center"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>">Direito: <%=Papel_Direito_Cliente_Sub%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="center"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>">Recibo: <%=Papel_Recibo_Cliente_Sub%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="right"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>">QTD : <%=FormatNumber(Qtd_Cliente_Sub,0)%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="right"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=FormatNumber(Valor_Cliente_Sub,2)%></font></td>
				</tr>
		
		<%		end if
		
				'incluir em deposito retirada
				'incluir em cliente_provisao
		
				Deposito_Retirada_Cliente_Deposito_Retirada = "R"
				Tipo_Cliente_Deposito_Retirada = "P"					
				Liquidacao_Cliente_Deposito_Retirada = 0
				Descricao_Cliente_Deposito_Retirada = "Pgto. de Subscricao " & Papel_Recibo_Cliente_Sub
				Descricao_Cliente_Deposito_Retirada_teste = "Pgto. de Subscricao"
							
				StrSql = ""
				StrSql = StrSql & " Delete From Cliente_Deposito_Retirada "
				StrSql = StrSql & " Where "
				StrSql = StrSql & " Data_Cliente_Deposito_Retirada = " & FixDate(data_pagamento_cmdf)
				StrSql = StrSql & " And "
				StrSql = StrSql & " (Tipo_Cliente_Deposito_Retirada = 'P') " 
				StrSql = StrSql & " And "
				StrSql = StrSql & " Codigo_Bvsp_Cliente_Deposito_Retirada = " &Cod_Bvsp_Cliente_Sub
				StrSql = StrSql & " And "
				StrSql = StrSql & " Valor_Cliente_Deposito_Retirada = " &Troca(Valor_Cliente_Sub)
				StrSql = StrSql & " And "
				StrSql = StrSql & " Descricao_Cliente_Deposito_Retirada Like '%"& Descricao_Cliente_Deposito_Retirada_teste&"%'"
				Set Deletar_Cliente_Deposito_Retirada_Subs = Conn.Execute(StrSql)
				'response.write "delete " & StrSql &"<BR>"&"<BR>"
				response.Flush()
				'****************************************************************************************************************
				
				StrSQl = ""
				StrSQl = StrSQl & " Insert Into Cliente_Deposito_Retirada "
				StrSQl = StrSQl & " ("
				StrSQl = StrSQl & " Data_Cliente_Deposito_Retirada"
				StrSQl = StrSQl & " , Codigo_Bvsp_Cliente_Deposito_Retirada"
				StrSQl = StrSQl & " , Deposito_Retirada_Cliente_Deposito_Retirada"
				StrSQl = StrSQl & " , Tipo_Cliente_Deposito_Retirada"
				StrSQl = StrSQl & " , Liquidacao_Cliente_Deposito_Retirada "
				StrSQl = StrSQl & " , Descricao_Cliente_Deposito_Retirada"
				StrSQl = StrSQl & " , Valor_Cliente_Deposito_Retirada"
				StrSQl = StrSQl & " )"
				StrSQl = StrSQl & " Values "
				StrSQl = StrSQl & " ("
				StrSQl = StrSQl & FixDate(data_pagamento_cmdf)' data da segunda coluna do CMDF que é a data pgto
				StrSQl = StrSQl & " ,   " & Cod_Bvsp_Cliente_Sub 
				StrSQl = StrSQl & " , '" & Deposito_Retirada_Cliente_Deposito_Retirada & "'"
				StrSQl = StrSQl & " , '" & Tipo_Cliente_Deposito_Retirada & "'"
				StrSQl = StrSQl & " ,   " & Liquidacao_Cliente_Deposito_Retirada 
				StrSQl = StrSQl & " , '" & Descricao_Cliente_Deposito_Retirada & "'"
				StrSQl = StrSQl & " ,   " &  Troca(Valor_Cliente_Sub)
				StrSQl = StrSQl & " )"
				Set Tab_Cliente_Deposito_Retirada = conn.execute(StrSql)
				'response.write "deposito " &  StrSql &"<BR>"&"<BR>"
				
				
				if cdate(data_local) <> cdate(data_pagamento_cmdf) then
					Tipo_Provisao_Cliente_Provisao_local = "P"
					Valor_Fator_Ativo_Modificacao_local = 0
					Valor_Cliente_Provisao_Aux = Valor_Cliente_Sub * (-1)
					
					StrSql = ""
					StrSql = StrSql & " Insert Into Cliente_Provisao "
					StrSql = StrSql & " Values "
					StrSql = StrSql & " ("
					StrSql = StrSql & Fixdate(Data_local)
					StrSql = StrSql & " , " & Cod_Bvsp_Cliente_Sub
					StrSql = StrSql & " , '" & Tipo_Provisao_Cliente_Provisao_local
					StrSql = StrSql & "' , '" & Descricao_Cliente_Deposito_Retirada
					StrSql = StrSql & "' , '" & Papel_Recibo_Cliente_Sub &"'"
					StrSql = StrSql & " , "& FixDate(Data_Local)
					StrSql = StrSql & " , " & FixDate(data_pagamento_cmdf)
					StrSql = StrSql & " , " & Troca(Qtd_Cliente_Sub)
					StrSql = StrSql & " , "  & Troca(Valor_Fator_Ativo_Modificacao_local)
					StrSql = StrSql & " , " & Troca(Valor_Cliente_Provisao_Aux)
					StrSql = StrSql & " )"				
					Set Tab_Insere_Provisao	= Conn.Execute(StrSql)
					'response.write "aqui 2 " & StrSql &"<BR>"
	
				end if
				
			end if
	
		end if	
	
	'********************AQUI fim do teste da subscricao*******************************
	
	'********************Inicio do teste da sobra de subscricao *******************************
		if Cod_Lan_Subscricao = "15068" or Cod_Lan_Subscricao = "10166" then
			Inclui_Sobra = "S"
	
			Cod_Bvsp_Cliente_Sobra 	= Tab_Consulta_CMDF_AUX("Codigo_bvsp_Cliente")
			Nome_Cliente_Sobra	 	= Tab_Consulta_CMDF_AUX("Nome_Espec_Cliente")
	
		
			if 	Inclui_Sobra = "S" then
		
				Codigo_Bvsp_Boleta_Bvsp				= Cod_Bvsp_Cliente_Sobra
				Codigo_Carteira_Boleta_Bvsp			= "00"
				Codigo_Megabolsa_Boleta_Bvsp		= Cod_Bvsp_Cliente_Sobra
				Papel_Boleta_Bvsp					= Papel_Lancamento_Cmdf
				Compra_Venda_Boleta_Bvsp			= "C"
				Numero_Boleta_Bvsp					= REPLACE(Tab_Consulta_CMDF_AUX("Qtd_cmdf_aux"), ",", "") 'Tab_Consulta_CMDF_AUX("Num_Ref_cmdf_aux")
				Numero_Boleta_Bvsp					= REPLACE(Numero_Boleta_Bvsp, ".", "")				
				Qtd_Espec_Boleta_Bvsp				= Tab_Consulta_CMDF_AUX("Qtd_cmdf_aux")
				Qtd_Total_Boleta_Bvsp				= Tab_Consulta_CMDF_AUX("Qtd_cmdf_aux")
				Tipo_Mercado_Boleta_Bvsp			= "VIS"
				Tipo_Operacao_Boleta_Bvsp			= "NOR"
				Codigo_Objeto_Papel_Boleta_Bvsp		= Papel_Lancamento_Cmdf
				Hora_teste							= time
				Hora								= DatePart("h", Hora_teste)
				Min									= DatePart("n", Hora_teste)
				Horario_Boleta_Bvsp 				= Hora&":"&Min
				Contra_Parte_Boleta_Bvsp			= 0
				if Qtd_Espec_Boleta_Bvsp > 0 then
					Preco_Boleta_Bvsp					= cdbl(Valor_Lancamento_Cmdf) / cdbl(Qtd_Espec_Boleta_Bvsp)
				else
					Preco_Boleta_Bvsp					= 0
				end if	
				Custodia_Boleta_Bvsp				= "CLC"
				Codigo_Cliente_Inst_Boleta_Bvsp		= 0 
				Codigo_Usuario_Inst_Boleta_Bvsp		= 0 
				Prazo_Vencimento_Termo_Boleta_Bvsp	= "000"
				Carteira_Gerencial_Boleta_Bvsp		= "S"
				Carteira_Contabil_Boleta_Bvsp		= "S"
				Integra_Boleta_Bvsp					= "N"
				Mdc_Boleta_Bvsp						= "N"
				Codigo_Bvsp_Copia_Boleta_Bvsp		= 0 
				Codigo_Carteira_Copia_Boleta_Bvsp	= "00"
				Codigo_Corretora_Boleta_Bvsp		= 0 
		
				strSQL = ""
				strSQL = strSQL & "INSERT INTO Boleta_Bvsp (Data_Boleta_Bvsp,"
				strSQL = strSQL &                         "Codigo_Bvsp_Boleta_Bvsp,"
				strSQL = strSQL &                         "Codigo_Carteira_Boleta_Bvsp,"
				strSQL = strSQL &                         "Codigo_Megabolsa_Boleta_Bvsp,"
				strSQL = strSQL &                         "Papel_Boleta_Bvsp,"
				strSQL = strSQL &                         "Compra_Venda_Boleta_Bvsp,"
				strSQL = strSQL &                         "Numero_Boleta_Bvsp,"
				strSQL = strSQL &                         "Qtd_Espec_Boleta_Bvsp,"
				strSQL = strSQL &                         "Qtd_Total_Boleta_Bvsp,"
				strSQL = strSQL &                         "Tipo_Mercado_Boleta_Bvsp,"
				strSQL = strSQL &                         "Tipo_Operacao_Boleta_Bvsp,"
				strSQL = strSQL &                         "Codigo_Objeto_Papel_Boleta_Bvsp,"
				strSQL = strSQL &                         "Horario_Boleta_Bvsp,"
				strSQL = strSQL &                         "Contra_Parte_Boleta_Bvsp,"
				strSQL = strSQL &                         "Preco_Boleta_Bvsp,"
				strSQL = strSQL &                         "Custodia_Boleta_Bvsp,"
				strSQL = strSQL &                         "Codigo_Cliente_Inst_Boleta_Bvsp,"
				strSQL = strSQL &                         "Codigo_Usuario_Inst_Boleta_Bvsp,"
				strSQL = strSQL &                         "Prazo_Vencimento_Termo_Boleta_Bvsp,"
				strSQL = strSQL &                         "Carteira_Gerencial_Boleta_Bvsp,"
				strSQL = strSQL &                         "Carteira_Contabil_Boleta_Bvsp,"
				strSQL = strSQL &                         "Integra_Boleta_Bvsp,"
				strSQL = strSQL &                         "Mdc_Boleta_Bvsp,"
				strSQL = strSQL &                         "Codigo_Bvsp_Copia_Boleta_Bvsp,"
				strSQL = strSQL &                         "Codigo_Carteira_Copia_Boleta_Bvsp,"
				strSQL = strSQL &                         "Codigo_Corretora_Boleta_Bvsp) "
				strSQL = strSQL & "VALUES (" & FixDate(Data_Local) & ","
				strSQL = strSQL &              Codigo_Bvsp_Boleta_Bvsp & ",'"
				strSQL = strSQL &              Codigo_Carteira_Boleta_Bvsp & "',"
				strSQL = strSQL &              Codigo_Megabolsa_Boleta_Bvsp & ",'"
				strSQL = strSQL &              Papel_Boleta_Bvsp & "','"
				strSQL = strSQL &              Compra_Venda_Boleta_Bvsp & "',"
				strSQL = strSQL &              Numero_Boleta_Bvsp & ","
				strSQL = strSQL &              Troca(Qtd_Espec_Boleta_Bvsp) & ","
				strSQL = strSQL &              Troca(Qtd_Total_Boleta_Bvsp) & ",'"
				strSQL = strSQL &              Tipo_Mercado_Boleta_Bvsp & "','"
				strSQL = strSQL &              Tipo_Operacao_Boleta_Bvsp & "','"
				strSQL = strSQL &              Codigo_Objeto_Papel_Boleta_Bvsp & "','"
				strSQL = strSQL &              Horario_Boleta_Bvsp & "',"
				strSQL = strSQL &              Contra_Parte_Boleta_Bvsp & ","
				strSQL = strSQL &              Troca(Preco_Boleta_Bvsp) & ",'"
				strSQL = strSQL &              Custodia_Boleta_Bvsp & "',"
				strSQL = strSQL &              Codigo_Cliente_Inst_Boleta_Bvsp & ","
				strSQL = strSQL &              Codigo_Usuario_Inst_Boleta_Bvsp & ",'"
				strSQL = strSQL &              Prazo_Vencimento_Termo_Boleta_Bvsp & "','"
				strSQL = strSQL &              Carteira_Gerencial_Boleta_Bvsp & "','"
				strSQL = strSQL &              Carteira_Contabil_Boleta_Bvsp & "','"
				strSQL = strSQL &              Integra_Boleta_Bvsp & "','"
				strSQL = strSQL &              Mdc_Boleta_Bvsp & "',"
				strSQL = strSQL &              Codigo_Bvsp_Copia_Boleta_Bvsp & ",'"
				strSQL = strSQL &              Codigo_Carteira_Copia_Boleta_Bvsp & "',"
				strSQL = strSQL &              Codigo_Corretora_Boleta_Bvsp & ") "
				'response.write strSQL &"<BR>" 
				if Cod_Lan_Subscricao <> "10166"  then
					Set Tab_Boleta_Bvsp = conn.execute(strSQL)
				end if
					
				Total_Sobra_Subscricao	= Total_Sobra_Subscricao + 1		
				
				'incluir deposito_retirada para clubes			
				if DebCred_Cmdf = "D" then
					Deposito_Retirada_Cliente_Deposito_Retirada = "R"
				else
					Deposito_Retirada_Cliente_Deposito_Retirada = "D"				
				end if
				
				Tipo_Cliente_Deposito_Retirada = "P"
				
				Liquidacao_Cliente_Deposito_Retirada = 0
				
				Descricao_Cliente_Deposito_Retirada = Descricao_LanFin_Cmdf & " " & Papel_Boleta_Bvsp
				
				Valor_Sobra_Sub = Tab_Consulta_CMDF_AUX("Valor_cmdf_aux")	
				
				StrSQl = ""
				StrSql = StrSql & " SELECT * From Cliente_Deposito_Retirada "
				StrSql = StrSql & " WHERE "
				StrSql = StrSql & " Data_Cliente_Deposito_Retirada = "& Fixdate(data_local)
				StrSql = StrSql & " And "
				StrSql = StrSql & " Codigo_Bvsp_Cliente_Deposito_Retirada = " &Codigo_Bvsp_Boleta_Bvsp
				StrSql = StrSql & " And "
				StrSql = StrSql & " Deposito_Retirada_Cliente_Deposito_Retirada = '" & Deposito_Retirada_Cliente_Deposito_Retirada&"' "
				StrSql = StrSql & " And "
				StrSql = StrSql & " Tipo_Cliente_Deposito_Retirada = '"&Tipo_Cliente_Deposito_Retirada&"'"
				StrSql = StrSql & " And "
				StrSql = StrSql & " Descricao_Cliente_Deposito_Retirada = '"&Descricao_Cliente_Deposito_Retirada&"'"
				StrSql = StrSql & " And "
				StrSql = StrSql & " Valor_Cliente_Deposito_Retirada = "&Troca(Valor_Sobra_Sub)
				'response.write strsql&"<BR>"
				set Tab_Consulta_DEP_RET 	= conn.execute(StrSql)
	 
				if Tab_Consulta_DEP_RET.eof then
					StrSQl = ""
					StrSQl = StrSQl & " Insert Into Cliente_Deposito_Retirada "
					StrSQl = StrSQl & " ("
					StrSQl = StrSQl & " Data_Cliente_Deposito_Retirada"
					StrSQl = StrSQl & " , Codigo_Bvsp_Cliente_Deposito_Retirada"
					StrSQl = StrSQl & " , Deposito_Retirada_Cliente_Deposito_Retirada"
					StrSQl = StrSQl & " , Tipo_Cliente_Deposito_Retirada"
					StrSQl = StrSQl & " , Liquidacao_Cliente_Deposito_Retirada "
					StrSQl = StrSQl & " , Descricao_Cliente_Deposito_Retirada"
					StrSQl = StrSQl & " , Valor_Cliente_Deposito_Retirada"
					StrSQl = StrSQl & " )"
					StrSQl = StrSQl & " Values "
					StrSQl = StrSQl & " ("
					StrSQl = StrSQl & FixDate(data_local)
					StrSQl = StrSQl & " ,   " & Codigo_Bvsp_Boleta_Bvsp 
					StrSQl = StrSQl & " , '" & Deposito_Retirada_Cliente_Deposito_Retirada & "'"
					StrSQl = StrSQl & " , '" & Tipo_Cliente_Deposito_Retirada & "'"
					StrSQl = StrSQl & " ,   " & Liquidacao_Cliente_Deposito_Retirada 
					StrSQl = StrSQl & " , '" & Descricao_Cliente_Deposito_Retirada & "'"
					StrSQl = StrSQl & " ,   " &  Troca(Valor_Sobra_Sub)
					StrSQl = StrSQl & " )"
					Set Tab_Cliente_Deposito_Retirada = conn.execute(StrSql)
				end if
				
			end if	
			
		end if
	'********************Fim do teste da sobra de subscricao *******************************
	
		' 11000 - DEPOSITO DE MARGEM DE GARANTIA  - EXIGENCIA DE MARGEM PARA GARANTIA. 
					' Debito de margem, sai dinheiro do conta corrente para pagamento de uma margem gerada.	
					' pago dinheiro de uma margem gerada.
					' Credito de Cliente_Provisao, verificando se nao é uma alteração pois pode ser um acrescimo de margfem na posicao deste cliente
					' F no conta corrente gerando um lançamento de exigencia de margem  (M) em Cliente_Deposito_Retirada (R)
		' 11048 e 11050 e 11144 e 11145 e 11222 e 11223 - RETIRADA DE MARGEM DE GARANTIA - DEVOLUCAO DO VALOR EXCEDENTE EM GARANTIA. 
					' Credito de Margem , entra dinheiro no conta corrente de uma margem depositada anteriormente.
					' Recebo dinheiro de uma margem ja depositada anteriormente.
					' Debito de Cliente_Provisao, verificando se nao é uma alteração pois pode ser uma devolucao parcial de margem
					' Credito no conta corrente gerando um lançamento de deposito de margem (M) em Cliente_Deposito_Retirada   (D)	
		' 15082 - PAGAMENTO DE DIVIDENDOS 21016 -  
					' Credito no Conta Corrente, entra dinheiro no conta corrente de dividendos de ações próprias.
					' Debito de Cliente_Provisao, verifico se existe o valor exato do dividendo
					' Credito no conta corrente gerando um lançamento de pagamento de conta corrente (D)
		' 15088 - JUROS SOBRE CAPITAL PROPRIO - Credito no Conta corrente, recebo dinheiro de juros sobre ações próprias.
					' Igual ao DIVIDENDO
					' Credito no Conta Corrente, entra dinheiro no conta corrente de juros de ações próprias.
					' Debito de Cliente_Provisao, verifico se existe o valor exato do Juros
					' Credito no conta corrente gerando um lançamento de pagamento de conta corrente (D)
		' 15092 - RENDIMENTO CAPITAL PROPRIO - Credito no Conta corrente, recebo dinheiro de juros sobre ações próprias.
					' Igual ao DIVIDENDO
					' Credito no Conta Corrente, entra dinheiro no conta corrente de juros de ações próprias.
					' Debito de Cliente_Provisao, verifico se existe o valor exato do Juros
					' Credito no conta corrente gerando um lançamento de pagamento de conta corrente (D)
		' 15076 - PAGAMENTO DE AMORTIZACAO - Credito no Conta corrente, recebo dinheiro de juros sobre ações próprias.
					' Igual ao DIVIDENDO
					' Credito no Conta Corrente, entra dinheiro no conta corrente da AMORTIZACAO.
					' Debito de Cliente_Provisao, verifico se existe o valor exato da AMORTIZACAO 
					' Credito no conta corrente gerando um lançamento de pagamento de conta corrente (D)
					
		' 15094 - RENDIMENTO CAPITAL PROPRIO LIQUIDO - Credito no Conta corrente, recebo dinheiro de juros sobre ações próprias.
					' Igual ao DIVIDENDO
					' Credito no Conta Corrente, entra dinheiro no conta corrente de juros de ações próprias.
					' Debito de Cliente_Provisao, verifico se existe o valor exato do Juros
					' Credito no conta corrente gerando um lançamento de pagamento de conta corrente (D)
		' 15084	- PAGAMENTO REFERENTE A FRAÇOES		
		' 15112 - PAGAMENTO RESTIT DE CAPITAL EM DINHEIRO
		' 15062 	- recibo de subscricao
		' 15098 - PAGAMENTO DE RESGATE DE RENDA VARIAVEL


		If Cdbl(Codigo_Lancamento_Cmdf) = 11048 Or Cdbl(Codigo_Lancamento_Cmdf) = 11050 Or Cdbl(Codigo_Lancamento_Cmdf) = 11144 Or Cdbl(Codigo_Lancamento_Cmdf) = 11145 Or Cdbl(Codigo_Lancamento_Cmdf) = 11222 Or Cdbl(Codigo_Lancamento_Cmdf) = 11223 Or Cdbl(Codigo_Lancamento_Cmdf) = 11000 Or Cdbl(Codigo_Lancamento_Cmdf) = 15082 Or Cdbl(Codigo_Lancamento_Cmdf) = 15088 Or Cdbl(Codigo_Lancamento_Cmdf) = 15092  Or Cdbl(Codigo_Lancamento_Cmdf) = 15094 Or Cdbl(Codigo_Lancamento_Cmdf) = 15084 or Cdbl(Codigo_Lancamento_Cmdf) = 15112  or Cdbl(Codigo_Lancamento_Cmdf) = 10162 or Cdbl(Codigo_Lancamento_Cmdf) = 10164  or Cdbl(Codigo_Lancamento_Cmdf) = 15098 Then
	
			Integra_Registro = "S"
			
			'RESPONSE.WRITE "Codigo_Cliente_Local (COM) " & Codigo_Cliente_Local &"<br>"
			'RESPONSE.WRITE Integra_Registro &"<br>"
			
			Nome_Cliente			= Tab_Consulta_CMDF_AUX("Nome_Espec_Cliente")
			Codigo_Cliente_Local	= Tab_Consulta_CMDF_AUX("Codigo_Bvsp_Cliente")
			Administra_Cota			= Tab_Consulta_CMDF_AUX("Administra_Cota_Carteira_Cliente")
			
			If Integra_Registro = "S" and Administra_Cota = "S" Then
				
					Inclui_Deposito_Retirada = "S"
					If Cdbl(Codigo_Lancamento_Cmdf) = 11000 Then
						' Debito de margem, sai dinheiro do conta corrente para pagamento de uma margem gerada.	
						' pago dinheiro de uma margem gerada.
						' Credito de Cliente_Provisao, verificando se nao é uma alteração pois pode ser um acrescimo de margfem na posicao deste cliente
						' Debito no conta corrente gerando um lançamento de exigencia de margem  (M) em Cliente_Deposito_Retirada (R)
						Tipo_Provisao_Cliente_Provisao					= "M"
						Tipo_Cliente_Deposito_Retirada 				= "M"
					ElseIf Cdbl(Codigo_Lancamento_Cmdf) = 11048 Or Cdbl(Codigo_Lancamento_Cmdf) = 11050 Or Cdbl(Codigo_Lancamento_Cmdf) = 11144 Or Cdbl(Codigo_Lancamento_Cmdf) = 11145 Or Cdbl(Codigo_Lancamento_Cmdf) = 11222 Or Cdbl(Codigo_Lancamento_Cmdf) = 11223 Then
						' Debito de Cliente_Provisao, verificando se nao é uma alteração pois pode ser uma devolucao parcial de margem
						' Credito no conta corrente gerando um lançamento de deposito de margem (M) em Cliente_Deposito_Retirada   (D)	
						Tipo_Provisao_Cliente_Provisao					= "M"
						Tipo_Cliente_Deposito_Retirada 				= "M"
					ElseIf Cdbl(Codigo_Lancamento_Cmdf) = 15082 Or Cdbl(Codigo_Lancamento_Cmdf) = 15088 Or Cdbl(Codigo_Lancamento_Cmdf) = 15092 Or Cdbl(Codigo_Lancamento_Cmdf) = 15094 Or Cdbl(Codigo_Lancamento_Cmdf) = 15084 or Cdbl(Codigo_Lancamento_Cmdf) = 15112 or Cdbl(Codigo_Lancamento_Cmdf) = 10162 or Cdbl(Codigo_Lancamento_Cmdf) = 10164 or Cdbl(Codigo_Lancamento_Cmdf) = 15098  Then
						Tipo_Provisao_Cliente_Provisao					= "D"
						Tipo_Cliente_Deposito_Retirada 				= "D"					
						' Credito no Conta Corrente, entra dinheiro no conta corrente de juros de ações próprias.
						' Debito de Cliente_Provisao, verifico se existe o valor exato do Juros
						' Credito no conta corrente gerando um lançamento de pagamento de conta corrente (D)
					End If
				
				'*****************************************************************************************************************************
					If Codigo_Lancamento_Cmdf = 11000 Or Codigo_Lancamento_Cmdf = 11048 Or Codigo_Lancamento_Cmdf = 11050 Or Codigo_Lancamento_Cmdf = 11144 Or Codigo_Lancamento_Cmdf = 11145 Or Codigo_Lancamento_Cmdf = 11222 Or Codigo_Lancamento_Cmdf = 11223 Then
						Descricao_Cliente_Provisao	= "%%"
					Else
						Descricao_Cliente_Provisao	= "%" & Papel_Lancamento_Cmdf & "%"
					End If
				
				Codigo_Bvsp_Cliente_Provisao	= Codigo_Cliente_Local
				Valor_Cliente_Provisao			= Valor_Lancamento_Cmdf
				Tipo_Filtro						= "Provisao_Valor_Data_Papel"
				%>
				<!-- #include file="../../../../../include/Filtra_Cliente_Provisao.asp" -->
				<%
				If Tab_Cliente_Provisao.eof then
					Tipo_Filtro						= "Provisao_Data_Papel"
					%>
					<!-- #include file="../../../../../include/Filtra_Cliente_Provisao.asp" -->
					<%
					'if 	Codigo_Bvsp_Cliente_Provisao = 24434  then
						'response.write "2 " & Sql&"<BR>"&"<BR>"
					'end if		
				end if			
				
				'response.write "aqui 1 " & Sql&"<BR>"&"<BR>"
				'if 	Codigo_Bvsp_Cliente_Provisao = 24434  then
					'response.write "1 " & Sql & "<BR>"&"<BR>"
				'end if	
					
				If Tab_Cliente_Provisao.eof then
					Tipo_Filtro						= "Provisao_Valor"
					%>
					<!-- #include file="../../../../../include/Filtra_Cliente_Provisao.asp" -->
					<%
					'if 	Codigo_Bvsp_Cliente_Provisao = 24434  then
						'response.write "2 " & Sql&"<BR>"&"<BR>"
					'end if		
				end if
				
				Achou_Descricao			= "N"
	
	
				Codigo_Bvsp_Clube_Cliente_local = Codigo_Cliente_Local
				Tipo_Filtro						= "Pegar_Dias_Clube"%>
				
				<!-- #include file="../../../../../include/Filtra_Clube_Cliente.asp" -->			
				
				<%
				if not Tab_Clube_Cliente.eof then
					'tipo_pgto = Tab_Clube_Cliente("tipo_pgto_divjur_clube_cliente")
					'a variavel vai continuar a existir mas sem funcionalidade, rodar updates
					'se for 0 - INCPL
					'se for 1 - APLICOT
					'se for 2 - PAGCOT
					
					tipo_pgto_prov_dividendo 	= Tab_Clube_Cliente("pgto_prov_dividendo_clube_cliente") 
					tipo_pgto_prov_juros 		= Tab_Clube_Cliente("pgto_prov_juros_clube_cliente")
					tipo_pgto_prov_rendimento	= Tab_Clube_Cliente("pgto_prov_rendimento_clube_cliente") 
				end if	
				'response.write "tipo_pgto" & tipo_pgto &"<BR>"
				
	
				'response.end
				
				'Se for provento e tipo_pgto for 1,3,4,5 mandar pro programa da Dani de distribuir cotas					
				If Cdbl(Codigo_Lancamento_Cmdf) = 15082 Or Cdbl(Codigo_Lancamento_Cmdf) = 15088 Or Cdbl(Codigo_Lancamento_Cmdf) = 15092 Or Cdbl(Codigo_Lancamento_Cmdf) = 15094 or Cdbl(Codigo_Lancamento_Cmdf) = 15112 or Cdbl(Codigo_Lancamento_Cmdf) = 10164 then
				'se for o 10162 q é débito, não distribuir
					If Codigo_Lancamento_Cmdf = 15082 Then
						OBS_Local = "DIVIDENDO de " & Papel_Lancamento_Cmdf 
					ElseIf Codigo_Lancamento_Cmdf = 15088 Then
						OBS_Local = "JRS CAP PROPRIO de " &  Papel_Lancamento_Cmdf 
					ElseIf Codigo_Lancamento_Cmdf = 15092 Then
						OBS_Local = "RENDIMENTO de " & Papel_Lancamento_Cmdf 
					ElseIf Codigo_Lancamento_Cmdf = 15094 Then
						OBS_Local = "RENDIMENTO LIQ de " & Papel_Lancamento_Cmdf 
					'ElseIf Codigo_Lancamento_Cmdf = 15084 Then
					'	OBS_Local = "FRACOES de " & Papel_Lancamento_Cmdf 
					ElseIf Codigo_Lancamento_Cmdf = 15112 Then
						OBS_Local = "RESTIT DE CAPITAL " & Papel_Lancamento_Cmdf 
					ElseIf Codigo_Lancamento_Cmdf = 10164 Then					
						OBS_Local = "REEMBOLSO PROVENTO BTC " & Papel_Lancamento_Cmdf 
					End If
					
					'response.write "Codigo_Lancamento_Cmdf " & Codigo_Lancamento_Cmdf &"<BR>"
					'response.write "tipo_pgto_prov_dividendo " & tipo_pgto_prov_dividendo  &"<BR>"
					'response.write "cod " & Codigo_Bvsp_Cliente_Provisao &"<BR>"				 
					
					IF 	(Cdbl(Codigo_Lancamento_Cmdf) = 15082 and (tipo_pgto_prov_dividendo = "PAGCOT" OR tipo_pgto_prov_dividendo 	= "APLICOT" OR tipo_pgto_prov_dividendo 	= "APLICOTPL" OR tipo_pgto_prov_dividendo  = "INCRESG" OR tipo_pgto_prov_dividendo  = "PAGCOI")) or _
						(Cdbl(Codigo_Lancamento_Cmdf) = 10164 and (tipo_pgto_prov_dividendo = "PAGCOT" OR tipo_pgto_prov_dividendo 	= "APLICOT" OR tipo_pgto_prov_dividendo 	= "APLICOTPL" OR tipo_pgto_prov_dividendo  = "INCRESG" OR tipo_pgto_prov_dividendo  = "PAGCOI")) or _
						(Cdbl(Codigo_Lancamento_Cmdf) = 15088 and (tipo_pgto_prov_juros 	= "PAGCOT" OR tipo_pgto_prov_juros 		= "APLICOT" OR tipo_pgto_prov_juros 	= "APLICOTPL" OR tipo_pgto_prov_juros 	 = "INCRESG" OR tipo_pgto_prov_juros 	  = "PAGCOI")) or _
						(Cdbl(Codigo_Lancamento_Cmdf) = 15092 and (tipo_pgto_prov_rendimento = "PAGCOT" OR tipo_pgto_prov_rendimento = "APLICOT" OR tipo_pgto_prov_rendimento 	= "APLICOTPL" OR tipo_pgto_prov_rendimento = "INCRESG" OR tipo_pgto_prov_rendimento = "PAGCOI")) or _
						(Cdbl(Codigo_Lancamento_Cmdf) = 15094 and (tipo_pgto_prov_rendimento = "PAGCOT" OR tipo_pgto_prov_rendimento = "APLICOT" OR tipo_pgto_prov_rendimento 	= "APLICOTPL" OR tipo_pgto_prov_rendimento = "INCRESG" OR tipo_pgto_prov_rendimento = "PAGCOI"))  then
					
						'response.write "entrei " & Codigo_Bvsp_Cliente_Provisao &"<BR>"
	
						Qtde_Total_Prov = Tab_Consulta_CMDF_AUX("Qtd_cmdf_aux")
						'response.write "Codigo_Bvsp_Cliente_Provisao " & Codigo_Bvsp_Cliente_Provisao &"<BR>"
						'response.write "Qtde_Total_Prov " & Qtde_Total_Prov &"<BR>"
						
						If not Tab_Cliente_Provisao.Eof then
							DISTRIBUI_COTAS = "S"
							Data_Auxiliar_Div_prov = Tab_Cliente_Provisao("Data_Provisao_Cliente_provisao")
						else
							Sql = ""
							Sql = Sql & " SELECT Cliente_Provisao.*, Cliente.nome_Espec_cliente, Cliente.nome_cliente "
							Sql = Sql & " FROM "
							Sql = Sql & " Cliente_Provisao "
							Sql = Sql & " INNER JOIN "
							Sql = Sql & " Cliente ON "
							Sql = Sql & " (Cliente_Provisao.Codigo_Bvsp_Cliente_Provisao = Cliente.codigo_bvsp_cliente "
							Sql = Sql & " And "
							Sql = Sql & " Cliente.Codigo_carteira_cliente = '00') "
							Sql = Sql & " Where "
							Sql = Sql & " Data_Cliente_Provisao = "&FixDate(Data_Local)
							Sql = Sql & " And "
							Sql = Sql & " Codigo_Bvsp_Cliente_Provisao = "&Codigo_Bvsp_Cliente_Provisao
							Sql = Sql & " And "
							Sql = Sql & " Tipo_Provisao_Cliente_Provisao = '"&Tipo_Provisao_Cliente_Provisao&"'"
							Sql = Sql & " And "
							Sql = Sql & " Data_Pagamento_Cliente_Provisao >= "&FixDate(Data_Local)
							Sql = Sql & " And "
							Sql = Sql & " Papel_Cliente_Provisao = '"&trim(Papel_Lancamento_Cmdf)&"'"
							Sql = Sql & " And "
							Sql = Sql & " Descricao_Ativo_Bvsp_Cliente_Provisao Like '"&Descricao_Cliente_Provisao&"'"
							Sql = Sql & " And "
							Sql = Sql & " Valor_Cliente_Provisao > "&Troca(Valor_Cliente_Provisao)
							set Tab_Cliente_Provisao_DIV_PROV = conn.execute(Sql)		
							'response.write "Não achou a provisao " & Sql &"<BR>"	
							if not Tab_Cliente_Provisao_DIV_PROV.eof then
								DISTRIBUI_COTAS = "S"
								Data_Auxiliar_Div_prov = Tab_Cliente_Provisao_DIV_PROV("Data_Provisao_Cliente_provisao")
								'response.write "agora achou " &"<BR>"&"<BR>"	
							else				
								'response.write "não achou a provisao " &"<BR>"&"<BR>"
								DISTRIBUI_COTAS = "N"
								Data_Auxiliar_Div_prov = Data_Local
							end if	
						end if	
						
						
						
						if DISTRIBUI_COTAS <> "N" then%>
						<tr>
						<td colspan="7">
						<!-- #include file="../../../../../include/distribui_prov_cotas.asp" -->
							</td>
						</tr>	
						<%	
						end if
					end if	
										
				end if 
				
				'response.end
	
					If Tab_Cliente_Provisao.Eof  And (Codigo_Lancamento_Cmdf = 15082 Or Codigo_Lancamento_Cmdf = 15088 Or Codigo_Lancamento_Cmdf = 15092 Or Codigo_Lancamento_Cmdf = 15094 Or Codigo_Lancamento_Cmdf = 15084  or Codigo_Lancamento_Cmdf = 15112 or Codigo_Lancamento_Cmdf = 15098 or Codigo_Lancamento_Cmdf = 10162 or Codigo_Lancamento_Cmdf = 10164 ) Then
						'response.write "Papel_Lancamento_Cmdf : " & Papel_Lancamento_Cmdf
						'response.write "Codigo_Bvsp_Cliente_Provisao : " & Codigo_Bvsp_Cliente_Provisao &"<BR>"
						'response.end
						Descricao_Cliente_Provisao = ""
						If Codigo_Lancamento_Cmdf = 15082  Then
							Descricao_Cliente_Provisao = "DIVIDENDO " & "%" & Papel_Lancamento_Cmdf & "%"
						ElseIf Codigo_Lancamento_Cmdf = 15088  Then
							Descricao_Cliente_Provisao = "JRS CAP PROPRIO " & "%" &  Papel_Lancamento_Cmdf & "%"
						ElseIf Codigo_Lancamento_Cmdf = 15092  Then
							Descricao_Cliente_Provisao = "RENDIMENTO " & "%" &  Papel_Lancamento_Cmdf & "%"
						ElseIf Codigo_Lancamento_Cmdf = 15094 Then
							Descricao_Cliente_Provisao = "RENDIMENTO LIQ " & "%" &  Papel_Lancamento_Cmdf & "%"	
						ElseIf Codigo_Lancamento_Cmdf = 15084 Then
							Descricao_Cliente_Provisao = "FRACOES " & "%" &  Papel_Lancamento_Cmdf & "%"
						ElseIf Codigo_Lancamento_Cmdf = 15112 Then
							Descricao_Cliente_Provisao = "RESTIT DE CAPITAL " & "%" &  Papel_Lancamento_Cmdf & "%"
						ElseIf Codigo_Lancamento_Cmdf = 15098 then
							Descricao_Cliente_Provisao = "RESGATE DE RENDA VARIAVEL " & "%" &  Papel_Lancamento_Cmdf & "%"						
						ElseIf Codigo_Lancamento_Cmdf = 10162 OR Codigo_Lancamento_Cmdf = 10164 then
							Descricao_Cliente_Provisao = "REEMBOLSO EVENTO EMPRESTIMO " & "%" &  Papel_Lancamento_Cmdf & "%"						
						End If
	
						
						Valor_Cliente_Provisao			= Valor_Lancamento_Cmdf
						Tipo_Filtro						= "Provisao_Descricao"
						%>
						<!-- #include file="../../../../../include/Filtra_Cliente_Provisao.asp" -->
					<%
						If Not Tab_cliente_provisao.Eof Then
	
							'if 	Codigo_Bvsp_Cliente_Provisao = 186421  then
							'	response.write "Valor_Cliente_Provisao : " & Tab_Cliente_Provisao("Valor_Cliente_Provisao")  &"<BR>"
							'	response.write "Valor_Lancamento_Cmdf : " & Valor_Lancamento_Cmdf&"<BR>"
							'end if	
	
							If (cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) - Cdbl(Valor_Lancamento_Cmdf) <= 0.01) And (cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) - Cdbl(Valor_Lancamento_Cmdf) >= -0.01) Then
								Achou_Descricao = "S"
								
								'if 	Codigo_Bvsp_Cliente_Provisao = 186421  then
								'	response.write "Valor_Cliente_Provisao : " & Tab_Cliente_Provisao("Valor_Cliente_Provisao")  &"<BR>"
								'	response.write "Valor_Lancamento_Cmdf : " & Valor_Lancamento_Cmdf&"<BR>"
								'	response.write sql&"<br>"&"<br>"
								'end if
							Else
								'if 	Codigo_Bvsp_Cliente_Provisao = 186421  then
								'	response.write "Valor_Cliente_Provisao : " & Tab_Cliente_Provisao("Valor_Cliente_Provisao")  &"<BR>"
								'	response.write "Valor_Lancamento_Cmdf : " & Valor_Lancamento_Cmdf&"<BR>"
								'	response.write "Valor_Cliente_Provisao - Valor_Lancamento_Cmdf  : " & formatnumber(Tab_Cliente_Provisao("Valor_Cliente_Provisao") - Valor_Lancamento_Cmdf,20) &"<BR>"
								'	response.write sql&"<br>"&"<br>"
								'end if
							End If
							'response.write sql&"<br>"
							'response.write Tab_cliente_provisao.Eof
							'response.end
						End If
					End If
				
					'response.write Codigo_Bvsp_Cliente_Provisao &"<BR>"
					'response.write Codigo_Lancamento_Cmdf &"<BR>"
					'response.write Tab_Cliente_Provisao.Eof&"<BR>"&"<BR>"
					If Tab_Cliente_Provisao.Eof Then
						If cdbl(Codigo_Lancamento_Cmdf) = 11000 Or cdbl(Codigo_Lancamento_Cmdf) = 11048 Or cdbl(Codigo_Lancamento_Cmdf) = 11050 Or cdbl(Codigo_Lancamento_Cmdf) = 11144 Or cdbl(Codigo_Lancamento_Cmdf) = 11145 Or cdbl(Codigo_Lancamento_Cmdf) = 11222 Or cdbl(Codigo_Lancamento_Cmdf) = 11223 Then
							'Descricao_Cliente_Provisao	= "%%"
							Descricao_Cliente_Provisao	= "DEPOSITO DE MARGEM%"
							'Descricao_Cliente_Provisao	= "%MARGEM%"
						'Else
							'Descricao_Cliente_Provisao = "BOVESPA"
						'End If
							Tipo_Filtro						=	"Alguns_Provisoes"				%>
							<!-- #include file="../../../../../include/Filtra_Cliente_Provisao.asp" -->
							<%
							'if 	Codigo_Bvsp_Cliente_Provisao = 24434  then
							'	response.write "PASSEI AQUI " & "<br>"
							'	response.write sql&"<BR>"&"<BR>"
							'end if
		
						End If
					End If
					
	
					'response.write Codigo_Lancamento_Cmdf&"<BR>"
				
					'response.write "tipo filtro : " & tipo_filtro&"<BR>"
					If Not Tab_Cliente_Provisao.Eof Then
						Cod_Cliente_Provisao = Tab_Cliente_Provisao("Cod_Cliente_Provisao")
						if Mostra_Valores_CMDF_Bvsp = "S" then%>
						<tr>
						<td bgcolor="#D7D6E4" colspan="7">
	                 	 <p align="center"><font face="Verdana" size="1">Provisão
	                	  Encontrada ! ! !&nbsp;</font></td>
						</tr>
						<tr>
						<td bgcolor="#D7D6E4"><p align="left"><font face="Verdana" size="1"><%=Tab_Cliente_Provisao("Codigo_Bvsp_Cliente_Provisao") & " " & mid(Tab_Cliente_Provisao("Nome_Espec_Cliente"),1,35)%></font></td>
						<td bgcolor="#D7D6E4"><p align="left"><font face="Verdana" size="1"><%=Tab_Cliente_Provisao("Tipo_Provisao_Cliente_Provisao")%></font></td>
						<td bgcolor="#D7D6E4"><p align="center"><font face="Verdana" size="1"><%=Codigo_Lancamento_Cmdf%></font></td>
						<td bgcolor="#D7D6E4"><p align="center"><font face="Verdana" size="1"><%=DebCred_Cmdf%></font></td>
						<td bgcolor="#D7D6E4" colspan="2"><p align="left"><font face="Verdana" size="1"><%=Tab_Cliente_Provisao("Descricao_Ativo_Bvsp_Cliente_Provisao")%></font></td>
						<td bgcolor="#D7D6E4"><p align="right"><font face="Verdana" size="1"><%=FormatNumber(Tab_Cliente_Provisao("Valor_Cliente_Provisao"),2)%></font></td>
						</tr>
					
						<%end if
	
				
						If Cdbl(Codigo_Lancamento_Cmdf) = 11000 or Cdbl(Codigo_Lancamento_Cmdf) = 11144 or  Cdbl(Codigo_Lancamento_Cmdf) = 11222 Then
							Valor_Cliente_Provisao					= (cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) + Cdbl(Valor_Lancamento_Cmdf))
							Descricao_Ativo_Bvsp_Cliente_Provisao	= Tab_Cliente_Provisao("Descricao_Ativo_Bvsp_Cliente_Provisao")
	
							StrSql = ""
							StrSQl = StrSQl & " Update Cliente_Provisao "
							StrSQl = StrSQl & " Set "
							StrSQl = StrSQl & " Valor_Cliente_Provisao =  " & Troca(Valor_Cliente_Provisao)
							StrSQl = StrSQl & " Where "
							StrSQl = StrSQl & " Cod_Cliente_Provisao = " & Cod_Cliente_Provisao
							Set Tab_Cliente_Provisao = conn.execute(StrSql)
							response.Flush()
						ElseIf (cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) - Cdbl(Valor_Lancamento_Cmdf) <= 0.01) And (cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) - Cdbl(Valor_Lancamento_Cmdf) >= -0.01) Then
						
							StrSQl = ""
							StrSQl = StrSQl & " Delete From Cliente_Provisao "
							StrSQl = StrSQl & " Where "
							StrSQl = StrSQl & " Cod_Cliente_Provisao = " & Cod_Cliente_Provisao
							'response.write strsql&"<BR>"
							set Tab_Deleta_Cliente_Provisao = conn.execute(StrSql)
						ElseIf cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) > Cdbl(Valor_Lancamento_Cmdf) Then
							If Cdbl(Codigo_Lancamento_Cmdf) <> 11000 And Cdbl(Codigo_Lancamento_Cmdf) <> 11048 And Cdbl(Codigo_Lancamento_Cmdf) <> 11050  And Cdbl(Codigo_Lancamento_Cmdf) <> 11144 And Cdbl(Codigo_Lancamento_Cmdf) <> 11145 And Cdbl(Codigo_Lancamento_Cmdf) <> 11222 And Cdbl(Codigo_Lancamento_Cmdf) <> 11223  Then
								if Mostra_Valores_CMDF_Bvsp = "S" then%>
									<tr>
									<td bgcolor="#800000" colspan="7"><font face="Verdana" size="1" color="#FFFFFF"><b>ATENÇÃO , O recebimento abaixo  de : <%= FormatNumber(Valor_Cliente_Provisao,2)%>  , é menor que o valor Provisionado de : <%= FormatNumber(Tab_Cliente_Provisao("Valor_Cliente_Provisao"),2)%> o Crédito sera feito e as provisões debitadas do valor total , verifique ! ! !</b></font> </td>
									</tr>
									<%		
									Cor_de_Fundo = "#800000"
								End If
							End If		
							Valor_Cliente_Provisao					= (cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) - Cdbl(Valor_Lancamento_Cmdf))
							Descricao_Ativo_Bvsp_Cliente_Provisao	= Tab_Cliente_Provisao("Descricao_Ativo_Bvsp_Cliente_Provisao")
						
							StrSQl = ""
							StrSQl = StrSQl & " Update Cliente_Provisao "
							StrSQl = StrSQl & " Set "
							StrSQl = StrSQl & " Valor_Cliente_Provisao = " & Troca(Valor_Cliente_Provisao)
							StrSQl = StrSQl & " Where "
							StrSQl = StrSQl & " Cod_Cliente_Provisao = " & Cod_Cliente_Provisao
							Set Tab_Cliente_Provisao = conn.execute(StrSql)
							'if 	Codigo_Bvsp_Cliente_Provisao = 186421  then
							'response.write StrSql &"<BR>" &"<BR>"
							'end if
							response.Flush()
						Else
						
							If (FormatNumber(Valor_Cliente_Provisao - Tab_Cliente_Provisao("Valor_Cliente_Provisao"),2)) > 0.01 Or (FormatNumber(Valor_Cliente_Provisao - Tab_Cliente_Provisao("Valor_Cliente_Provisao"),2)) < -0.01 Then
							'if 	Codigo_Bvsp_Cliente_Provisao = 186421  then
							'	response.write cdbl(FormatNumber(Valor_Cliente_Provisao,2))&"<BR>"
							'	response.write Cdbl(FormatNumber(Tab_Cliente_Provisao("Valor_Cliente_Provisao"),2))&"<BR>"
							'	response.write "AQUI : " & FormatNumber(Valor_Cliente_Provisao - Tab_Cliente_Provisao("Valor_Cliente_Provisao"),2)&"<BR>"
							'end if	
								if Mostra_Valores_CMDF_Bvsp = "S" then%>
									<tr>
									<td bgcolor="#800000" colspan="7"><font face="Verdana" size="1" color="#FFFFFF"><b>ATENÇÃO
		         	   	        	, O recebimento abaixo  de : <%= FormatNumber(Valor_Cliente_Provisao,2)%>  ,
		                      	excede ao valor depositado em Provisões de : <%= FormatNumber(Tab_Cliente_Provisao("Valor_Cliente_Provisao"),2)%> o Crédito sera feito e as provisões zeradas</b></font> </td>
									</tr>
									<%
									Cor_de_Fundo = "#800000"
								end if
							End If
							'response.write linha&"<BR>"
							'response.end
							'Inclui_Deposito_Retirada = "N"
							StrSQl = ""
							StrSQl = StrSQl & " Delete From Cliente_Provisao "
							StrSQl = StrSQl & " Where "
							StrSQl = StrSQl & " Cod_Cliente_Provisao = " & Cod_Cliente_Provisao
							'response.write strsql&"<BR>"
							set Tab_Deleta_Cliente_Provisao = conn.execute(StrSql)
							response.Flush()
						End If
						
					ElseIf Cdbl(Codigo_Lancamento_Cmdf) = 11000 or Cdbl(Codigo_Lancamento_Cmdf) = 11144 or  Cdbl(Codigo_Lancamento_Cmdf) = 11222  Then
						Descricao_Ativo_Bvsp_Cliente_Provisao	= Mid((Trim(Descricao_LanFin_Cmdf)),1,50)
						'Descricao_Ativo_Bvsp_Cliente_Provisao	= Mid((Trim(Descricao_LanFin_Cmdf)& " " & Trim(Descricao_RefLan_Cmdf)),1,50)
						StrSQl = ""
						StrSQl = StrSQl & " Insert Into Cliente_Provisao "
						StrSQl = StrSQl & " ("
						StrSQl = StrSQl & " Data_Cliente_Provisao "
						StrSQl = StrSQl & " , Codigo_Bvsp_Cliente_Provisao"
						StrSQl = StrSQl & " , Tipo_Provisao_Cliente_Provisao"
						StrSQl = StrSQl & " , Descricao_Ativo_Bvsp_Cliente_Provisao"
						StrSql = StrSql & " , Papel_Cliente_Provisao "
						StrSql = StrSql & " , Data_Provisao_Cliente_Provisao "
						StrSql = StrSql & " , Data_Pagamento_Cliente_Provisao "
						StrSql = StrSql & " , Quantidade_Papel_Cliente_Provisao "
						StrSql = StrSql & " , Percentual_Cliente_Provisao "																				
						StrSQl = StrSQl & " , Valor_Cliente_Provisao "
						StrSQl = StrSQl & " ) "
						StrSQl = StrSQl & " Values "
						StrSQl = StrSQl & " ("
						StrSQl = StrSQl & FixDate(Data_Local)
						StrSQl = StrSQl & " ,   " & Codigo_Bvsp_Cliente_Provisao 
						StrSQl = StrSQl & " , '" & Tipo_Provisao_Cliente_Provisao & "'"
						StrSQl = StrSQl & " , '" & Descricao_Ativo_Bvsp_Cliente_Provisao & "'"
						StrSql = StrSql & " , ' '" 
						StrSql = StrSql & " , " & FixDate("2001/01/01") 
						StrSql = StrSql & " , " & FixDate("2001/01/01")
						StrSql = StrSql & " , 0" 
						StrSql = StrSql & " , 0" 				
						StrSQl = StrSQl & " ,   " & Troca(Valor_Cliente_Provisao)
						StrSQl = StrSQl & " )"
						'if 	Codigo_Bvsp_Cliente_Provisao = 24434  then
						'	response.write strsql
						'end if
						'response.end
						Set Tab_Cliente_Provisao = conn.execute(StrSql)
						'response.write "aqui 3 " & StrSql &"<BR>"
						response.Flush()
					Else
						if Mostra_Valores_CMDF_Bvsp = "S" and Cdbl(Codigo_Lancamento_Cmdf) <> 10162 and Cdbl(Codigo_Lancamento_Cmdf) <> 10164  then
						'reembolso de BTC não tem provisao%>
						<tr>
						<td bgcolor="#800000" colspan="7"><font face="Verdana" size="1" color="#FFFFFF">ATENÇÃO , Não foi encontrada a provisão abaixo. O Crédito em conta corrente no valor de : <%= FormatNumber(Valor_Cliente_Provisao,2)%> , foi realizado</font> </td>
						</tr>
					<%	end if
						'Inclui_Deposito_Retirada = "N"
						if Cdbl(Codigo_Lancamento_Cmdf) <> 10162 and Cdbl(Codigo_Lancamento_Cmdf) <> 10164 then
							Cor_de_fundo = "#800000"
						end if	
					End If
	
				'credito sempre em todos os tipo_pgto, depois testo se for diferente de 0 faço o estorno os proventos
				If Inclui_Deposito_Retirada = "S" Then		
					Data_Cliente_Deposito_Retirada					= Data_Local 
					Codigo_Bvsp_Cliente_Deposito_Retirada			= Codigo_Cliente_Local
					If DebCred_Cmdf									= "D" Then
						Deposito_Retirada_Cliente_Deposito_Retirada 	= "R"
					ElseIf DebCred_Cmdf								= "C" Then
						Deposito_Retirada_Cliente_Deposito_Retirada 	= "D"
					End If				
							Liquidacao_Cliente_Deposito_Retirada			= 0
							'Descricao_Cliente_Deposito_Retirada 			= Mid((Trim(Descricao_LanFin_Cmdf)	& " " & Trim(Descricao_RefLan_Cmdf)),1,50)
							Descricao_Cliente_Deposito_Retirada 			= Mid((Trim(Descricao_LanFin_Cmdf)),1,50)
							Valor_Cliente_Deposito_Retirada				= Valor_Lancamento_Cmdf
						
						
							StrSQl = ""
							StrSQl = StrSQl & " Insert Into Cliente_Deposito_Retirada "
							StrSQl = StrSQl & " ("
							StrSQl = StrSQl & " Data_Cliente_Deposito_Retirada"
							StrSQl = StrSQl & " , Codigo_Bvsp_Cliente_Deposito_Retirada"
							StrSQl = StrSQl & " , Deposito_Retirada_Cliente_Deposito_Retirada"
							StrSQl = StrSQl & " , Tipo_Cliente_Deposito_Retirada"
							StrSQl = StrSQl & " , Liquidacao_Cliente_Deposito_Retirada "
							StrSQl = StrSQl & " , Descricao_Cliente_Deposito_Retirada"
							StrSQl = StrSQl & " , Valor_Cliente_Deposito_Retirada"
							StrSQl = StrSQl & " )"
							StrSQl = StrSQl & " Values "
							StrSQl = StrSQl & " ("
							StrSQl = StrSQl & FixDate(Data_Cliente_Deposito_Retirada)
							StrSQl = StrSQl & " ,   " & Codigo_Bvsp_Cliente_Deposito_Retirada 
							StrSQl = StrSQl & " , '" & Deposito_Retirada_Cliente_Deposito_Retirada & "'"
							StrSQl = StrSQl & " , '" & Tipo_Cliente_Deposito_Retirada & "'"
							StrSQl = StrSQl & " ,   " & Liquidacao_Cliente_Deposito_Retirada 
							StrSQl = StrSQl & " , '" & Descricao_Cliente_Deposito_Retirada & "'"
							StrSQl = StrSQl & " ,   " &  Troca(Valor_Cliente_Deposito_Retirada)
							StrSQl = StrSQl & " )"
							Set Tab_Cliente_Deposito_Retirada = conn.execute(StrSql)
							'if 	Codigo_Bvsp_Cliente_Provisao = 186421  then
							'RESPONSE.WRITE "DEP " &  StrSql &"<br>" &"<br>"
							'end if
							response.Flush()
				
							if (DISTRIBUI_COTAS <> "N") and (tipo_pgto_prov_dividendo <> "INCPL" or tipo_pgto_prov_juros <> "INCPL" or tipo_pgto_prov_rendimento <> "INCPL") and (tipo_pgto_prov_dividendo <> "INCRESG" or tipo_pgto_prov_juros <> "INCRESG" or tipo_pgto_prov_rendimento <> "INCRESG") and (tipo_pgto_prov_dividendo <> "PAGCOI" or tipo_pgto_prov_juros <> "PAGCOI" or tipo_pgto_prov_rendimento <> "PAGCOI") and (Cdbl(Codigo_Lancamento_Cmdf) = 15082 Or Cdbl(Codigo_Lancamento_Cmdf) = 15084 Or Cdbl(Codigo_Lancamento_Cmdf) = 15088 Or Cdbl(Codigo_Lancamento_Cmdf) = 15092 Or Cdbl(Codigo_Lancamento_Cmdf) = 15094 Or Cdbl(Codigo_Lancamento_Cmdf) = 10164) THEN
							
								Deposito_Retirada_Cliente_Deposito_Retirada 	= "R"
								Descricao_Cliente_Deposito_Retirada_estorno 			= "A distribuir " & Descricao_Cliente_Deposito_Retirada
	
								'If (tipo_pgto = 1) or (tipo_pgto = 2) or (Cdbl(Codigo_Lancamento_Cmdf) = 1856 and tipo_pgto = 3) or (Cdbl(Codigo_Lancamento_Cmdf) = 1957 and tipo_pgto = 4) or (Cdbl(Codigo_Lancamento_Cmdf) = 1965 and tipo_pgto = 5) or (Cdbl(Codigo_Lancamento_Cmdf) = 2164 and tipo_pgto = 5) then						
									'lanço o estorno em deposito_retirada se for PAGCOT ou APLICOT
								IF 	(Cdbl(Codigo_Lancamento_Cmdf) = 15082 and (tipo_pgto_prov_dividendo = "PAGCOT" OR tipo_pgto_prov_dividendo = "APLICOT" OR tipo_pgto_prov_dividendo = "APLICOTPL")) or _ 
								    (Cdbl(Codigo_Lancamento_Cmdf) = 15084 and (tipo_pgto_prov_dividendo = "PAGCOT" OR tipo_pgto_prov_dividendo = "APLICOT" OR tipo_pgto_prov_dividendo = "APLICOTPL")) or _ 
								    (Cdbl(Codigo_Lancamento_Cmdf) = 10164 and (tipo_pgto_prov_dividendo = "PAGCOT" OR tipo_pgto_prov_dividendo = "APLICOT" OR tipo_pgto_prov_dividendo = "APLICOTPL")) or _ 																
									(Cdbl(Codigo_Lancamento_Cmdf) = 15088 and (tipo_pgto_prov_juros 	= "PAGCOT" OR tipo_pgto_prov_juros = "APLICOT" OR tipo_pgto_prov_juros = "APLICOTPL")) or _ 
									(Cdbl(Codigo_Lancamento_Cmdf) = 15092 and (tipo_pgto_prov_rendimento = "PAGCOT" OR tipo_pgto_prov_rendimento = "APLICOT" OR tipo_pgto_prov_rendimento = "APLICOTPL")) or _ 
									(Cdbl(Codigo_Lancamento_Cmdf) = 15094 and (tipo_pgto_prov_rendimento = "PAGCOT" OR tipo_pgto_prov_rendimento = "APLICOT" OR tipo_pgto_prov_rendimento = "APLICOTPL"))  then
	
									StrSQl = ""
									StrSQl = StrSQl & " Insert Into Cliente_Deposito_Retirada "
									StrSQl = StrSQl & " ("
									StrSQl = StrSQl & " Data_Cliente_Deposito_Retirada"
									StrSQl = StrSQl & " , Codigo_Bvsp_Cliente_Deposito_Retirada"
									StrSQl = StrSQl & " , Deposito_Retirada_Cliente_Deposito_Retirada"
									StrSQl = StrSQl & " , Tipo_Cliente_Deposito_Retirada"
									StrSQl = StrSQl & " , Liquidacao_Cliente_Deposito_Retirada "
									StrSQl = StrSQl & " , Descricao_Cliente_Deposito_Retirada"
									StrSQl = StrSQl & " , Valor_Cliente_Deposito_Retirada"
									StrSQl = StrSQl & " )"
									StrSQl = StrSQl & " Values "
									StrSQl = StrSQl & " ("
									StrSQl = StrSQl & FixDate(Data_Cliente_Deposito_Retirada)
									StrSQl = StrSQl & " ,   " & Codigo_Bvsp_Cliente_Deposito_Retirada 
									StrSQl = StrSQl & " , '" & Deposito_Retirada_Cliente_Deposito_Retirada & "'"
									StrSQl = StrSQl & " , '" & Tipo_Cliente_Deposito_Retirada & "'"
									StrSQl = StrSQl & " ,   " & Liquidacao_Cliente_Deposito_Retirada 
									StrSQl = StrSQl & " , '" & Descricao_Cliente_Deposito_Retirada_estorno & "'"
									StrSQl = StrSQl & " ,   " &  Troca(Valor_Cliente_Deposito_Retirada)
									StrSQl = StrSQl & " )"
									'response.write StrSQl
									'response.end
									Set Tab_Estorno_Cliente_Deposito_Retirada = conn.execute(StrSql)
									'RESPONSE.WRITE "ESTORNO " &  StrSql &"<br>" &"<br>"
								end if	
	
								'baixa do estorno em cliente provisao
		
								'response.write " Codigo_Bvsp_Cliente_Deposito_Retirada " & Codigo_Bvsp_Cliente_Deposito_Retirada & "<BR>"
								'response.write " Descricao_Cliente_Provisao " & Descricao_Cliente_Provisao & "<BR>"
								'response.write " Valor_Cliente_Deposito_Retirada " & Valor_Cliente_Deposito_Retirada & "<BR>"
								'response.write " Descricao_Cliente_Provisao " & Descricao_Cliente_Provisao & "<BR>"
								'response.write " Papel_Lancamento_Cmdf " & Papel_Lancamento_Cmdf & "<BR>"																					
								'response.write " Data_Pagamento_Cliente_Provisao " & Data_Cliente_Deposito_Retirada & "<BR>"																												
								
								'ver filtros: 
								'Provisao_Valor_Data_Papel
								'Provisao_Valor
								'Provisao_Descricao
								
								'Tenho que multiplicar por -1 porque o a distribuir a lógica é ao contrário
								Valor_Cliente_Provisao_a_Distribuir = Valor_Cliente_Deposito_Retirada * (-1)
								
								Sql = ""
								Sql = Sql & " SELECT Cliente_Provisao.*, Cliente.nome_Espec_cliente, Cliente.nome_cliente "
								Sql = Sql & " FROM "
								Sql = Sql & " Cliente_Provisao "
								Sql = Sql & " INNER JOIN "
								Sql = Sql & " Cliente ON "
								Sql = Sql & " (Cliente_Provisao.Codigo_Bvsp_Cliente_Provisao = Cliente.codigo_bvsp_cliente "
								Sql = Sql & " And "
								Sql = Sql & " Cliente.Codigo_carteira_cliente = '00') "
								Sql = Sql & " Where "
								Sql = Sql & " Data_Cliente_Provisao = "&FixDate(Data_Cliente_Deposito_Retirada)
								Sql = Sql & " And "
								Sql = Sql & " Codigo_Bvsp_Cliente_Provisao = "&Codigo_Bvsp_Cliente_Deposito_Retirada
								Sql = Sql & " And "
								Sql = Sql & " Tipo_Provisao_Cliente_Provisao = '"&Tipo_Provisao_Cliente_Provisao&"'"
								Sql = Sql & " And "
								Sql = Sql & " Data_Pagamento_Cliente_Provisao = "&FixDate(Data_Cliente_Deposito_Retirada)
								Sql = Sql & " And "
								Sql = Sql & " Papel_Cliente_Provisao = '"&trim(Papel_Lancamento_Cmdf)&"'"
								Sql = Sql & " And "
								Sql = Sql & " Descricao_Ativo_Bvsp_Cliente_Provisao Like 'A distribuir "&Descricao_Cliente_Provisao&"'"							
								Sql = Sql & " And "
								Sql = Sql & " ( "
								Sql = Sql & " Valor_Cliente_Provisao = "&Troca(Valor_Cliente_Provisao_a_Distribuir)
								Sql = Sql & " Or "	
								Sql = Sql & " Valor_Cliente_Provisao = "&Troca(Valor_Cliente_Provisao_a_Distribuir+0.01)
								Sql = Sql & " Or "	
								Sql = Sql & " Valor_Cliente_Provisao = "&Troca(Valor_Cliente_Provisao_a_Distribuir-0.01)
								Sql = Sql & " ) order by Data_Provisao_Cliente_Provisao "
								'response.write Sql &"<BR>"
								set Tab_Cliente_Provisao = conn.execute(Sql)	
								if Tab_Cliente_Provisao.eof then
									'procurar sem data de pagamento e sem papel
									Sql = ""
									Sql = Sql & " SELECT Cliente_Provisao.*, Cliente.nome_Espec_cliente, Cliente.nome_cliente "
									Sql = Sql & " FROM "
									Sql = Sql & " Cliente_Provisao "
									Sql = Sql & " INNER JOIN "
									Sql = Sql & " Cliente ON "
									Sql = Sql & " (Cliente_Provisao.Codigo_Bvsp_Cliente_Provisao = Cliente.codigo_bvsp_cliente "
									Sql = Sql & " And "
									Sql = Sql & " Cliente.Codigo_carteira_cliente = '00') "
									Sql = Sql & " Where "
									Sql = Sql & " Data_Cliente_Provisao = "&FixDate(Data_Cliente_Deposito_Retirada)
									Sql = Sql & " And "
									Sql = Sql & " Codigo_Bvsp_Cliente_Provisao = "&Codigo_Bvsp_Cliente_Deposito_Retirada
									Sql = Sql & " And "
									Sql = Sql & " Tipo_Provisao_Cliente_Provisao = '"&Tipo_Provisao_Cliente_Provisao&"'"
									Sql = Sql & " And "
									Sql = Sql & " Descricao_Ativo_Bvsp_Cliente_Provisao Like 'A distribuir "&Descricao_Cliente_Provisao&"'"
									Sql = Sql & " And "
									Sql = Sql & " Papel_Cliente_Provisao <> 'BMF'"
									Sql = Sql & " And "
									Sql = Sql & " ( "
									Sql = Sql & " Valor_Cliente_Provisao = "&Troca(Valor_Cliente_Provisao_a_Distribuir)
									Sql = Sql & " Or "	
									Sql = Sql & " Valor_Cliente_Provisao = "&Troca(Valor_Cliente_Provisao_a_Distribuir+0.01)
									Sql = Sql & " Or "	
									Sql = Sql & " Valor_Cliente_Provisao = "&Troca(Valor_Cliente_Provisao_a_Distribuir-0.01)
									Sql = Sql & " ) order by Data_Provisao_Cliente_Provisao "								
									'response.write "Entrando no IF 1 " &  Sql&"<BR>"&"<BR>"
									'response.end
									set Tab_Cliente_Provisao = conn.execute(Sql)
									if Tab_Cliente_Provisao.eof then
										Sql = ""
										Sql = Sql & " SELECT Cliente_Provisao.*, Cliente.nome_espec_cliente, Cliente.nome_cliente "
										Sql = Sql & " FROM "
										Sql = Sql & " Cliente_Provisao "
										Sql = Sql & " INNER JOIN "
										Sql = Sql & " Cliente ON "
										Sql = Sql & " (Cliente_Provisao.Codigo_Bvsp_Cliente_Provisao = Cliente.codigo_bvsp_cliente "
										Sql = Sql & " And "
										Sql = Sql & " Cliente.Codigo_carteira_cliente = '00') "
										Sql = Sql & " Where "
										Sql = Sql & " Data_Cliente_Provisao = "&FixDate(Data_Cliente_Deposito_Retirada)
										Sql = Sql & " And "
							 			Sql = Sql & " Codigo_Bvsp_Cliente_Provisao = "&Codigo_Bvsp_Cliente_Deposito_Retirada
										Sql = Sql & " And "
										Sql = Sql & " Tipo_Provisao_Cliente_Provisao = '"&Tipo_Provisao_Cliente_Provisao&"'"
										Sql = Sql & " And "
										Sql = Sql & " Descricao_Ativo_Bvsp_Cliente_Provisao Like 'A distribuir "&Descricao_Cliente_Provisao&"'"
										'response.write "Entrando no IF 2 " & Sql&"<BR>"&"<BR>"
										'response.end
										set Tab_Cliente_Provisao = conn.execute(Sql)
									end if
								end if
								
								If Not Tab_Cliente_Provisao.Eof Then
									Cod_Cliente_Provisao = Tab_Cliente_Provisao("Cod_Cliente_Provisao")
									If (cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) - Cdbl(Valor_Cliente_Deposito_Retirada) <= 0.01) And (cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) - Cdbl(Valor_Cliente_Deposito_Retirada) >= -0.01) Then
										StrSQl = ""
										StrSQl = StrSQl & " Delete From Cliente_Provisao "
										StrSQl = StrSQl & " Where "
										StrSQl = StrSQl & " Cod_Cliente_Provisao = " & Cod_Cliente_Provisao
										'response.write StrSQl &"<BR>"
										set Tab_Deleta_Cliente_Provisao_estorno = conn.execute(StrSql)								
									elseif cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) > Cdbl(Valor_Cliente_Deposito_Retirada) Then
										Valor_Cliente_Provisao_estorno	= (cdbl(Tab_Cliente_Provisao("Valor_Cliente_Provisao")) - Cdbl(Valor_Lancamento_Cmdf))
										StrSQl = ""
										StrSQl = StrSQl & " Update Cliente_Provisao "
										StrSQl = StrSQl & " Set "
										StrSQl = StrSQl & " Valor_Cliente_Provisao = " & Troca(Valor_Cliente_Provisao_estorno)
										StrSQl = StrSQl & " Where "
										StrSQl = StrSQl & " Cod_Cliente_Provisao = " & Cod_Cliente_Provisao
										'response.write StrSQl &"<BR>"									
										Set Tab_Cliente_Provisao_estorno = conn.execute(StrSql)
									else
										StrSQl = ""
										StrSQl = StrSQl & " Delete From Cliente_Provisao "
										StrSQl = StrSQl & " Where "
										StrSQl = StrSQl & " Cod_Cliente_Provisao = " & Cod_Cliente_Provisao
										'response.write StrSQl &"<BR>"									
										set Tab_Deleta_Cliente_Provisao_estorno = conn.execute(StrSql)
									end if	
								end if	
							end if
				
							Total_Incluidos	= Total_Incluidos	 + 1
						End If
	
					
				Cor_Da_Letra = "#000000"
				If Cor_de_fundo <> "#800000" Then
					If Cor_de_fundo = "#DEE0D6" Then
						Cor_de_fundo 			= ""
					Else
						Cor_de_fundo			= "#DEE0D6"
					End If
				Else
					Cor_Da_Letra = "#FFFFFF"
				End If
				
				
				If Flag_Fechamento_Clube <> "S" Then
				if Mostra_Valores_CMDF_Bvsp = "S" then%>
				<tr>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Codigo_Cliente_Cmdf & " " & mid(Nome_Cliente,1,35)%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Codigo_Desc_Grupo%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="center"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Codigo_Lancamento_Cmdf%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="center"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=DebCred_Cmdf%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Descricao_LanFin_Cmdf%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Descricao_RefLan_Cmdf%></font></td>
				<td bgcolor="<%=Cor_de_fundo%>"><p align="right"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=FormatNumber(Valor_Lancamento_Cmdf,2)%></font></td>
				</tr>
				<%
				end if
				End If
				
				Total_Checados	= Total_Checados	 + 1
				
				
			End If ' do integra = S
		End If ' do teste do cod. CMDF
	 End If ' cod inicial e final 
	 
		 If Cor_de_fundo = "#800000" Then
			Cor_de_fundo			= ""
		End If
	 
	 	Tab_Consulta_CMDF_AUX.movenext
		'linha = TheFile.ReadLine
	Wend
	
	'response.write "saí do while  " &"<BR>"
	'***************************************** pegando os empréstimos para jogar em Deposito_retirada apenas se o cliente tiver empréstimo (somando com o IR) ********************************************
	
	IF BVSP_DBTC = "S" or BVSP_EMP_Sinacor = "S" then
	' SE A CORRETORA TIVER EMPRESTIMO
	
		  'SE A CORRETORA FOR DAR BAIXA PELO CMDF
		  if BAIXA_BTC_CMDF = "S" then
	
		  		'response.write "entrei aqui " &"<BR>"
		  
				StrSQl = ""
				StrSQl = StrSQl & " SELECT 	Tipo_Reg_cmdf_aux, Num_Ref_cmdf_aux, Data_cmdf_aux, Cod_Cliente_cmdf_aux, "
				StrSQl = StrSQl & " 		SUM(case when DebCred_cmdf_aux = 'D' then Valor_Emprestimo_cmdf_aux * (-1) else Valor_Emprestimo_cmdf_aux end) AS Valor_BTC_IR, "
				StrSQl = StrSQl & "         Qtd_cmdf_aux, Papel_cmdf_aux  "
				StrSQl = StrSQl & " FROM Cmdf_Aux inner join cliente on Cod_Cliente_cmdf_aux = Cliente.Codigo_Bvsp_Cliente "
				StrSQl = StrSQl & " WHERE Data_cmdf_aux = "&fixdate(data_local)
				StrSQl = StrSQl & " 	  AND Cod_Cliente_cmdf_aux >=  "&CDBL(Codigo_Inicial) &" AND Cod_Cliente_cmdf_aux <=  "&cdbl(Codigo_Final)
				StrSQl = StrSQl & "       AND controla_carteira_cliente = 'S' and Codigo_carteira_cliente = '00' "
				StrSQl = StrSQl & "       AND Cod_Lanca_cmdf_aux IN (10192,10194,10196,10188,10174,10176, 12112)  "			
				
				StrSQl = StrSQl & " GROUP BY Tipo_Reg_cmdf_aux, Num_Ref_cmdf_aux, Data_cmdf_aux, Cod_Cliente_cmdf_aux, "
				StrSQl = StrSQl & "          Qtd_cmdf_aux, Papel_cmdf_aux "
				StrSQl = StrSQl & " ORDER BY Cod_Cliente_cmdf_aux, Num_Ref_cmdf_aux "
				set Tab_EMP_CMDF_AUX = conn.execute(StrSql)
				
				while not Tab_EMP_CMDF_AUX.eof
				
			  		'response.write "entrei aqui 2 " &"<BR>"
					Tipo_Filtro 			= "Busca_Cliente"
					Codigo_Cliente 			= Tab_EMP_CMDF_AUX("Cod_Cliente_cmdf_aux")
					Codigo_Carteira_Cliente = "00"
						%>
					<!-- #include file="../../../../../include/Filtra_Cliente.asp" -->
						<%
					If Not Tab_Cliente.eof Then
						controla_carteira	 = Tab_Cliente("controla_carteira_cliente")
						Administra_Cota		 = Tab_Cliente("Administra_Cota_Carteira_Cliente")
						Cod_Bvsp_Cliente_Emp = Tab_Cliente("Codigo_Bvsp_Cliente")
						Nome_Cliente_Emp	 = Tab_Cliente("Nome_Espec_Cliente")
						if controla_carteira = "S" then
							Inclui_Emp = "S"
						elseif (controla_carteira = "N" and Administra_Cota = "S") then
							Cod_Bvsp_Cliente_Emp	= Tab_Cliente("codigo_Conta_Mae_Bvsp_cliente")
							Inclui_Emp = "S"
						else	
							Inclui_Emp = "N"		
						end if	
					Else
						Inclui_Emp = "N"
					end if	
								
					
					If Inclui_Emp = "S" then
					
						If Tab_EMP_CMDF_AUX("Valor_BTC_IR") < 0 Then
							Deposito_Retirada_Cliente_Deposito_Retirada = "R"
						Else
							Deposito_Retirada_Cliente_Deposito_Retirada = "D"
						End If
						
						Tipo_Filtro	=	"Provisao_Descricao"	
						Codigo_Bvsp_Cliente_Provisao	= Cod_Bvsp_Cliente_Emp
						Tipo_Provisao_Cliente_Provisao	= "P"
						Nome_Ativo_EMP					= Tab_EMP_CMDF_AUX("Papel_cmdf_aux")&" EMP"
						Num_Contrato_EMP				= Tab_EMP_CMDF_AUX("Num_Ref_cmdf_aux")
						Qtd_cmdf_aux					= formatnumber(Tab_EMP_CMDF_AUX("Qtd_cmdf_aux"),0)
						
						Descricao_Cliente_Provisao			= "%Juros sobre Emprestimo " & Nome_Ativo_EMP & " " & Num_Contrato_EMP & "%" 
					
						%>
						<!-- #include file="../../../../../include/Filtra_Cliente_Provisao.asp" -->
						<%
						'response.write sql&"<BR>"
						
						If Tab_Cliente_Provisao.Eof Then
							Descricao_Cliente_Deposito_Retirada = "Juros sobre Emprestimo " & Nome_Ativo_EMP & " " & Num_Contrato_EMP
							if Mostra_Valores_CMDF_Bvsp = "S" then%>
							<!--	<tr>
								<td bgcolor="#800000" colspan="7"><font face="Verdana" size="1" color="#FFFFFF">ATENÇÃO , Não foi encontrada a provisão abaixo. O Crédito em conta corrente no valor de : <%= FormatNumber(Tab_EMP_CMDF_AUX("Valor_BTC_IR"),2)%> , foi realizado</font> </td>
								</tr>-->
							<%	end if
						Else			
			'				Descricao_Cliente_Deposito_Retirada	= "Juros sobre Emprestimo " & Nome_Ativo_EMP & " " & Num_Contrato_EMP & " " & FormatNumber(Tab_EMP_CMDF_AUX("Qtd_cmdf_aux"),0) 
							Descricao_Cliente_Deposito_Retirada	= Tab_Cliente_Provisao("Descricao_Ativo_Bvsp_Cliente_Provisao")
						end if
										
							Valor_Cliente_Deposito_Retirada		= FormatNumber(Tab_EMP_CMDF_AUX("Valor_BTC_IR"),2)
							if Valor_Cliente_Deposito_Retirada < 0 then
							 	Valor_Cliente_Deposito_Retirada = Valor_Cliente_Deposito_Retirada * (-1)
							end if 
							'response.write "Valor_Cliente_Deposito_Retirada : " & Valor_Cliente_Deposito_Retirada &"<BR>"
							
							sql = ""
							sql = sql & " Insert into Cliente_Deposito_Retirada "
							sql = sql & " ("
							sql = sql & " Data_Cliente_Deposito_Retirada "
							sql = sql & " , Codigo_Bvsp_Cliente_Deposito_Retirada "
							sql = sql & " , Deposito_Retirada_Cliente_Deposito_Retirada "
							sql = sql & " , Tipo_Cliente_Deposito_Retirada "
							sql = sql & " , Liquidacao_Cliente_Deposito_Retirada "
							sql = sql & " , Descricao_Cliente_Deposito_Retirada "
							sql = sql & " , Valor_Cliente_Deposito_Retirada "
							sql = sql & " ) "
							sql = sql & " Values ("
							sql = sql & Fixdate(Data_Local) 
							sql = sql & " ," & Cod_Bvsp_Cliente_Emp
							sql = sql & " ,'"&Deposito_Retirada_Cliente_Deposito_Retirada&"'"
							sql = sql & " ,'P'"
							sql = sql & " ,0" 
							sql = sql & " ,'" & Descricao_Cliente_Deposito_Retirada & "'"
							sql = sql & " ," &  Troca(Valor_Cliente_Deposito_Retirada)
							sql = sql & " )"
							set Tab_Deposito_Retirada_EMP = conn.execute(sql)
							INCLUI_EMPR = INCLUI_EMPR + 1
						'	response.write sql&"<BR>"&"<BR>"
							'response.end
							
						Cor_Da_Letra = "#000000"
						If Cor_de_fundo <> "#800000" Then
							If Cor_de_fundo = "#DEE0D6" Then
								Cor_de_fundo 			= ""
							Else
								Cor_de_fundo			= "#DEE0D6"
							End If
						Else
							Cor_Da_Letra = "#FFFFFF"
						End If
			
							
							if Mostra_Valores_CMDF_Bvsp = "S" then%>
							<tr>
							<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Cod_Bvsp_Cliente_Emp & " " & mid(Nome_Cliente_Emp,1,35)%></font></td>
							<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Num_Contrato_EMP%></font></td>
							<td bgcolor="<%=Cor_de_fundo%>"><p align="center"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Nome_Ativo_EMP%></font></td>
							<td bgcolor="<%=Cor_de_fundo%>"><p align="center"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Deposito_Retirada_Cliente_Deposito_Retirada%></font></td>
							<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Descricao_Cliente_Deposito_Retirada%></font></td>
							<td bgcolor="<%=Cor_de_fundo%>"><p align="left"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=Qtd_cmdf_aux%></font></td>
							<td bgcolor="<%=Cor_de_fundo%>"><p align="right"><font face="Verdana" size="1" color="<%=Cor_Da_Letra%>"><%=FormatNumber(Valor_Cliente_Deposito_Retirada,2)%></font></td>
							</tr>			
			<%				end if
						
					End if ' o cliente controla carteira	
			
					
					If Cor_de_fundo = "#800000" Then
						Cor_de_fundo			= ""
					End If	
					
				 	Tab_EMP_CMDF_AUX.movenext
				Wend
		end if ' BAIXA_btc_cmdf
	end if ' se a corretora integra o BTC
		
	'***********************************  fim do CC dos empréstimos ***********************************************
	
	
	'***************************************** BAIXA PROV DISTRIBUI COTAS ************************************************
	'dando a baixa na provisao que veio do distribui cotas e foi lançado no CC no dia do pagamento.
	Sql = ""
	Sql = Sql & " SELECT * "
	Sql = Sql & " FROM "
	Sql = Sql & " Cliente_Deposito_Retirada "
	Sql = Sql & " Where "
	Sql = Sql & " Data_Cliente_Deposito_Retirada = "&FixDate(Data_Local)
	Sql = Sql & " And "
	Sql = Sql & " Descricao_Cliente_Deposito_Retirada Like 'Dividendos a distribuir aos cotistas%' "
	Sql = Sql & " And "
	Sql = Sql & " Codigo_Bvsp_Cliente_Deposito_Retirada >= " & cdbl(Codigo_Inicial)
	Sql = Sql & " And "
	Sql = Sql & " Codigo_Bvsp_Cliente_Deposito_Retirada <= " & cdbl(Codigo_Final)
	'response.write Sql&"<BR>"&"<BR>"
	'response.end
	set Tab_PGTO_DEP_RET = conn.execute(Sql)	
	While Not Tab_PGTO_DEP_RET.Eof 
	
		Deposito_Retirada_Cliente_Deposito_Retirada = Tab_PGTO_DEP_RET("Deposito_Retirada_Cliente_Deposito_Retirada")
		Valor_Cliente_Deposito_Retirada 		= Tab_PGTO_DEP_RET("Valor_Cliente_Deposito_Retirada")
		Codigo_Bvsp_Cliente_Provisao 			= Tab_PGTO_DEP_RET("Codigo_Bvsp_Cliente_Deposito_Retirada")
		Tipo_Provisao_Cliente_Provisao 			= "D"
		Descricao_Ativo_Bvsp_Cliente_Provisao	= Tab_PGTO_DEP_RET("Descricao_Cliente_Deposito_Retirada")
		
		Sql = ""
		Sql = Sql & " SELECT * "
		Sql = Sql & " FROM "
		Sql = Sql & " Cliente_Provisao "
		Sql = Sql & " Where "
		Sql = Sql & " Data_Cliente_Provisao  = "&FixDate(Data_Local)
		Sql = Sql & " And "
		Sql = Sql & " Codigo_Bvsp_Cliente_Provisao = "&Codigo_Bvsp_Cliente_Provisao
		Sql = Sql & " And "
		Sql = Sql & " Tipo_Provisao_Cliente_Provisao = '"&Tipo_Provisao_Cliente_Provisao&"'"	
		Sql = Sql & " And "	
		Sql = Sql & " Descricao_Ativo_Bvsp_Cliente_Provisao = '"&Descricao_Ativo_Bvsp_Cliente_Provisao&"'"
		'response.write Sql&"<BR>"&"<BR>"
		'response.end
		set Tab_PROCURA_PROV = conn.execute(Sql)			
		if not Tab_PROCURA_PROV.eof then
			
			Cod_Cliente_Provisao = Tab_PROCURA_PROV("Cod_Cliente_Provisao")
			
			Valor_Cliente_Provisao = Tab_PROCURA_PROV("Valor_Cliente_Provisao")
			
			If Deposito_Retirada_Cliente_Deposito_Retirada = "D" Then
				Valor_Cliente_Provisao	= Cdbl(Valor_Cliente_Provisao - Valor_Cliente_Deposito_Retirada)
			ElseIf Deposito_Retirada_Cliente_Deposito_Retirada = "R" Then
				Valor_Cliente_Provisao	= Cdbl(Valor_Cliente_Provisao + Valor_Cliente_Deposito_Retirada)
			End If
	
			If Valor_Cliente_Provisao > 0.01  OR  Valor_Cliente_Provisao < -0.01 Then
				StrSql = StrSql & " Update Cliente_Provisao "
				StrSql = StrSql & " Set "
				StrSql = StrSql & " Valor_Cliente_Provisao = "&Troca(Valor_Cliente_Provisao)
			Else
				StrSql = StrSql & " Delete Cliente_Provisao "
			End If
				StrSql = StrSql & " Where "
				StrSql = StrSql & " Cod_Cliente_Provisao = "&Cod_Cliente_Provisao
				'response.write "Dando baixa na Provisao q veio da Distribuição " &StrSql &"<BR>"
				Set Tab_Baixa_Provisao = conn.execute(StrSql)
		end if	
		Tab_PGTO_DEP_RET.movenext
	wend	
	
	if FLAG_PAP_REC = "S" then%>
		<form name="phoenix" method="POST" action="Altera_Grava_Subscricao.asp">
		</form>
	<%end if%>	
	
	</table>
	  </center>
	</div>
	<%
'****************** INICIO - 15080 - Pagamento de Dissidencia *******************
	
	strSQL = " delete FROM LOG_ERRO_PROC "
	Set Tab_delete_ERRO = Conn.Execute(strSQL)

	strSQL = ""
	strSQL = strSQL & "BEGIN TRY EXEC [dbo].[SP_PGTO_DISSIDENCIA] @DATA_LOCAL = "&fixdate(data_integra_sinacor)&", @CLIENTE_INICIAL = "&Codigo_Inicial&" , @CLIENTE_FINAL ="&Codigo_Final&"  END TRY "
	'response.write strSQL & "<BR>"
	strSQL = strSQL & " BEGIN CATCH BEGIN 	exec SP_GRAVA_ERRO_PROC 'SP_PGTO_DISSIDENCIA'  END END CATCH "
	Set Tab_Executa_Proc = Conn.Execute(strSQL)
	
	strSQL = " SELECT * FROM LOG_ERRO_PROC "
	Set Tab_Executa_ERRO = Conn.Execute(strSQL)
	IF NOT Tab_Executa_ERRO.EOF THEN
	%>
		<table>
		<tr>
		<td><p align="center">ERRO: <%=Tab_Executa_ERRO("DS_ERRO_SQL")%>  - LINHA: <%=Tab_Executa_ERRO("LINHA")%>  - <%=Tab_Executa_ERRO("NM_PROC")%></p></td>
		</tr>
		</table>
		<%
		'************************** LOG *****************************
		Cadastro_Modificacao	= "SP_PGTO_DISSIDENCIA_err"
		Acao_Modificacao		= "I"
		strSQL					= "SP_PGTO_DISSIDENCIA_err - " & Tab_Executa_ERRO("DS_ERRO_SQL") & " linha " & Tab_Executa_ERRO("LINHA")
		%>
		<!-- #include file="../../../../../include/Grava_Alteracao.Asp" -->
		<!-- #include file="../../../../../include/FechaBancodados.inc" -->
	<%
		response.end
	'************************** LOG *****************************
	Else
	%>
	<table>
	<tr>
	<td><p align="center">Procedure de Pagamento de Dissidencia do dia <%=cdate(data_integra_sinacor)%> executada com sucesso ! ! !</p></td>
	</tr>
	</table>	
<%
	end if

'****************** FIM - Pagamento de Dissidencia *******************
	
	
	
	'************************** LOG *****************************
	Cadastro_Modificacao	= "Integra_CMDF_Bvsp"
	Acao_Modificacao		= "I"
	strSQL					= ""
	%>
	<!-- #include file="../../../../../include/Grava_Alteracao.Asp" -->
	<%
	
	If Flag_Fechamento_Clube <> "S" Then%>
		<div align="center">
		 <center>
		<%if Mostra_Valores_CMDF_Bvsp = "S" then%>	
		
			<table border="0" width="99%" cellspacing="1" cellpadding="0">
			<tr>
			<td bgcolor="#A0A0A0" colspan="2" align="center"><font face="Verdana" size="2"><b>INTEGRAÇÃO CMDF</b></font></td>
			</tr>
			<tr>
			<td width="50%" bgcolor="#A0A0A0"><font face="Verdana" size="2">Data da Integração</font></td>
			<td width="50%" bgcolor="#A0A0A0"><p align="left"><font face="Verdana" size="2"><%=Data_ativo_bvsp_Sai%></font></td>
			</tr>
			<tr>
			<td width="50%" bgcolor="#A0A0A0"><font face="Verdana" size="2">Total de Registros</font></td>
			<td width="50%" bgcolor="#A0A0A0">	<p align="left"><font face="Verdana" size="2"><%=Total_Registros%></font></td>
			</tr>
			<tr>
			<td width="50%" bgcolor="#A0A0A0"><font face="Verdana" size="2">Total de Subscrições</font></td>
			<td width="50%" bgcolor="#A0A0A0">	<p align="left"><font face="Verdana" size="2"><%=INCLUI_SUB%></font></td>
			</tr>
			<tr>
			<td width="50%" bgcolor="#A0A0A0"><font face="Verdana" size="2">Total Sobra de Subscrições</font></td>
			<td width="50%" bgcolor="#A0A0A0">	<p align="left"><font face="Verdana" size="2"><%=Total_Sobra_Subscricao%></font></td>
			</tr>
			<tr>
			<td width="50%" bgcolor="#A0A0A0"><font face="Verdana" size="2">Total de Empréstimos</font></td>
			<td width="50%" bgcolor="#A0A0A0">	<p align="left"><font face="Verdana" size="2"><%=INCLUI_EMPR%></font></td>
			</tr>
			
			<tr>
			<td width="50%" bgcolor="#A0A0A0"><font face="Verdana" size="2">Total de Registros disponiveis para inclusão </font></td>
			<td width="50%" bgcolor="#A0A0A0">	<p align="left"><font face="Verdana" size="2"><%=Total_Checados%></font></td>
			</tr>
			
			<tr>
			<td width="50%" bgcolor="#A0A0A0"><font face="Verdana" size="2">Registros Incluídos</font></td>
			<td width="50%" bgcolor="#A0A0A0">	<p align="left"><font face="Verdana" size="2"><%=Total_Incluidos%></font></td>
			</tr>
			 <tr>
				<td width="50%"><font face="Verdana" size="2">&nbsp;&nbsp;</font></td>
			  </tr>
			  <tr>
				<td align="center" colspan="2"><font face="Verdana" size="2" color="#872d16"><B>INTEGRAÇÃO REALIZADA COM SUCESSO!</B></font></td>
			  </tr>  
			</table>
		<%end if%>
	</center>
	</div>
	<%End If%>
	
<%End if%>