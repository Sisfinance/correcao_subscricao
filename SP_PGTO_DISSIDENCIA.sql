CREATE PROCEDURE SP_PGTO_DISSIDENCIA
	 @DATA_LOCAL SMALLDATETIME 
	, @CLIENTE_INICIAL INT
	, @CLIENTE_FINAL INT
AS
BEGIN
	
	SET NOCOUNT ON;

DELETE FROM 
CLIENTE_DEPOSITO_RETIRADA 
WHERE Data_Cliente_Deposito_Retirada = @DATA_LOCAL
AND 
Liquidacao_Cliente_Deposito_Retirada = 15080 
AND 
Codigo_Bvsp_Cliente_Deposito_Retirada >= @CLIENTE_INICIAL
AND 
Codigo_Bvsp_Cliente_Deposito_Retirada <= @CLIENTE_FINAL


INSERT INTO CLIENTE_DEPOSITO_RETIRADA 
	(
	Data_Cliente_Deposito_Retirada,
	Codigo_Bvsp_Cliente_Deposito_Retirada,
	Deposito_Retirada_Cliente_Deposito_Retirada,
	Tipo_Cliente_Deposito_Retirada,
	Liquidacao_Cliente_Deposito_Retirada,
	Descricao_Cliente_Deposito_Retirada,
	Valor_Cliente_Deposito_Retirada
	) 

SELECT
	Data_Arquivo_cmdf_aux,
	Cod_Cliente_cmdf_aux, 
	CASE WHEN DebCred_cmdf_aux = 'C' THEN 'D' ELSE 'R' END AS DR,
	'P' AS TipoCliDepRet,
	15080 as Liq,
	Descr_Lanca_cmdf_aux,
	abs(Valor_cmdf_aux)
FROM CMDF_AUX INNER JOIN CLIENTE ON Cod_Cliente_cmdf_aux = Codigo_Bvsp_Cliente and '00' = Codigo_carteira_Cliente
Where 
Data_Arquivo_cmdf_aux = @DATA_LOCAL 
AND
Cod_Lanca_cmdf_aux = 15080
AND 
Cod_Cliente_cmdf_aux >= @CLIENTE_INICIAL
AND 
Cod_Cliente_cmdf_aux <= @CLIENTE_FINAL
AND 
Administra_Cota_Carteira_Cliente = 'S'



DELETE FROM BOLETA_BVSP WHERE DATA_BOLETA_BVSP = @DATA_LOCAL  
AND
CODIGO_BVSP_BOLETA_BVSP >= @CLIENTE_INICIAL
AND
CODIGO_BVSP_BOLETA_BVSP <= @CLIENTE_FINAL
AND
Custodia_Boleta_Bvsp = 'DIS' 
AND
Contra_Parte_Boleta_Bvsp = 15080
AND
Integra_Boleta_Bvsp = 'N'


INSERT INTO Boleta_Bvsp (Data_Boleta_Bvsp,
                         Codigo_Bvsp_Boleta_Bvsp,
                         Codigo_Carteira_Boleta_Bvsp,
                         Codigo_Megabolsa_Boleta_Bvsp,
                         Papel_Boleta_Bvsp,
                         Compra_Venda_Boleta_Bvsp,
                         Numero_Boleta_Bvsp,
                         Qtd_Espec_Boleta_Bvsp,
                         Qtd_Total_Boleta_Bvsp,
                         Tipo_Mercado_Boleta_Bvsp,
                         Tipo_Operacao_Boleta_Bvsp,
                         Codigo_Objeto_Papel_Boleta_Bvsp,
                         Horario_Boleta_Bvsp,
                         Contra_Parte_Boleta_Bvsp,
                         Preco_Boleta_Bvsp,
                         Custodia_Boleta_Bvsp,
                         Codigo_Cliente_Inst_Boleta_Bvsp,
                         Codigo_Usuario_Inst_Boleta_Bvsp,
                         Prazo_Vencimento_Termo_Boleta_Bvsp,
                         Carteira_Gerencial_Boleta_Bvsp,
                         Carteira_Contabil_Boleta_Bvsp,
                         Integra_Boleta_Bvsp,
                         Mdc_Boleta_Bvsp,
                         Codigo_Bvsp_Copia_Boleta_Bvsp,
                         Codigo_Carteira_Copia_Boleta_Bvsp,
                         Codigo_Corretora_Boleta_Bvsp, TOTNEG_Boleta_Bvsp, LIQOPER_Boleta_Bvsp, DATA_LIQUIDACAO_BOLETA_BVSP) 
SELECT
Data_Arquivo_cmdf_aux, 
Cod_Cliente_cmdf_aux, 
'00' as cart, 
Cod_Cliente_cmdf_aux, 
Papel_cmdf_aux,
'V' as Compra_Venda,
Ident_Lanc_cmdf_aux, 
Qtd_cmdf_aux,  
Qtd_cmdf_aux, 
'VIS' as Mercado,
'NOR' as Tipo, 
Papel_cmdf_aux, 
'12:00' as Hora,
15080 as cp,  
(Valor_Emprestimo_cmdf_aux / Qtd_cmdf_aux) as Preco,
'DIS' as Custodia,
0 as CliInst,
0 as UsuInst,  
'000' as PrazoTermo,
'S' as GER,
'S' as Cont , 
'N' as Int,
'S' as mdc, 
0 as CodCopia, 
'00' as CodCartCopia, 
0 as Corret,
Valor_cmdf_aux AS TOTNEG,
0 AS LIQOPER,
Data_Arquivo_cmdf_aux AS DTLIQ
From CMDF_AUX 
Where 
Data_Arquivo_cmdf_aux = @DATA_LOCAL 
AND Cod_Lanca_cmdf_aux = 15080
AND Cod_Cliente_cmdf_aux >= @CLIENTE_INICIAL and Cod_Cliente_cmdf_aux <= @CLIENTE_FINAL

END