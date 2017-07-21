<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

	<!-- #include file = "conexao.asp" -->

	<Script language="JavaScript" type="text/javascript">

		function GravaPgto() {
		//alert();
				document.frmPgtoDespesa.submit();
			return true;
			var vlDoc = document.frmALtDespesa.vlDocumento.value;
			vlDoc = vlDoc.replace(",","")
			var vlAlt = document.frmALtDespesa.vlAlternativo.value;
				if(vlDoc < 1 || isNaN(vlDoc)) {
					window.alert("Selecionar o usuario para alteração");
					document.frmALtDespesa.vlDocumento.focus();
					return false
					}
				{document.frmALtDespesa.submit()
				return true
				}
			}
		</Script>
<%
'	mesRef		= Right(request("competencia"),2)
'	ano			= Left(request("competencia"),4)
'	mesDespesa 	= "15/" & mesRef & "/" & ano
'	competencia = ano & mesRef
	codDespesa = request("codDespesa")
	
	set rstDespesa = Server.CreateObject ("ADODB.recordset")
	sqlDespesa = 				"SELECT * FROM tb_despesas AS d "
	sqlDespesa = sqlDespesa & 	"INNER JOIN tb_orcamento AS o ON d.cod_compromisso = o.cod_compromisso "
	sqlDespesa = sqlDespesa & 	"WHERE d.cod_despesa = " & codDespesa
	
	rstDespesa.Open sqlDespesa, conn
	
	if isnull(rstDespesa("vl_alternativo")) then
		vlAlternativo = FormatNumber(0,2)
	else
		vlAlternativo = FormatNumber(rstDespesa("vl_alternativo"),2)
	end if
			
	if isnull(rstDespesa("vl_documento")) then
		vlDocumento = FormatNumber(0,2)
	else
		vlDocumento = FormatNumber(rstDespesa("vl_documento"),2)
	end if

	
	vlPago 		= vlDocumento
	vlResidual 	= FormatNumber(0,2)
	dtPagamento	= right("0"&rstDespesa("vencimento"),2) & "/" & right(rstDespesa("competencia"),2) & "/" & left(rstDespesa("competencia"),4)
	codFormaPgto= rstDespesa("cod_forma_pgto")
	codBanco	= rstDespesa("cod_banco")
%>

<body>
<div id="tudo">

	<div id= "header">
		<h1>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h1>
	</div>
		
	<div id="conteudo">
		<hr>
			<center><h1>Pagamento de Despesa</h1>
		<hr>
	</div>

	<form name="frmPgtoDespesa" id="frmPgtoDespesa" action="Grv_pagamento.asp" method="post">
		<input type=hidden name=codDespesa value=<%=codDespesa%>>
		<input type=hidden name=vlDocumento value=<%=vlDocumento%>>
		<input type=hidden name=competencia value=<%=rstDespesa("competencia")%>>
		<TABLE align="center" border="1">
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Compromisso: &nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<%=rstDespesa("compromisso")%>&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Valor do Documento: &nbsp;&nbsp;</B></td>
				<td></td>
				<td align="right"><B>&nbsp;&nbsp;<%=vlDocumento%>&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Valor Pago: &nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type=text name=vlPago style="text-align:right" value=<%=vlPago%>>&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Valor Residual&nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type= text name=vlResidual style="text-align:right" value=<%=vlResidual%>>&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Data do Pagamento&nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type= text name=dtPagamento style="text-align:right" value=<%=dtPagamento%>>&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Forma de Pagamento&nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp; <select align=right name=codFormaPgto style="text-align:right">
										<option value="0" disabled >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Selecione a Forma</option>
										<option value="1" <%if codFormaPgto=1 then response.write("selected")%>>Transferência</option>
										<option value="2" <%if codFormaPgto=2 then response.write("selected")%>>Cheque</option>
										<option value="3" <%if codFormaPgto=3 then response.write("selected")%>>Pgto OnLine</option>
										<option value="4" <%if codFormaPgto=4 then response.write("selected")%>>Débito Autorizado</option>
										<option value="5" <%if codFormaPgto=5 then response.write("selected")%>>Cartão</option>
										<option value="6" <%if codFormaPgto=6 then response.write("selected")%>>Saque</option>
										<option value="7" <%if codFormaPgto=7 then response.write("selected")%>>Parcelamento</option>
									</select>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Banco&nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp; <select align=right name=codBanco style="text-align:right">
										<option value="0" disabled selected>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Selecione o Banco</option>
										<option value="1" <%if codBanco=1 then response.write("selected")%>>Banco do Brasil</option>
										<option value="2" <%if codBanco=2 then response.write("selected")%>>Bradesco</option>
										<option value="3" <%if codBanco=3 then response.write("selected")%>>Itaú</option>
										<option value="4" <%if codBanco=4 then response.write("selected")%>>MCMM</option>
										<option value="5" <%if codBanco=5 then response.write("selected")%>>Outros</option>
									</select>
				<td></td>
			</tr>
		<TABLE>
<%
	rstDespesa.Close
	set rstDespesa=nothing
		
%>
<p><p>
	<center><input type = "button" value = "  P A G A R  " onClick="GravaPgto()"/></center>
	<center><input type = "button" value = "  V O L T A R  " onClick="history.go(-1)"/></center>
	<center><input type = "button" value = "  M E N U  " onClick="window.open('CtrlFin.asp','_self')"/></center>

	</form>
</body>
</html>

<%
		
Function teste()
	response.write("teste")
	response.end
End function	

Function NomeMes(data)
	if isnull(data) then
		mes = "-"
	else
		mes = Month(data)
		Select Case mes
			Case 1 : mes = "Janeiro"
			Case 2 : mes = "Fevereiro"
			Case 3 : mes = "Março"
			Case 4 : mes = "Abril"
			Case 5 : mes = "Maio"
			Case 6 : mes = "Junho"
			Case 7 : mes = "Julho"
			Case 8 : mes = "Agosto"
			Case 9 : mes = "Setembro"
			Case 10 : mes = "Outubro"
			Case 11 : mes = "Novembro"
			Case 12 : mes = "Dezembro"
		End Select
		mes = mes & "/" & Year(data) 
	end if
  NomeMes =  mes
End Function

%>