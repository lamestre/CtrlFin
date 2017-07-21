<%
Option Explicit
Dim codFormaPgto, codBanco
%>
<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

	<Script language="JavaScript" type="text/javascript">

		function GravaCompromisso() {
		//alert();
				document.frmCompromisso.submit();
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

	' mesRef		= Right(request("competencia"),2)
	' ano			= Left(request("competencia"),4)
	' mesDespesa 	= "15/" & mesRef & "/" & ano
	' competencia = ano & mesRef
	
	' codDespesa = request("codDespesa")
	
	' set rstDespesa = Server.CreateObject ("ADODB.recordset")
	' sqlDespesa = 				"SELECT * FROM tb_despesas AS d "
	' sqlDespesa = sqlDespesa & 	"INNER JOIN tb_orcamento AS o ON d.cod_compromisso = o.cod_compromisso "
	' sqlDespesa = sqlDespesa & 	"WHERE d.cod_despesa = " & codDespesa
	
	' rstDespesa.Open sqlDespesa, conn
	
	' if isnull(rstDespesa("vl_alternativo")) then
		' vlAlternativo = FormatNumber(0,2)
	' else
		' vlAlternativo = FormatNumber(rstDespesa("vl_alternativo"),2)
	' end if
			
	' if isnull(rstDespesa("vl_documento")) then
		' vlDocumento = FormatNumber(0,2)
	' else
		' vlDocumento = FormatNumber(rstDespesa("vl_documento"),2)
	' end if
	
	' vlPago 		= vlDocumento
	' vlResidual 	= FormatNumber(0,2)
	' dtPagamento	= right("0"&rstDespesa("vencimento"),2) & "/" & right(rstDespesa("competencia"),2) & "/" & left(rstDespesa("competencia"),4)
	' codFormaPgto= rstDespesa("cod_forma_pgto")
	' codBanco	= rstDespesa("cod_banco")
%>

<body>
<div id="tudo">

	<div id= "header">
		<h1>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h1>
	</div>
		
	<div id="conteudo">
		<hr>
			<center><h1>Compromisso</h1>
		<hr>
	</div>

	<form name="frmCompromisso" id="frmCompromisso" action="Grv_compromisso.asp" method="post">
		<TABLE align="center" border="1">
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Compromisso: &nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type=text name=compromisso style="text-align:right">&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Favorecido: &nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type=text name=favorecido style="text-align:right">&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Valor Previsto: &nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type=text name=vlPrevisto style="text-align:right">&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Vencimento: &nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type=text name=vencimento style="text-align:right">&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Forma de Pagamento&nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp; <select align=right name=codFormaPgto style="text-align:right">
										<option value="0" disabled selected >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Selecione a Forma</option>
										<option value="1">Transferência</option>
										<option value="2">Cheque</option>
										<option value="3">Pgto OnLine</option>
										<option value="4">Débito Autorizado</option>
										<option value="5">Cartão</option>
										<option value="6">Saque</option>
									</select>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Banco&nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp; <select align=right name=codBanco style="text-align:right">
										<option value="0" disabled selected>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Selecione o Banco</option>
										<option value="1">Banco do Brasil</option>
										<option value="2">Bradesco</option>
										<option value="3">Itaú</option>
									</select>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Mês Inicial: &nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type=text name=mesInicial style="text-align:right">&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Mês Final&nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type= text name=mesFinal style="text-align:right">&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
		<TABLE>
<p><p>
	<center><input type = "button" value = "  G R A V A R   " onClick="GravaCompromisso()"/></center>
	<center><input type = "button" value = "  V O L T A R   " onClick="history.go(-1)"/></center>
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