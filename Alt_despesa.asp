<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

	<!-- #include file = "conexao.asp" -->

	<Script language="JavaScript" type="text/javascript">

		function GravaDespesa() {
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
	codDespesa = request("codDespesa")
	
	set rstDespesa = Server.CreateObject ("ADODB.recordset")
	sqlDespesa = " SELECT * FROM tb_despesas AS d INNER JOIN tb_orcamento AS o ON d.cod_compromisso = o.cod_compromisso WHERE d.cod_despesa = " & codDespesa
	
	rstDespesa.Open sqlDespesa, conn
	
	competencia = rstDespesa("competencia")
	
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
	
%>

<body>
<div id="tudo">

	<div id= "header">
		<h1>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h1>
	</div>
		
	<div id="conteudo">
		<hr>
			<center><h1>Alteração de Despesa</h1>
		<hr>
	</div>

	<form name="frmALtDespesa" id="frmALtDespesa" action="Grv_despesa.asp" method="post">
		<td><input type=hidden name=codDespesa value=<%=codDespesa%>></td>
		<td><input type=hidden name=competencia value=<%=competencia%>></td>
		<TABLE align="center" border="1">
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Compromisso: &nbsp;&nbsp;</B></td>
				<td></td>
				<td align="center"><B>&nbsp;&nbsp;<%=rstDespesa("compromisso")%>&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr align="center">
				<td></td>
				<td><B>&nbsp;&nbsp;Valor do Documento: &nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type=text name=vlDocumento style="text-align:right" value=<%=vlDocumento%>>&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td><B>&nbsp;&nbsp;Valor Alternativo&nbsp;&nbsp;</B></td>
				<td></td>
				<td><B>&nbsp;&nbsp;<input type= text name=vlAlternativo style="text-align:right" value=<%=vlAlternativo%>>&nbsp;&nbsp;</B></td>
				<td></td>
			</tr>
		<TABLE>
<%
	rstDespesa.Close
	set rstDespesa=nothing
		
%>
<p><p>
	<center><input type = "button" value = "  G R A V A R    " onClick="GravaDespesa()"/></center>
	<center><input type = "submit" value = "  V O L T A R    " onClick="history.go(-1)"/></center>
	<center><input type = "submit" value = "  M E N U  " onClick="window.open('CtrlFin.asp','_self')"/></center>

	</form>
</body>
</html>

<%
response.end
		set rstDespesas = Server.CreateObject ("ADODB.recordset")
		rstDespesas.Open sqlDespesas, conn
		ordem = 1
		
		do while not rstDespesas.EOF
			sqlGrava = ""
			sqlGrava = sqlGrava & "insert into tb_despesas(cod_compromisso, competencia, vl_documento) "
			sqlGrava = sqlGrava & "values (" & rstDespesas("cod_compromisso") & "," & competencia & "," & replace(rstDespesas("vl_previsto"),",",".") & ")"

response.write(ordem & " - " & sqlGrava &"<p>")
'			conn.execute sqlGrava
			ordem = ordem + 1
			rstDespesas.MoveNext
		loop
		
		rstDespesas.Close
		set rstDespesas=nothing
		conn.Close

		response.write("Incluidos " & ordem-1 & "registros.")
		
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