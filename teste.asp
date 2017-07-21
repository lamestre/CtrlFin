<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div id="tudo">

	<div id= "header">
		<h1>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h1>
	</div>
		
	<div id="conteudo">
		<hr>
			<center><h1>O R Ç A M E N T O</h1>
		<hr>
	</div>

	<!-- #include file = "conexao.asp" -->
	
	<TABLE align="center" border="1">
		<tr align="center">
			<td rowspan="2"><B>&nbsp;&nbsp;Compromisso&nbsp;&nbsp;</B></td>
			<td></td>
			<td rowspan="2"><B>&nbsp;&nbsp;Favorecido&nbsp;&nbsp;</B></td>
			<td></td>
			<td colspan="2"><B>&nbsp;&nbsp;Valor&nbsp;&nbsp;</B></td>
			<td></td>
			<td colspan="3"><B>&nbsp;&nbsp;Banco&nbsp;&nbsp;</B></td>
			<td></td>
			<td colspan="3"><B>&nbsp;&nbsp;Forma de Pgto&nbsp;&nbsp;</B></td>
			<td></td>
			<td colspan="3"><B>&nbsp;&nbsp;Vencimento&nbsp;&nbsp;</B></td>
			<td></td>
			<td colspan="3"><B>&nbsp;&nbsp;Fim&nbsp;&nbsp;</B></td>
		</tr>

<%
		 set rstOrcamento = Server.CreateObject ("ADODB.recordset")
		
		sqlOrcamento = " select * from tb_orcamento "
		
		'Abre a conexão do Record Set com o BD
		 rstOrcamento.Open sqlOrcamento, conn
		
		' rstOrcamento.Close
		set rstOrcamento=nothing
		
		
	%>
	</font>
	</TABLE>
	

	<center><input type = "submit" value = "  V O L T A R    " onClick="history.go(-1)"/></center>
	<center><input type = "submit" value = "  M E N U  " onClick="window.open('index.html','_self')"/></center>
	
	

</body>
</html>

