<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

	<!-- #include file = "conexao.asp" -->

<%

	codDespesa		= request("codDespesa")
	vlDocumento		= replace(request("vlDocumento"),",",".")
	vlAlternativo	= replace(request("vlAlternativo"),",",".")
	competencia		= request("competencia")

	sqlDespesa 		= " UPDATE tb_despesas SET vl_documento = " & vlDocumento & ", vl_alternativo = " & vlAlternativo & " WHERE cod_despesa = " & codDespesa

	conn.Execute sqlDespesa
	
%>

<body>
<div id="tudo">

	<div id= "header">
		<h1>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h1>
	</div>
		
	<div id="conteudo">
	<form name="frmGrvDespesa" id="frmGrvDespesa" action="Rel_Despesas.asp" method="post">
		<td><input type=hidden name=competencia value=<%=competencia%>></td>
		<hr>
			<center><h1>Alteração de Despesa realizada</h1>
		<hr>
	<center><input type = "submit" value = " Ok " /></center>
	</form>
	</div>
		
</div>
</body>
</html>
