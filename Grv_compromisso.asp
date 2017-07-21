<%
Option Explicit
Dim conn, sqlCompromisso, compromisso, favorecido, vlPrevisto, vencimento, codBanco, codFormaPgto, mesInicial, mesFinal
%>
<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

	<!-- #include file = "conexao.asp" -->

<%
	compromisso		= request("compromisso")
	favorecido		= request("favorecido")
	vlPrevisto		= replace(request("vlPrevisto"),",",".")
	vencimento		= request("vencimento")
	codBanco		= request("codBanco")
	codFormaPgto	= request("codFormaPgto")
	mesInicial		= "01/" & request("mesInicial")
	mesFinal		= request("mesFinal")
	
	if mesFinal <> "" then mesFinal	= "30/" & request("mesFinal")
	
	sqlCompromisso = ""
	sqlCompromisso = "INSERT INTO tb_orcamento(compromisso, favorecido, vl_previsto, vencimento, cod_banco, cod_forma_pgto, mes_inicio, mes_final) "
	sqlCompromisso = sqlCompromisso & "VALUES ('" & compromisso & "','" & favorecido & "'," & vlPrevisto & "," & vencimento & "," & codBanco & "," & codFormaPgto & ",DATE('" & mesInicial & "'), "
	if mesFinal = "" then 
		sqlCompromisso = sqlCompromisso & "Null)"
	else 
		sqlCompromisso = sqlCompromisso & " date('" & mesFinal & "')"
	end if

	conn.Execute sqlCompromisso
	
%>

<body>
<div id="tudo">

	<div id= "header">
		<h1>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h1>
	</div>
		
	<div id="conteudo">
		<hr>
			<center><h1>Novo Compromisso incluído</h1>
		<hr>
	</div>
	<center><input type = "button" value = " Ok " onClick="window.open('Inc_compromisso.asp','_self')"/></center>
</div>
</body>
</html>
