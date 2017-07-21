<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

	<!-- #include file = "conexao.asp" -->

<%
	competencia		= request("competencia")
	codDespesa		= request("codDespesa")
	vlDocumento		= replace(request("vlDocumento"),".","")
	vlDocumento		= replace(vlDocumento,",",".")
	vlDocumento		= CasaDec(vlDocumento)
	vlPago			= replace(request("vlPago"),".","")
	vlPago			= replace(vlPago,",",".")
	vlPago			= CasaDec(vlPago)
	vlResidual		= replace(request("vlResidual"),".","")
	vlResidual		= replace(vlResidual,",",".")
	vlResidual		= CasaDec(vlResidual)
	dtPagamento		= request("dtPagamento")
	codFormaPgto	= request("codFormaPgto")
	codBanco		= request("codBanco")
'response.write "*" &dtPagamento & "*"
'response.end
'	if dtPagamento 	= "" then dtPagamento = null

	if vlResidual = 0.00 then vlResidual = CasaDec(replace((vlDocumento - vlPago)/100,",","."))
	

	
'	set rstDespesa = Server.CreateObject ("ADODB.recordset")
	sqlPgto = 			" UPDATE tb_despesas SET vl_pago 			= " & vlPago
	sqlPgto = sqlPgto & ", vl_alternativo	= " & vlResidual
	sqlPgto = sqlPgto & ", dt_pgto			= "
												if dtPagamento = "" then 
													sqlPgto = sqlPgto & "Null" '& dtPagamento 
												else 
													sqlPgto = sqlPgto & " date('" & dtPagamento & "')"
												end if
	sqlPgto = sqlPgto & ", cod_forma_pgto	= " & codFormaPgto
	sqlPgto = sqlPgto & ", cod_banco		= " & codBanco
	sqlPgto = sqlPgto & " WHERE cod_despesa = " & codDespesa

'response.write(sqlPgto)
	conn.Execute sqlPgto
'	rstDespesa.Close
'	set rstDespesa=nothing
	
%>

<body>
<div id="tudo">

	<div id= "header">
		<h1>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h1>
	</div>
		
	<div id="conteudo">
		<hr>
			<center><h1>Pagamento de Despesa realizado</h1>
		<hr>
	</div>
		
	<center><input type = "button" value = " Ok " onClick="window.open('Rel_Despesas.asp?competencia=<%=competencia%>','_self')"/></center>
</div>
</body>
</html>
<%
function CasaDec(valor)
	if instr(valor,".") = 0 then valor = valor & ".00"
	CasaDec = valor
end function

%>
