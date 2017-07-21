<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

<%
	mesRef		= Right(request("competencia"),2)
	ano			= Left(request("competencia"),4)
	mesOrcado 	= "15/" & mesRef & "/" & ano
	competencia	= request("competencia")
%>

<body>
<div id="tudo">

	<div id= "header">
		<h1>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h1>
	</div>
		
	<div id="conteudo">
		<hr>
			<center><h1>O R Ç A M E N T O</h1>
			<%=NomeMes(mesOrcado)%>
		<hr>
	</div>

	<!-- #include file = "conexao.asp" -->
	
	<TABLE align="center" border="1">
		<tr align="center">
			<td></td>
			<td></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Compromisso&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Favorecido&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Valor&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Banco&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Forma de Pgto&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Vencimento&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Final&nbsp;&nbsp;</B></td>
			<td></td>
		</tr>

<%

		'Cria o Record Set para o BD
		set rstOrcamento = Server.CreateObject ("ADODB.recordset")
			
		sqlOrcamento = 					" select * from tb_orcamento	as o "
		sqlOrcamento = sqlOrcamento & 	"inner join tb_banco			as b on o.cod_banco		=b.cod_banco "
		sqlOrcamento = sqlOrcamento &	"inner join tb_forma_pgto 		as f on o.cod_forma_pgto=f.cod_forma_pgto "
		sqlOrcamento = sqlOrcamento &	"where mes_inicio < date('" & mesOrcado & "') and (mes_final > date('" & mesOrcado & "') or mes_final is null) order by 7,5 asc "

		'Abre a conexão do Record Set com o BD
		rstOrcamento.Open sqlOrcamento, conn
		
		total = 0
		ordem = 1
		
		'Abre loop
		do while not rstOrcamento.EOF
%>	
			<TR align="center">
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=ordem%> &nbsp;</TD>
				<td></td>
				<TD align="left">&nbsp;<%=rstOrcamento("Compromisso")%></TD>
				<td></td>
				<TD><%=rstOrcamento("Favorecido")%></TD>
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=FormatNumber(rstOrcamento("vl_previsto"),2)%> &nbsp;</TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=rstOrcamento("banco")%>&nbsp;&nbsp;</TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=rstOrcamento("forma_pgto")%>&nbsp;&nbsp;</TD>
				<td></td>
				<TD><%=rstOrcamento("vencimento")%></TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=NomeMes(rstOrcamento("mes_final"))%>&nbsp;&nbsp;</TD>
				<td></td>
			</TR>			
<%	
			total = total + CDbl(rstOrcamento("vl_previsto"))
			ordem = ordem + 1
			rstOrcamento.MoveNext

		loop
		'Fecha loop
		
		rstOrcamento.Close
		set rstOrcamento=nothing
		
%>
		<TR align="center">
			<td></td>
			<TD colspan=5>&nbsp;T o t a l</TD>
			<td></td>
			<TD align="right">&nbsp;&nbsp;<%=FormatNumber(total,2)%> &nbsp;</TD>
			<td></td>
			<TD colspan=7>&nbsp</TD>
			<td></td>

		</TR>	
	</font>
	</TABLE>
	
<p><p>
	<center><input type = "submit" value = "  I M P R I M I R  " onClick="window.open('Imp_orcamento.asp?competencia=<%=competencia%>')"/></center>
	<center><input type = "submit" value = "  M E N U  " onClick="window.open('CtrlFin.asp','_self')"/></center>
	
	

</body>
</html>

<%

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