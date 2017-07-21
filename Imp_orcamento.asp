<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

<%
	mesRef		= Right(request("competencia"),2)
	ano			= Left(request("competencia"),4)
	mesOrcado 	= "15/" & mesRef & "/" & ano
'	response.write mesRef&"*"&ano&"*"&mesOrcado
%>

<body>
<div id="tudo">
		
	<div id="conteudo">
		<center><h2>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h2>
			<center><B>O R Ç A M E N T O &nbsp;&nbsp;-&nbsp;&nbsp; <%=NomeMes(mesOrcado)%></B>
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
			<td><B>&nbsp;&nbsp;Valor&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Banco&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Meio&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Venc.&nbsp;&nbsp;</B></td>
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
				<TD align="left">&nbsp;<%=rstOrcamento("Compromisso_abrv")%></TD>
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=FormatNumber(rstOrcamento("vl_previsto"),2)%> &nbsp;</TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=rstOrcamento("banco_abrv")%>&nbsp;&nbsp;</TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=rstOrcamento("forma_pgto_abrv")%>&nbsp;&nbsp;</TD>
				<td></td>
				<TD><%=rstOrcamento("vencimento")%></TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=NumMes(rstOrcamento("mes_final"))%>&nbsp;&nbsp;</TD>
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
			<TD colspan=3>&nbsp;T o t a l</TD>
			<td></td>
			<TD align="right">&nbsp;&nbsp;<%=FormatNumber(total,2)%> &nbsp;</TD>
			<td></td>
			<TD colspan=7>&nbsp</TD>
			<td></td>

		</TR>	
	</font>
	</TABLE>
	
<p><p>

<script>
	print();
	window.close();
</script>

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

Function NumMes(data)
	if isnull(data) then
		mes = "-"
	else
		mes = Month(data)
		Select Case mes
			Case 1 : mes = "01"
			Case 2 : mes = "02"
			Case 3 : mes = "03"
			Case 4 : mes = "04"
			Case 5 : mes = "05"
			Case 6 : mes = "06"
			Case 7 : mes = "07"
			Case 8 : mes = "08"
			Case 9 : mes = "09"
			'Case 10 : mes = "Outubro"
			'Case 11 : mes = "Novembro"
			'Case 12 : mes = "Dezembro"
		End Select
		mes = mes & "/" & Year(data) 
	end if
  NumMes =  mes
End Function
%>