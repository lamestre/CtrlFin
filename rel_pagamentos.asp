<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

	<!-- #include file = "conexao.asp" -->

<%
	mesRef		= Right(request("competencia"),2)
	ano			= Left(request("competencia"),4)
	mesDespesa 	= "15/" & mesRef & "/" & ano
	competencia = ano & mesRef
	
	set rstDespesa = Server.CreateObject ("ADODB.recordset")
	sqlDespesa = " SELECT * FROM tb_despesas AS d WHERE d.competencia = "&competencia&" AND dt_pgto IS NOT null AND vl_pago IS NOT NULL"
	rstDespesa.Open sqlDespesa, conn
	
	if rstDespesa.eof then
		response.write"<div align=center><font style=""font-family:verdana;font-size:9pt"" color=#FF0000<b>Ainda não foi paga nenhuma Despesa para esta competencia.</font></div>"
		response.write"<br>"
		response.write"<div align=center><font style=""font-family:verdana;font-size:9pt"" color=#3300CC><b><a href=""javascript:history.back()"">Voltar</a></b>"
		response.end
	end if
	
%>

<body>
<div id="tudo">

	<div id= "header">
		<h1>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h1>
	</div>
		
	<div id="conteudo">
		<hr>
			<center><h1>P A G A M E N T O S</h1>
			<%=NomeMes(mesDespesa)%>
		<hr>
	</div>

	
	<TABLE align="center" border="1">
		<tr align="center">
			<td></td>
			<td></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Compromisso&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Valor Documento&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Valor Pago&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Valor Residual&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Data do Pagamento&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Forma de Pgto&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Banco&nbsp;&nbsp;</B></td>
			<td></td>
		</tr>

<%
		'Cria o Record Set para o BD
		set rstDespesas = Server.CreateObject ("ADODB.recordset")
			
		sqlDespesas = 				" select * from tb_despesas	as d "
		sqlDespesas = sqlDespesas & "inner join tb_orcamento		as o on d.cod_compromisso	=	o.cod_compromisso "
		sqlDespesas = sqlDespesas & "left join tb_banco			as b on d.cod_banco			=	b.cod_banco "
		sqlDespesas = sqlDespesas &	"left join tb_forma_pgto 		as f on d.cod_forma_pgto	=	f.cod_forma_pgto "
		sqlDespesas = sqlDespesas &	"where competencia = " & competencia & " and (dt_pgto IS NOT NULL AND vl_pago IS NOT NULL) order by d.dt_pgto, o.vencimento, o.cod_banco asc "

		'Abre a conexão do Record Set com o BD
		rstDespesas.Open sqlDespesas, conn
		
		total = 0
		ordem = 1
		
		'Abre loop
		do while not rstDespesas.EOF
			if isnull(rstDespesas("vl_alternativo"))  then
				vlResidual = " - &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			elseif rstDespesas("vl_alternativo") = "0" then
				vlResidual = " - &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			else
				vlResidual = FormatNumber(rstDespesas("vl_alternativo"),2)
			end if
%>	
			<TR align="center">
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=ordem%> &nbsp;</TD>
				<td></td>
				<TD align="left">&nbsp;<%=rstDespesas("compromisso")%></TD>
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=FormatNumber(rstDespesas("vl_documento"),2)%> &nbsp;</TD>
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=FormatNumber(rstDespesas("vl_pago"),2)%> &nbsp;</TD>
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=vlResidual%> &nbsp;</TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=rstDespesas("dt_pgto")%>&nbsp;&nbsp;</TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=rstDespesas("forma_pgto")%>&nbsp;&nbsp;</TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=rstDespesas("banco")%>&nbsp;&nbsp;</TD>
				<td></td>
			</TR>			
<%	
			total = total + CDbl(rstDespesas("vl_documento"))
			ordem = ordem + 1
			rstDespesas.MoveNext

		loop
		'Fecha loop
		
		rstDespesas.Close
		set rstDespesas=nothing
		
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
	<center><input type = "submit" value = "  I M P R I M I R  " onClick="window.open('Imp_pagamentos.asp?competencia=<%=competencia%>')"/></center>
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