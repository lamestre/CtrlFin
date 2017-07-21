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
	sqlDespesa = " SELECT * FROM tb_despesas AS d WHERE d.competencia = " & competencia
	rstDespesa.Open sqlDespesa, conn
	
	if rstDespesa.eof then
		response.write"<div align=center><font style=""font-family:verdana;font-size:9pt"" color=#FF0000<b>Ainda não foi gerado controle de Despesa para esta competencia.</font></div>"
		response.write"<br>"
		response.write"<div align=center><font style=""font-family:verdana;font-size:9pt"" color=#3300CC><b><a href=""javascript:history.back()"">Voltar</a></b>"
		response.end
	end if
	
%>

<body>
<div id="tudo">

		
	<div id="conteudo">
		<center><h2>. &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; .</h2>
			<center><B>D E S P E S A S &nbsp;&nbsp;-&nbsp;&nbsp; <%=NomeMes(mesDespesa)%></B>
		<hr>
	</div>

	
	<TABLE align="center" border="1">
		<tr align="center">
			<td></td>
			<td></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Compromisso&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Vlr. Doc.&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Vlr. Alt.&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Banco&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Meio&nbsp;&nbsp;</B></td>
			<td></td>
			<td><B>&nbsp;&nbsp;Venc.&nbsp;&nbsp;</B></td>
			<td></td>
		</tr>

<%
		'Cria o Record Set para o BD
		set rstDespesas = Server.CreateObject ("ADODB.recordset")
			
		sqlDespesas = 				" select * from tb_despesas	as d "
		sqlDespesas = sqlDespesas & "inner join tb_orcamento		as o on d.cod_compromisso	=	o.cod_compromisso "
		sqlDespesas = sqlDespesas & "inner join tb_banco			as b on o.cod_banco			=	b.cod_banco "
		sqlDespesas = sqlDespesas &	"inner join tb_forma_pgto 		as f on o.cod_forma_pgto	=	f.cod_forma_pgto "
		sqlDespesas = sqlDespesas &	"where competencia = " & competencia & " and (dt_pgto is null or vl_pago is null) order by o.vencimento, o.cod_banco asc "

		'Abre a conexão do Record Set com o BD
		rstDespesas.Open sqlDespesas, conn
		
		total = 0
		ordem = 1
		
		'Abre loop
		do while not rstDespesas.EOF
			if isnull(rstDespesas("vl_alternativo")) then
				vlAlternativo = " - &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			else
				vlAlternativo = FormatNumber(rstDespesas("vl_alternativo"),2)
			end if
			
			codDespesa = rstDespesas("cod_despesa")

%>	
			<TR align="center">
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=ordem%> &nbsp;</TD>
				<td></td>
				<TD align="left">&nbsp;<%=rstDespesas("compromisso_abrv")%></TD>
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=FormatNumber(rstDespesas("vl_documento"),2)%> &nbsp;</TD>
				<td></td>
				<TD align="right">&nbsp;&nbsp;<%=vlAlternativo%> &nbsp;</TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=rstDespesas("banco_abrv")%>&nbsp;&nbsp;</TD>
				<td></td>
				<TD>&nbsp;&nbsp;<%=rstDespesas("forma_pgto_abrv")%>&nbsp;&nbsp;</TD>
				<td></td>
				<TD><%=rstDespesas("vencimento")%></TD>
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
			<TD colspan=3>&nbsp;T o t a l</TD>
			<td></td>
			<TD align="right">&nbsp;&nbsp;<%=FormatNumber(total,2)%> &nbsp;</TD>
			<td></td>
			<TD colspan=7>&nbsp</TD>
			<td></td>

		</TR>	
	</font>
	</TABLE>

<script>
	print();
	window.close();
</script>

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