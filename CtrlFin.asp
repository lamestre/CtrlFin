<%
Option Explicit
Dim conn, competencia
%>
<html>
<head>
	<title> CtrlFin </title>
	<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>


	<!-- #include file = "conexao.asp" -->

	<Script language="JavaScript" type="text/javascript">

		function SetaComp() {
//				alert(document.frmCtrlFin.comp.value + " * " + document.frmCtrlFin.competencia.value);
//				document.frmCtrlFin.comp.value=document.frmCtrlFin.competencia.value;
				document.frmCtrlFin.submit();
		}

		
		function TrocaComp() {
			var comp = document.getElementById("competencia").value;
			document.frmCtrlFin.comp.value=comp;
			document.frmCtrlFin.submit();
		}
		</Script>

<%

	Dim comp
	
'	Session("competencia") = year(Date()) & right("0" & month(Date()),2)
	competencia = request.form("comp")
	if isNull(competencia) or competencia = "" then
		competencia = year(Date()) & right("0" & month(Date()),2)
	end if
	
'	comp = Session("competencia")

%>
	
<body <%if request.form("comp")="" then response.write("OnLoad=""SetaComp()""")%> >
<div id="tudo">

	<div id= "header">
		<h1> . &nbsp; : &nbsp;&nbsp; Controle Financeiro &nbsp;&nbsp; : &nbsp; . </h1>
	</div>
		<form name="frmCtrlFin"  method="post">
		
	<div id="conteudo">
		<div id="menu">
			<ul><hr>
			<h2>
				<li><a href="Rel_orcamento.asp?competencia=<%=request.form("comp")%>" title="Lista de Compromissos">Or&ccedil;amento</a></li>
				<li><a href="Rel_despesas.asp?competencia=<%=request.form("comp")%>" title="Lista de Despesas">Despesas </a></li>
				<li><a href="Rel_pagamentos.asp?competencia=<%=request.form("comp")%>" title="Lista de Pagamentos Realizados">Pagamentos</a></li>
				<li><a href="Tabelas.asp?competencia=<%=request.form("comp")%>" title="Mantençao de Tabelas">Tabelas</a></li>
				
				<input type=hidden name=comp  value=<%=competencia%>>
				
				<TABLE align="center" border="1">
					<tr>
						<td></td>
						<td><B>&nbsp;&nbsp;Competencia: &nbsp;&nbsp;</B></td>
						<td></td>
						<td><B>&nbsp;&nbsp;<select align=right id=competencia OnChange="TrocaComp()" style="text-align:right">
										<option value="201705" <%if competencia="201705" then response.write("selected")%>>05/2017</option>
										<option value="201706" <%if competencia="201706" then response.write("selected")%>>06/2017</option>
										<option value="201707" <%if competencia="201707" then response.write("selected")%>>07/2017</option>
										<option value="201708" <%if competencia="201708" then response.write("selected")%>>08/2017</option>
										<option value="201709" <%if competencia="201709" then response.write("selected")%>>09/2017</option>
										<option value="201710" <%if competencia="201710" then response.write("selected")%>>10/2017</option>
										<option value="201711" <%if competencia="201711" then response.write("selected")%>>11/2017</option>
										<option value="201712" <%if competencia="201712" then response.write("selected")%>>12/2017</option>
									</select>&nbsp;&nbsp;</B></td>
						<td></td>
					</tr>
				</TABLE>
				</form>
			</h2>
			</ul><hr>
		</div>
	</div>
</div>

<div id="footer">
	<img src="imgs/foot.png">
</div>

</body>
</html>
