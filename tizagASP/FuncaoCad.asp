<% @LANGUAGE = "VBSCRIPT"   %>
<%OPTION EXPLICIT           %>
<%
dim objLog, Retorno
dim Empresa_Id, Usuario_Internet_id, vntConecta_DB
dim Banco_Dados, Nome_Produto, Msg_Erro, Msg_Ret
dim idioma
dim strUser, strPWD, strQuery, strConexao, objConn, rs, x
dim CONTADOR
dim tam_arr_teste


idioma = ""
Session("grupo_usuario_id") = 4
empresa_id = 1
'Ang�lica 26/11/2003 - N�o � mais usado o idioma em ingl�s, estava dando problema na Rodorei
Session("idioma") = "P"
'
SESSION("NOME_BANCO") = "SQLSERVER"
SESSION("NOME_PRODUTO") = "NOVO"
'SESSION("GL_CONECTA_DB") = "DSN=DSNECARGO;UID=SISTEMA;PWD=SYSUSER;APP=WEB"
SESSION("GL_CONECTA_DB_NET") = "Data Source=192.168.1.2; User ID=SISTEMA; Password=SYSUSER;Initial Catalog=NOVO; Application Name=NOVO"
SESSION("GL_ASP_ORIGEM") = "INI.ASP"
SESSION("TEMP_FLAG_GRUPO_HABILITADO") = ""


vntConecta_DB = "192.168.1.2" 'colocar aqui a localiza��o de sua base de dados 
Banco_Dados ="NOVO" 'colocar aqui o nome do banco 
strUser = "SISTEMA" 'colocar aqui o nome do usu�rio 
strPWD = "SYSUSER" 'colocar aqui a senha 

'Geramos a query SQL que ir� acessar os dados na base de dados 
'Conforme altera��o 1 
strQuery = "SELECT * FROM EMPLOYEE" 

'Montamos a string de conex�o 
strConexao = "Provider=SQLOLEDB.1;SERVER=" & vntConecta_DB
strConexao = strConexao & "; DATABASE=" & Banco_Dados
strConexao = strConexao & ";UID="& strUser
strConexao = strConexao & " ;PWD="& strPWD

set objConn = server.CreateObject("ADODB.Connection") 
objConn.open strConexao
set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open strQuery, strConexao
Response.Write(rs.STATE)
objConn.close



set objConn = nothing

%>
<html>
	<head>
		<title>Cadastro</title>
		<script type="text/javascript" language="javascript" src="funcoes.js"></script>
	</head>

	<body>
		<table>			
			<tr>
				<td align="right">
					<%=GL_FTCA("Ponto Opera��o:")%>
				</td>
				<td>
					<input type="text" id="operacao">
				</td>
			</tr>
			<tr>
				<td align="right">
					<%=GL_FTPG("Tipo:")%>
				</td>
				<td>
					<input type="text" id="tipo_Doc">
				</td>
			</tr>
			<tr>		
			</tr>
		</table>
		<form>
			<input type="button" id="consultar" value="Pesquisar" onclick="Msg()">
		</form>
	</body>
</html>