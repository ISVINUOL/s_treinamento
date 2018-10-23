<%@ LANGUAGE="VBSCRIPT" %>
<%OPTION EXPLICIT       %>
<%


%>
<html>
	<head>
		<title>tela8.asp</title>
	</head>
	<body>
		<font size="+1">
			<form id="cad_pessoa" method="POST" ACTION="tela9.asp">
				<table align="center">
					<tr>
						<td align="left"><b>Nome:</b></td>
						<td><input type="Text" name="txtNome"></td>
					</tr>
					<tr>
						<td align="left"><b>Senha:</b></td>
						<td><input type="password" name="txtSenha"></td>
					</tr>
					<tr>
						<td align="left"><b>Endereço:</b></td>
						<td><input type="Text" name="txtEndereco"></td>
					</tr>
					<tr>
						<td align="left"><b>Possui em casa:</b></td>
						<td>
							<input type="checkbox" name="chkEmcasa" Value="TV">TV
							<input type="checkbox" name="chkEmcasa" Value="Rádio"> Rádio
							<input type="checkbox" name="chkEmcasa" Value="Microondas"> Microondas
						</td>
					</tr>
					<tr>
						<td align="left"><b>Time Preferido:</b></td>
						<td>
							<select name="optTime" size="3">
								<option name="corinthians">Corinthians</option>
								<option name="flamengo">Flamengo</option>
								<option name="ambos">Ambos</option>
							</select> 
						</td>
					</tr>
					<tr>
						<td align="left"><b>Estado Cívil:</b></td>
						<td>
							<input type="radio" name="rdoEstado_civil" value="casado">Casado
							<input type="radio" name="rdoEstado_civil" value="solteiro">Solteiro
							<input type="radio" name="rdoEstado_civil" value="outro">Outro
						</td>
					</tr>
					<tr>
						<td align="left"><b>Observações:</b></td>
						<td>
							<TEXTAREA NAME="txtObs" COLS="50" ROWS="5"></TEXTAREA> 
						</td>
					</tr>
					<tr>
						<td align="left"><input type="submit" value="Botão OK"></td>
						<td align="right"><input type="Reset" value="Botão Limpar"></td>
					</tr>
				</table>
			</form>
		</font>
		<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
		<p align="right"> 
			<a href="C:\Users\Desenv\Desktop\Apostila\ExercÃ­cios HTML\Exercí­cio10.html">Copright.</a>
		</p>
	</body>
</html>