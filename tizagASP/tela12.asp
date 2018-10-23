<% @LANGUAGE = "VBSCRIPT"   %>
<%OPTION EXPLICIT           %>
<%
dim objLog, Retorno
dim Empresa_Id, Usuario_Internet_id, vntConecta_DB
dim Banco_Dados, Nome_Produto, msg_erro, Msg_Ret, msg_aviso
dim idioma
dim strUser, strPWD, strQuery, strConexao, objConn, rs, x, a 
dim CONTADOR, crud
dim tam_arr_teste, flg_segue
dim gl_txtID, gl_ID_Usuario, gl_txtNome, gl_txtSenha, gl_txtEndereco, gl_txtBens, gl_txtTime, gl_txtEstadoCV, gl_txtObs
dim vntId, vntNome, vntSenha, vntEndereco, vntTime, vntEstadoCV, vntObs
dim vntchkTV, vntchkRADIO, vntchkMICROONDAS


' SESSION("gl_txtNome")
' SESSION("gl_txtSenha")
' SESSION("gl_txtEndereco")
' SESSION("gl_txtBens")
' SESSION("gl_txtTime")
' SESSION("gl_txtEstadoCV")
' SESSION("gl_txtObs")

vntId = SESSION("gl_ID_Usuario")
vntNome = SESSION("gl_txtNome")
vntSenha = SESSION("gl_txtSenha")
vntEndereco = SESSION("gl_txtEndereco")
vntTime = SESSION("gl_txtTime")
vntEstadoCV = SESSION("gl_txtEstadoCV")
vntObs = SESSION("gl_txtObs")

a = split(SESSION("gl_txtBens"),",")

for each x in a
    select case trim(x)
        case "Microondas"
            vntchkMICROONDAS =  "checked"
        case "TV"
            vntchkTV = "checked"
        case "Rádio"
            vntchkRADIO =  "checked"
    end select
next
crud = ""
function operacao(crud)
    select case crud
        case "limpar"
            SESSION("gl_ID_Usuario") = ""
            SESSION("gl_txtNome") = ""
            SESSION("gl_txtSenha") = ""
            SESSION("gl_txtEndereco") = ""
            SESSION("gl_txtBens") = ""
            SESSION("gl_txtTime") = ""
            SESSION("gl_txtEstadoCV") = ""
            SESSION("gl_txtObs") = ""
        case "consultar"
            Response.redirect("consulta.asp")
    end select
end function

%>
<html>
	<head><title>Cadastro</title></head>
    <!--#include FILE = "CONFIG2.ASP"-->
	<script type="text/javascript" language="javascript" src="funcoes.js"></script>
	<body>
            <font size="+1">
                    <table width=600 ALIGN = CENTER>
                        <%if session("msg_erro") <> "" then%>
                            <tr><td align=center><%=GL_FMER(session("msg_erro"))%></td></tr>
                        <%end if%>
                        <%if session("msg_aviso") <> "" then%>
                            <tr><td align=center><%=GL_FMAV(session("msg_aviso"))%></td></tr>
                        <%end if%>
                        <%session("msg_aviso") = ""%>
                        <%session("msg_erro") = ""%>
                    </table>
                    <form id="cad_pessoa" method="POST" ACTION="">
                        <table align="center">
                            <tr>
                                <td align="left"><b>Nome:</b></td>
                                <td>
                                <input type="text" name "txtID" value="<%=vntId%>">
                                <input type="Text" name="txtNome" maxlength="50" value="<%=vntNome%>"></td>
                            </tr>
                            <tr>
                                <td align="left"><b>Senha:</b></td>
                                <td><input type="password" name="txtSenha" maxlength="10" value="<%=vntSenha%>"></td>
                            </tr>
                            <tr>
                                <td align="left"><b>Endereço:</b></td>
                                <td><input type="Text" name="txtEndereco"maxlength="200" value="<%=vntEndereco%>"></td>
                            </tr>
                            <tr>
                                <td align="left"><b>Possui em casa:</b></td>
                                <td>
                                    <input type="checkbox" name="chkBens" Value="TV" 
                                        <% if vntchkTV="checked" then%>checked<%end if%>>TV
                                    <input type="checkbox" name="chkBens" Value="Rádio"
                                        <% if vntchkRADIO="checked" then%>checked<%end if%>> Rádio
                                    <input type="checkbox" name="chkBens" Value="Microondas"
                                        <% if vntchkMICROONDAS="checked" then%>checked<%end if%>> Microondas
                                </td>
                            </tr>
                            <tr>
                                <td align="left"><b>Time Preferido:</b></td>
                                <td>
                                    <select name="optTime" size="3">
                                        <option name="Corinthians" <% if vntTime="Corinthians" then%>selected<%end if%>>Corinthians</option>
                                        <option name="Flamengo" <% if vntTime="Flamengo" then%>selected<%end if%>>Flamengo</option>
                                        <option name="Ambos" <% if vntTime="Ambos" then%>selected<%end if%>>Ambos</option>
                                    </select> 
                                </td>
                            </tr>
                            <tr>
                                <td align="left"><b>Estado Cívil:</b></td>
                                <td>
                                    <input type="radio" name="rdoEstado_civil" value="Casado"<%if vntEstadoCV="Casado" then%>checked<%end if%>>Casado
                                    <input type="radio" name="rdoEstado_civil" value="Solteiro"<%if vntEstadoCV="Solteiro" then%>checked<%end if%>>Solteiro
                                    <input type="radio" name="rdoEstado_civil" value="Outro"<%if vntEstadoCV="Outro" then%>checked<%end if%>>Outro
                                </td>
                            </tr>
                            <tr>
                                <td align="left"><b>Observações:</b></td>
                                <td>
                                    <TEXTAREA NAME="txtObs" COLS="50" ROWS="5" ><%=vntObs%></TEXTAREA> 
                                </td>
                            </tr>
                            <tr>
                                <td ></td>
                                <td align="right"><a href="SalvarDados.asp" /><img src="imagens\save.ico" width="40px" height="40px">
                                <a href="consulta.asp" /><img src="imagens\search.ico" width="40px" height="40px">
                                <a href="SalvarDados.asp" /><img src="imagens\edit.ico" width="40px" height="40px">
                                <a href="consulta.asp" /><img src="imagens\remove.ico" width="40px" height="40px"></td>
                                &nbsp;
                                <td align="right"><a href="<%=operacao("limpar")%>" /><img src="imagens\clear.ico" width="40px" height="40px"></td>
                            </tr>
                        </table>
                    </form>
                </font>
                <br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
                <p align="right"><a href="C:\Users\Desenv\Desktop\Apostila\ExercÃ­cios HTML\Exercí­cio10.html">Copright.</a>
                </p>
	    </body>
</html>