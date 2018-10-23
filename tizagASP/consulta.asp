<% @LANGUAGE = "VBSCRIPT"   %>
<%OPTION EXPLICIT           %>
<%
dim opc(5)
dim i
dim vntConecta_DB, Banco_Dados, strUser, strPWD, strQuery, objConn, rs, strConexao, Retorno, Qtde, arrDescr, arrID, CONTADOR
dim flag_ID_USUARIO, flag_Nome_Usuario, flag_Senha_Usuario, flag_Endereco_Usuario, flag_Bens_Usuario, flag_Time_Usuario, flag_Estado_Civil_Usuario, flag_Observacoes, outvntmsg_ret
dim gl_ID_Usuario,gl_txtNome, gl_txtSenha, gl_txtEndereco, gl_txtBens, gl_txtTime, gl_txtEstadoCV, gl_txtObs

Set objConn = CreateObject("Conexao_Base.Conect")
objConn.gConectar

Retorno = objConn.gComboConsulta(Qtde, arrDescr, arrID, outvntmsg_ret)

set objConn = nothing

function retorna_data()
    retorna_data = date
end function

function monta_session()
    
    Set objConn = CreateObject("Conexao_Base.Conect")
    objConn.gConectar
    flag_ID_USUARIO = Request.Form("optCad")

    Retorno = objConn.gConsulta_Cadastro_Usuario(flag_ID_USUARIO, flag_Nome_Usuario, flag_Senha_Usuario, flag_Endereco_Usuario, flag_Bens_Usuario, flag_Time_Usuario, flag_Estado_Civil_Usuario, flag_Observacoes, outvntmsg_ret)
    SESSION("gl_ID_Usuario") = flag_ID_USUARIO
    SESSION("gl_txtNome") = flag_Nome_Usuario
    SESSION("gl_txtSenha") = flag_Senha_Usuario
    SESSION("gl_txtEndereco") = flag_Endereco_Usuario
    SESSION("gl_txtBens") = flag_Bens_Usuario
    SESSION("gl_txtTime") = flag_Time_Usuario
    SESSION("gl_txtEstadoCV") = flag_Estado_Civil_Usuario
    SESSION("gl_txtObs") = flag_Observacoes

    set objConn = nothing
    if flag_Nome_Usuario <> "" then
        Response.redirect("tela12.asp")
    end if
end function

' function limpa_campos()
'     Qtde = ""
'     arrID = ""
'     arrDescr = ""
'     flag_ID_USUARIO = ""
'     flag_Nome_Usuario = ""
'     flag_Senha_Usuario = ""
'     flag_Endereco_Usuario = ""
'     flag_Bens_Usuario = ""
'     flag_Time_Usuario = ""
'     flag_Estado_Civil_Usuario = ""
'     flag_Observacoes = ""
' end function
%>
<html>
    <head>
        <title>Consulta Cadastro</title>
        <script LANGUAGE="javascript"> 
            function teste(){
                alert("<%=monta_session%>")
            }
        </script>
    </head>
    <body>
        <form id="consulta" METHOD="POST" action="">
            <table>
                <tr>
                    <td>
                        Nome:
                    </td>
                    <td>
                        <select id="optCad" name="optCad" size="">
                            <% FOR CONTADOR = 0 TO Qtde -1 %>
                                <OPTION VALUE="<%=arrID(CONTADOR)%>"><%=arrDescr(CONTADOR)%>
                            <% NEXT %>
                        </select>
                    </td>
                    <td>
                        <input type="submit" value="Ver Registro" onclick="<%=monta_session%>">
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>