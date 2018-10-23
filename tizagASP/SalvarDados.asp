<% @LANGUAGE = "VBSCRIPT"   %>
<%OPTION EXPLICIT           %>
<%
dim invntID, invntNome, invntSenha, invntEndereco, invntTime, invntEstadoCV, invntObs, invntBens, objConn
dim gl_txtID, gl_txtNome, gl_txtSenha, gl_txtEndereco, gl_txtBens, gl_txtTime, gl_txtEstadoCV, gl_txtObs, msg_erro, msg_aviso, Retorno, outvntMsg_Ret

    ' SESSION(gl_txtNome) = request.form("txtNome")
    ' SESSION(gl_txtSenha) = request.form("txtSenha")
    ' SESSION(gl_txtEndereco) = request.form("txtEndereco")
    ' SESSION(gl_txtTime) = request.form("optTime")
    ' SESSION(gl_txtBens) = request.form("chkBens")
    ' SESSION(gl_txtEstadoCV) = request.form("rdoEstado_civil")
    ' SESSION(gl_txtObs) = request.form("txtObs")

    Retorno = ""
    invntID = request.form("txtID")
    invntNome = request.form("txtNome")
    invntSenha = request.form("txtSenha")
    invntEndereco = request.form("txtEndereco")
    invntTime = request.form("optTime")
    invntBens = request.form("chkBens")
    invntEstadoCV = request.form("rdoEstado_civil")
    invntObs = request.form("txtObs")

    SESSION("gl_txtID") = invntID
    SESSION("gl_txtNome") = invntNome
    SESSION("gl_txtSenha") = invntSenha
    SESSION("gl_txtEndereco") = invntEndereco
    SESSION("gl_txtTime") = invntTime
    SESSION("gl_txtBens") = invntBens
    SESSION("gl_txtEstadoCV") = invntEstadoCV
    SESSION("gl_txtObs") = invntObs

    Set objConn = NOTHING
    Set objConn = CreateObject("Conexao_Base.Conect")
    
    objConn.gConectar
    if invntID <> "" then 
        Retorno = objConn.gAtualiza_Cadastro_Usuario(invntID,invntNome,invntSenha,invntEndereco,invntBens,invntTime,invntEstadoCV,invntObs,outvntMsg_Ret)
    else
        'Retorno = objConn.gSalva_Cadastro_Usuario(invntNome,invntSenha,invntEndereco,invntBens,invntTime,invntEstadoCV,invntObs,outvntMsg_Ret)
    end if

    if RETORNO <> 0 THEN
        SESSION("msg_erro") = " " & outvntMsg_Ret
        objConn.gDesconectar
		set objConn = NOTHING
		RESPONSE.REDIRECT "tela12.asp"
	else
        SESSION("msg_aviso") = "Cadastro Salvo com Sucesso!!! - " & outvntMsg_Ret
        objConn.gDesconectar
        set objConn = NOTHING
        RESPONSE.REDIRECT "tela12.asp"
    end if
%>
<html>
    <head><title>Salvar Dados</title></head>
    <body>
        <table>
            <tr><%=invntNome%></tr>
            <tr><%=invntSenha%></tr>
            <tr><%=invntEndereco%></tr>
        </table>
    </body>
</html>