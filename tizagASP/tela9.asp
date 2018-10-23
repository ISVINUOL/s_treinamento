<%@ LANGUAGE="VBSCRIPT" %>
<%OPTION EXPLICIT       %>
<%


dim nome, senha, endereco
dim emcasa, time, est_civil
dim obs

nome = request.form("txtNome")
endereco = request.form("txtEndereco")
senha = request.form("txtSenha")
emcasa = request.form("chkEmcasa")
time = request.form("optTime")
est_civil = request.form("rdoEstado_civil")
obs = request.form("txtObs")

%>

<html>
    <head><title>tela9.asp</title>
    </head>
    <body>
        <div align="center">
            <form id="recebe_val" method="POST" action="tela8.asp">
                <table>
                    <tr>
                        <th align="left">
                            Nome:
                        </th>
                        <td>
                            <p><%=nome%></p>
                        </td>
                    </tr>
                    <tr>
                        <th align="left">
                            Senha:
                        </th>
                        <td>
                            <p><%=senha%></p>
                        </td>
                    </tr>    
                    <tr>
                        <th align="left">
                            Endereço:
                        </th>
                        <td>
                            <p><%=endereco%></p>
                        </td>
                    </tr>
                    <tr>
                        <th align="left">
                            Possui em casa:
                        </th>
                        <td>
                            <p><%=emcasa%></p>
                        </td>
                    </tr>
                    <tr>
                        <th align="left">
                            Time Preferido:
                        </th>
                        <td>
                            <p><%=time%></p>
                        </td>
                    </tr>
                    <tr>
                        <th align="left">
                            Estado Cívil:
                        </th>
                        <td>
                            <p><%=est_civil%></p>
                        </td>
                    </tr>
                    <tr>
                        <th align="left">
                            Observações:
                        </th>
                        <td>
                            <p><%=obs%></p>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <input type="submit" value="Voltar">
                        </td>
                    </tr>
                </table>
            </form>
        </div>
    </body>
</html>
