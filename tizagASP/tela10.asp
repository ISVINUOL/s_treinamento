<%@ LANGUAGE="VBSCRIPT" %>
<%OPTION EXPLICIT 

dim gl_asp_origem, gl_Valor, gl_Valor2


%>
<html>
    <head>
        <title>tela10.asp</title>
    </head>
    <body>
        <form method="POST" id="percentual" ACTION="tela11.asp">
            <table>
                <tr>
                    <td>
                        <label for="valor">Valor Total:</label>
                    </td>
                    <td>
                        <input type="text" name="valor" size="3" value="<%=session("gl_valTo")%>">
                    </td>
                </tr>
                <tr>
                    <td>
                        <label for="valor">Valor para cálculo do percentual:</label>
                    </td>
                    <td>
                        <input type="text" name="valor2" size="3" value="<%=session("gl_valPe")%>">
                    </td
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <input type="submit" value="Calcular"><br>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>