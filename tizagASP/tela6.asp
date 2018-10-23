<%@ LANGUAGE="VBSCRIPT" %>
<%OPTION EXPLICIT 


%>
<html>
    <head>
        <title>tela6.asp</title>
    </head>
    <body>
        <form method="POST" id="percentual" ACTION="tela7.asp">
            <table>
                <tr>
                    <td>
                        <label for="valor">Valor Total:</label>
                    </td>
                    <td>
                        <input type="text" name="valor" size="3">
                    </td>
                </tr>
                <tr>
                    <td>
                        <label for="valor">Valor para cálculo do percentual:</label>
                    </td>
                    <td>
                        <input type="text" name="valor2" size="3">
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