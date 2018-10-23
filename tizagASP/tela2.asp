<%@ LANGUAGE="VBSCRIPT" %>
<%OPTION EXPLICIT
dim x, y, z


function calc_perc(valor, valor2)
    y = valor
    x = valor2

    if y <> "" and x <> "" then
        z = x * (y/100)
        z = z & "%"
    end if
end function


%>
<html>
    <head><title>tela2.asp</title></head>
    <body>
        <form method="POST" id="percentual" ACTION="tela2.asp">
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
                        <label for="valor">Parcela:</label>
                    </td>
                    <td>
                        <input type="text" name="valor2" size="3">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <input type="submit" value="Calcular" onclick="<%=calc_perc(request.form("valor"),request.form("valor2"))%>"><br>
                    </td>
                </tr>
                <tr>
                    <td>
                        Resultado:
                    </td>
                    <td>
                        <label for = "valor_res"><%=z%></label>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>