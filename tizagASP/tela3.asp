<%@ LANGUAGE="VBSCRIPT" %>
<%OPTION EXPLICIT

dim x, y, z
dim result

function calc_perc(valor, valor2)
    y = valor
    x = valor2

    if y <> "" and x <> "" then
        ' if z >= 50 then
        '     response.redirect("tela4.asp")
        '     z = z & "%"
        ' end if
        z = x * (y/100)

        select case true
            case x = 50
                z=0
                result = "O valor é igual a 50%."
            case x > 50
                result=""
                response.redirect("tela4.asp")
            case x < 50
                result=""
                response.redirect("tela5.asp")
        end select
        
        if x <> 50 then
            result = ""
        end if
    end if
end function


%>
<html>
    <head>
        <title>tela3.asp</title>
    </head>
    <body>
        <form method="POST" id="percentual">
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
                        <label id="valor_res" for="valor_res"><font face="arial" size="5"><%=result%></font></p></label>
                    </td>
                </tr>
            </table>
        </form>
        <script type="text/javascript" language="javascript">
            function LBLSHOW(){
                alert("O percentual é de 50%.");
            }
        </script>
    </body>
</html>