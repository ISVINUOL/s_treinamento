<%@ LANGUAGE="VBSCRIPT" %>
<%OPTION EXPLICIT       %>
<%

dim valTo, valPe, valCal

valTo = request.form("valor")
valPe = request.form("valor2")

if valTo <> "" and valPe <>"" then    
    valCal = (valPe/valTo ) * 100
end if

%>
<html>
    <head>
        <title>tela7.asp</title>
    </head>
    <body>
        <p align="center"> 
            <font face="arial" size="5">
                O valor total digitado foi: <%=valTo%><br>
	            O valor para cálculo do percentual foi: <%=valPe%><br>
	            O percentual calculado é: <%=valCal%> %
            </font>
            <br>
        </p> 
        <input type="button" id="voltar" value="voltar" onclick="teste();">
    </body>

    <script LANGUAGE="javascript"> 
        function teste(){
            window.location.assign("tela6.asp")
        }
    </script>
    
</html>