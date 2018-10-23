<%@ LANGUAGE="VBSCRIPT" %>
<%OPTION EXPLICIT
dim tamanho
dim titulo 
dim cor_tela
dim font

titulo = "Primeiro Asp"
cor_tela = "red"
font = "arial"

%>
<html>
    <title><%=titulo%></title>

    <body bgcolor=<%=cor_tela%>>

        <%for tamanho = 1 to 7%>
            <font face=<%=font%> size=<%=tamanho%>>Sou feliz por ser como sou, por ter o que tenho, mas amanhã quero ser ainda mais e para isso luto diariamente.</font> <br>
            <font face=<%=font%> size=<%=tamanho%>>Amar a vida é amar seus amigos, pois sem eles nada mais faz sentido no seu dia a dia.</font> <br><br>
        <%next%>

    </body>
</html>