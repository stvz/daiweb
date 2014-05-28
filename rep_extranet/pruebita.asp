<% 
scr="hola"
width = 40
height = 30
%>
<html>
	<body>
		<p>Pagina de prueba</p>
	</body>
</html>
<script language=javascript>
    funtion popUpWindow(){
        var x = window.open(‘<%=src%>’,’pictureWindow’,’width=<%=width%>,height=<%height%>’);
    }
</script>