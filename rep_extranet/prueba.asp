<%Server.ScriptTimeout=15000000
	
	info =  "<div style='display:table; background-color:red; border:2px solid pink; table-layout:fixed; '>" & _
			"<div style='background-color:white; border:2px solid blue; display:table-row;  '>" & _
				"<div style='background-color:lime; border:2px solid black; display:table-cell; '>TH 1</div>" & _
				"<div style='background-color:yellow; border:2px solid black; display:table-cell; '>TD 1</div>" & _
				"<div style='background-color:yellow; border:2px solid black; display:table-cell; '>TD 2</div>" & _
			"</div>" & _
			"<div style='background-color:white; border:2px solid blue; display:table-row; '>" & _
				"<div style='background-color:lime; border:2px solid black; display:table-cell; '>TH 2</div>" & _
				"<div style='background-color:yellow; border:2px solid black; display:table-cell; '>TD 3</div>" & _
				"<div style='background-color:yellow; border:2px solid black; display:table-cell; '>TD 4</div>" & _
			"</div>" & _
			"<div style='background-color:white; border:2px solid blue; display:table-row; '>" & _
			   "<div style='background-color:lime; border:2px solid black; display:table-cell; '>TH 3</div>" & _
			   "<div style='background-color:yellow; border:2px solid black; display:table-cell; '>TD 5</div>" & _
			   "<div style='background-color:yellow; border:2px solid black; display:table-cell; '>TD 6</div>" & _
			"</div>" & _
			"</div>"
	response.write(info)
	response.end()
	html = info & "<br>"

%>

<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE CARTERA .... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>