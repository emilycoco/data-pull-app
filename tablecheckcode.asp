


<html>
<head>
<title>Temp Table</title>
<link rel="stylesheet"type="text/css" href="http://www.researchnowsurveys.com/survey/codefolder/stylesheet.css"/>
</head>

<body>



<div>
	<%
	
response.end	
	dim mytemptable
	dim mycountry
	
	mytemptable = request("temptable")
	mycountry = request("panel")
	


	
	Dim eConnection, existRecordset
	Dim eSQL, eConnString
	





	

	eSQL= "SELECT count(tabname) as total from SYSCAT.tables where TABNAME= '" & mytemptable & "'"
	
	eConnString="DSN=VOP" & mycountry & "UID=sa;PWD=fr87bi209"
	
	Set econnection = Server.CreateObject("ADODB.Connection")
	Set existrecordset = Server.CreateObject("ADODB.Recordset")
	econnection.Open eConnString
	existrecordset.Open eSQL, econnection
	
	
	
	myexist = existrecordset("total")
	existrecordset.close

If myexist = 0 then

	Response.Write " <br/> <br/> <h1 class='headstyle'>No data in table, or table does not 	exist</h1>"
	Response.end
End If

If myexist = 1 then
	

		Dim cConnection, CountRecordset
	Dim cSQL, cConnString
	
	

	
	
	response.write "<table class='table'>"
	
	
	Response.write "<th class='cellhead'>Aggregate Data</th>"




	

	cSQL= "SELECT COUNT(DISTINCT SUBSID) AS TOTAL FROM " & mytemptable
	
	cConnString="DSN=VOP" & mycountry & "UID=sa;PWD=fr87bi209"
	
	Set cconnection = Server.CreateObject("ADODB.Connection")
	Set Countrecordset = Server.CreateObject("ADODB.Recordset")
	cconnection.Open cConnString
	Countrecordset.Open cSQL, cconnection
	
	
	
	mycount = countrecordset("total")

	Response.Write "<tr><td class='cell'>Entries: " & mycount & "</td>"
	
	Countrecordset.close
	

	response.write "</table>"
	
	
	
	%>
</div>

</br>

<div>

<p>

</p>

</br>

<p>
<%
Dim Connection, TableRecordset
Dim sSQL, sConnString 

sSQL= "SELECT * from " & mytemptable & " ORDER BY SUBSID"
	
	sConnString="DSN=VOP" & mycountry & "UID=sa;PWD=fr87bi209"
		Set connection = Server.CreateObject("ADODB.Connection")
	Set tablerecordset = Server.CreateObject("ADODB.Recordset")
	connection.Open sConnString
	tablerecordset.Open sSQL, connection



	response.write "<table class='table'>"
	Response.Write "<tr>" 
	For Each objField in tablerecordset.Fields
 	Response.Write "<th class='cellhead'>" & objField.Name & "</th>"
	Next
	Response.write "</tr>"

	Do While Not tablerecordset.EOF
	Response.Write "<tr>" 
	For Each objField in tablerecordset.Fields
	Response.Write "<td class='cell'>" & objField & "</td>"
	Next
	Response.write "</tr>"
	tableRecordset.MoveNext
	Loop
 	response.write "</table>"
End if



tablerecordset.close
connection.close
%>
</p>
</div>

</body>

</html> 