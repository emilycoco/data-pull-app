


<html>
<head>
<title>Survey Table</title>
<link rel="stylesheet"type="text/css" href="http://www.researchnowsurveys.com/survey/codefolder/stylesheet.css"/>


</head>

<body>



<div>

	<%

	
	dim mymailid
	dim mycountry
	
	mymailid = request("mailid")
	mycountry = request("panel")
	

	
	Dim eConnection, existRecordset
	Dim eSQL, eConnString
	





	

	eSQL= "SELECT count(mailcampaignid) as total from survey where mailcampaignid = " & mymailid
	
	eConnString="DSN=VOP" & mycountry & "UID=sa;PWD=fr87bi209"
	
	Set econnection = Server.CreateObject("ADODB.Connection")
	Set existrecordset = Server.CreateObject("ADODB.Recordset")
	econnection.Open eConnString
	existrecordset.Open eSQL, econnection
	
	
	
	myexist = existrecordset("total")
	existrecordset.close
If myexist = 0 then
	Response.write "<br/> <br/> <h1 class='headstyle'> No data in survey table, or survey table does not 	exist </h1>"
	Response.end
End If

If myexist > 0 then
	

	



	
	
	Dim cConnection, CountRecordset
	Dim cSQL, cConnString
	
	

	
	
	response.write "<table class='table table-bordered table-striped'>"
	
	
	Response.write "<th class='text-info'>Aggregate Data</th>"



	Dim oConnection, openrecordset
	Dim oSQL, oConnString
	Dim myopen
	
	oSQL= "SELECT COUNT(DISTINCT cpr_MAILCAMPAIGNID) AS STATUS FROM closedprojects WHERE cpr_MAILCAMPAIGNID = " & mymailid
	
	oConnString="DSN=VOP" & mycountry & "UID=sa;PWD=fr87bi209"
	
	Set oconnection = Server.CreateObject("ADODB.Connection")
	Set openrecordset = Server.CreateObject("ADODB.Recordset")
	oconnection.Open oConnString
	openrecordset.Open oSQL, oconnection
	
	
	
	myopen = openrecordset("STATUS")
	
		If myopen = 1 then
			mystatus = "This project is <span class='text-error'>closed</span>"
		End If

		If myopen = 0 then
			mystatus = "This project is open"
		End If
	

	Response.Write "<tr><td class='cell'>" & mystatus & "</td>"
	
	openrecordset.close
	

	cSQL= "SELECT COUNT(DISTINCT SUBSID) AS TOTAL FROM SURVEY WHERE MAILCAMPAIGNID = " & mymailid
	
	cConnString="DSN=VOP" & mycountry & "UID=sa;PWD=fr87bi209"
	
	Set cconnection = Server.CreateObject("ADODB.Connection")
	Set Countrecordset = Server.CreateObject("ADODB.Recordset")
	cconnection.Open cConnString
	Countrecordset.Open cSQL, cconnection
	
	
	
	mycount = countrecordset("total")

	Response.Write "<tr><td class='cell'>Entries: " & mycount & "</td>"
	
	Countrecordset.close
	
	dim onesql
	dim completesrecordset
	oneSQL= "SELECT COUNT(DISTINCT SUBSID) AS TOTAL FROM SURVEY WHERE MAILCAMPAIGNID = " & mymailid & " and status = 1"
	Set completesrecordset = Server.CreateObject("ADODB.Recordset")
	completesrecordset.Open oneSQL, cconnection
	
	mycompletes = completesrecordset("total")
	
	Response.Write "<tr><td class='cell'>Completes: " & mycompletes & "</td>"
	completesrecordset.close
	
	dim twosql
	dim sosrecordset
	twoSQL= "SELECT COUNT(DISTINCT SUBSID) AS TOTAL FROM SURVEY WHERE MAILCAMPAIGNID = " & mymailid & " and status = 2"
	Set sosrecordset = Server.CreateObject("ADODB.Recordset")
	sosrecordset.Open twoSQL, cconnection
	
	mysos = sosrecordset("total")
	
	Response.Write "<tr><td class='cell'>Screen Outs: " & mysos & "</td>"
	
	sosrecordset.close
	
	dim threesql
	dim qfsrecordset
	threeSQL= "SELECT COUNT(DISTINCT SUBSID) AS TOTAL FROM SURVEY WHERE MAILCAMPAIGNID = " & mymailid & " and status = 3"
	Set qfsrecordset = Server.CreateObject("ADODB.Recordset")
	qfsrecordset.Open threeSQL, cconnection
	
	myqfs = qfsrecordset("total")
	
	Response.Write "<tr><td class='cell'>Quota Fulls: " & myqfs & "</td>"
	
	qfsrecordset.close
	
	
	cconnection.close
	response.write "</table>"
	
	
	
	%>
</div>
<div>
    <form class='form-search'>
    <input type="text" class="input-medium search-query">
    <button type="submit" class="btn btn-info">Search</button>
    </form>
</div>
</br>

<div>

<%
Dim Connection, TableRecordset
Dim sSQL, sConnString 

sSQL= "SELECT SUBSID, STATUS, EXECID, DT AS DATE FROM SURVEY WHERE MAILCAMPAIGNID = " & mymailid & " ORDER BY SUBSID"
	
	sConnString="DSN=VOP" & mycountry & "UID=sa;PWD=fr87bi209"
		Set connection = Server.CreateObject("ADODB.Connection")
	Set tablerecordset = Server.CreateObject("ADODB.Recordset")
	connection.Open sConnString
	tablerecordset.Open sSQL, connection


	response.write "<table class='table table-bordered table-striped table-hover table-condensed'>"
	Response.Write "<tbody class='table-hover'><tr>" 
	For Each objField in tablerecordset.Fields
 	Response.Write "<th class='text-info'>" & objField.Name & "</th>"
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
 	response.write "</tbody></table>"
End if


tablerecordset.close
connection.close
%>

</div>


</body>

</html> 