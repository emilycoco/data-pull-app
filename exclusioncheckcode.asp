


<html>
<head>
<title>Exclusions</title>
<link rel="stylesheet"type="text/css" href="http://www.researchnowsurveys.com/survey/codefolder/stylesheet.css"/>


</head>

<body>



<div>
	<%

	
	dim mymailid
	dim mycountry
	dim mydate1
	dim mydate2
	
	mymailid = request("mailid")
	mycountry = request("panel")
	mydate1 = request("year") & "-" & request("month") & "-" & request("day") & " 00:00:00"
	mydate2 = request("year2") & "-" & request("month2") & "-" & request("day2") & " 24:00:00"


	dim mailidarray

	mailidarray = split(mymailid," ")




Dim Connection, tableRecordset
Dim  strsql, sConnString 

strsql= "SELECT DISTINCT SUBSID, MAILCAMPAIGNID AS MAILID, DT AS COMPLETE_DATE FROM SURVEY WHERE MAILCAMPAIGNID in ("
'separates string of mailids into separate variables in an array
For i = 0 To UBound(mailidarray) - 1
strsql = strsql  & mailidarray(i) & ", "
Next
strsql = strsql  & mailidarray(UBound(mailidarray)) & ")"

	strSql = strSql & " and status = 1 and dt >= '" & mydate1 & "' and dt <= '" & mydate2 & "'"
	sConnString="DSN=VOP" & mycountry & "UID=sa;PWD=fr87bi209"
	Set connection = Server.CreateObject("ADODB.Connection")
	Set tablerecordset = Server.CreateObject("ADODB.Recordset")
	connection.Open sConnString
	tablerecordset.Open strsql, connection





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



tablerecordset.close
connection.close
%>

</div>

</body>

</html> 