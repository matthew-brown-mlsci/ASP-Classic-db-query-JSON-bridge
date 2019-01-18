<%
' This bit of ASP classic accepts HTTP POST requests, looks for the "sql" parameter, and
' executes the SQL statement against a static defined database.  Databases connections are setup
' on the windows server as ODBC system DSN's.  This code accesses those DSN's, and returns
' a JSON compliant return that can be parsed via regular javascript JSON methods.
'
' It's a bit unpolished, and could use some work, but definitely works! 
'
' Example DSN used here: iferror_mlab_prod
' 
' Example http POST request:
' http://example.com/sql_bridge.asp&sql=SELECT column_name, column_phonenumber, column_established FROM example_DSN
'
' (Note, you can't try this in a web browser directly, this code is setup for POST requests, not GET requests)
'
' example response
'
' {
' 	"sql_query": "SELECT column_name, column_phonenumber, column_established FROM example_DSN"
'	"sql_response_column_headers": {
'		"column_name":"string",
'		"column_phonenumber":"number",
'		"column_established":"datetime",
'		}
'	"sql_response": [
'			{
'			"column_name":"John Doe",
'			"column_phonenumber":"4144654040",
'			"column_established":"2015-01-20T16:26:19.500Z"
'			},
'			{
'			"column_name":"Steve Noone",
'			"column_phonenumber":"1234567890",
'			"column_established":"2015-01-20T16:29:19.500Z"
'			}
'		]
'	"sql_stats": {
'		"total_rows":"2",
'		"total_columns":"3",
'		"query_time_in_seconds":"1",
'		"error_message":""
'		}
'	}
'

Call Response.AddHeader("Access-Control-Allow-Origin", "http://example.com")
Call Response.AddHeader("Access-Control-Allow-Credentials", "true")

'These constants are set so we open a read-only copy of the database pointer 'rs'
Const adOpenStatic = 3
Const adUseClient = 3
Const adLockOptimistic = 3

' Some other variables that should only be changed on the server side
Dim conn, rs
Dim qry, connectstr
Dim db_name, db_username, db_userpassword
Dim db_server
Dim sql_query 

' Check and see if we recieved a sql query, if we don't, send appropriate json and exit
sql_query = Request.Form("sql")
'Response.Write(sql_query)

' Clean up not-nice characters in the sql query
'sql_query = Replace(sql_query, """", "\""")

' If query is really short, error immediately
If Len(sql_query) < 1 Then
	Response.Write("{" & vbCrlf)
	Response.Write("""sql_query"":""" & sql_query & """," & vbCrlf)
	Response.Write("""sql_response_column_headers"":{}" & vbCrlf)
	Response.Write("""sql_response"": {}" & vbCrlf)
	Response.Write("""sql_stats"": {" & vbCrlf)
	Response.Write("	""error_message"":""Error, SQL Query too short""" & vbCrlf)
	Response.Write("	}" & vbCrlf)
	Response.End()
End If

' Connect to Database, create recordset, send the sql query
connectstr = "iferror_mlab_prod"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connectstr
set rs=Server.CreateObject("ADODB.recordset")
rs.Open sql_query, conn, adOpenStatic, adLockOptimistic

'Response.Write("<p>rs.Properties.RecordCount = " & rs.Fields(4).Name)
'Response.Write("<br>")

' First check if sql is a SELECT statement or something else
'		- initially, we'll support select, insert and update statements
'		- descr perhaps in the future...
sql_statement_type = Split(sql_query, " ")
'Response.Write("sql_statement_type(0) " & sql_statement_type(0) & vBCrlf)
'Response.Write("sql_statement_type(1) " & sql_statement_type(1) & vBCrlf)
'Response.Write(StrComp(sql_statement_type(0), "UDATE", 1) & vBCrlf)
If (StrComp(sql_statement_type(0), "SELECT", 1) = 0) Then
	Response.Write("{" & vbCrlf)
'	echo back the sql query sent
	Response.Write("""sql_query"":""" & sql_query & """," & vbCrlf)
	Response.Write("""sql_response_column_headers"": {" & vbCrlf)
	field_headers = ""
	For each line in rs.Fields
		field_headers = field_headers & "	""" & line.Name & """:""" & line.Type & """," & vbCrlf
	next
	' Trim the last comma (and crlf chars)
	field_headers = Left(field_headers, Len(field_headers) - 3) & vbCrlf
	Response.Write(field_headers)
	Response.Write("	}," & vbCrlf)
	
	' now the section where the tables get sent over
	' //todo: account for special characters / dates
	Response.Write("""sql_response"": [" & vbCrlf)
	safe_line_value = ""
	do until rs.EOF
		sql_row = "	{"
		for each line in rs.Fields
			If IsNull(line.Value) Then
				safe_line_value = ""
			Else
				safe_line_value = line.Value
				safe_line_value = Replace(safe_line_value, ChrW(0022), ("\" & ChrW(0022)))
				safe_line_value = Replace(safe_line_value, ChrW(0092), ("\" & ChrW(0092)))
				safe_line_value = Replace(safe_line_value, ChrW(0047), ("\" & ChrW(0047)))
				safe_line_value = Replace(safe_line_value, ChrW(0008), "\" & ChrW(0008))
				safe_line_value = Replace(safe_line_value, ChrW(0034), "\" & ChrW(0034))
				safe_line_value = Replace(safe_line_value, ChrW(0012), "\r")
				safe_line_value = Replace(safe_line_value, ChrW(0010), "\r")
				safe_line_value = Replace(safe_line_value, ChrW(0013), "\r")
				safe_line_value = Replace(safe_line_value, ChrW(0009), "\" & "t")	
			End If
			sql_row = sql_row & "		""" & line.Name & """:""" & safe_line_value & """," & vbCrlf
		next
		sql_row = Left(sql_row, Len(sql_row) - 3) & vbCrlf
		Response.Write(sql_row)
		rs.MoveNext
		If rs.EOF Then
			Response.Write("		}" & vbCrlf)
		Else
			Response.Write("		}," & vbCrlf)
		End If
	loop
	Response.Write("	]" & vbCrlf)
	Response.Write("}" & vbCrlf)
	rs.close
End If

' UPDATE statements
If (StrComp(sql_statement_type(0), "UPDATE", 1) = 0) Then
	Response.Write("{" & vbCrlf)
	Response.Write("""sql_query"":""" & sql_query & """," & vbCrlf)
	Response.Write("""sql_response_column_headers"": {" & vbCrlf)
	Response.Write("	}," & vbCrlf)
	Response.Write("""sql_response"": [" & vbCrlf)
	Response.Write("	]" & vbCrlf)
	Response.Write("}" & vbCrlf)
End If

' UPDATE statements
If (StrComp(sql_statement_type(0), "INSERT", 1) = 0) Then
	Response.Write("{" & vbCrlf)
	Response.Write("""sql_query"":""" & sql_query & """," & vbCrlf)
	Response.Write("""sql_response_column_headers"": {" & vbCrlf)
	Response.Write("	}," & vbCrlf)
	Response.Write("""sql_response"": [" & vbCrlf)
	Response.Write("	]" & vbCrlf)
	Response.Write("}" & vbCrlf)
End If

'Response.Write("<p>" & Request.QueryString("sql") & "<br>")
' change this to error out in JSON (START HERE!)
sub DisplayErrorInfo()
  Response.Write("{" & vbCrlf)
  Response.Write("""sql_query"":""" & sql_query & """," & vbCrlf)
  Response.Write("""sql_response_column_headers"":{}" & vbCrlf)
  Response.Write("""sql_response"": {}" & vbCrlf)
  Response.Write("""sql_stats"": {" & vbCrlf)
  Response.Write("	""error_message"":""")
  For Each errorObject In rs.ActiveConnection.Errors
    Response.Write "Description :" & errorObject.Description
    Response.Write "   Number:" & Hex(errorObject.Number) & "" & vbCrlf
  Next
  Response.Write("	}" & vbCrlf)
  Response.End()
End Sub


conn.close
%>