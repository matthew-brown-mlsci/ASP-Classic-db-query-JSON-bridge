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