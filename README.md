# powershell_mssql_splunk
A powershell script which creates Splunk config files for MSSQL Server and SQLAgent logs.
Works with multiple instances and on clusters. Writes clustername to "host" parameter if available.

Example Splunk inputs.conf
```
[monitor://D:\SQLData\MSSQL11.MSSQLSERVER\MSSQL\Log\ERRORLOG]
disabled=false
index=i_splunk_appl_mssql
sourcetype=mssql_error
source=MSSQLSERVER
followTail=false
host=HOST1
[monitor://D:\SQLData\MSSQL11.MSSQLSERVER\MSSQL\Log\SQLAGENT.OUT]
disabled=false
index=i_splunk_appl_mssql
sourcetype=mssql_error
source=MSSQLSERVER(Agent)
followTail=false
host=HOST1
```

Example Splunk props.conf
```
[mssql_error]
SHOULD_LINEMERGE = true 
BREAK_ONLY_BEFORE_DATE = true 
MAX_TIMESTAMP_LOOKAHEAD = 22 
CHARSET = UTF-16LE
NO_BINARY_CHECK = true

```
