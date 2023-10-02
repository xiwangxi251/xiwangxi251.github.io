<%
Dim conn, rs, sql

' 创建数据库连接对象
set conn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "D:\lab\lab3\score.mdb"

' SQL查询
sql = "SELECT s_name, s_score FROM score"

' 执行查询
rs.Open sql, conn

' 显示查询结果
Response.Write "<table border='1' width='50%' cellspacing='0' cellpadding='2'>"
Response.Write "<tr><th>Name</th><th>Score</th></tr>"

Do While Not rs.EOF
    Response.Write "<tr>"
    Response.Write "<td>" & rs("s_name") & "</td>"
    Response.Write "<td>" & rs("s_score") & "</td>"
    Response.Write "</tr>"
    rs.MoveNext
Loop

Response.Write "</table>"

' 清理并关闭连接
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
%>
