<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <title>
            历史成绩
    </title>
    <style>
      body {
  background-color: #f4f4f4;
}
.table1{
  background-color: #fff;
  margin: 0 auto;
  text-align: center;
  
}
    </style>
    
</head>
<body>

<p>
     set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "D:\lab\lab3\score.mdb"

sql="INSERT INTO score (s_name,s_score)"
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("fname") & "',"
sql=sql & "'" & Request.Form("score") & "')"

on error resume next
conn.Execute sql,recaffected
if err<>0 then
  Response.Write("No update permissions!")
 
end if
conn.close

<h2 align="center" font ><font color="red"> 提交成功</font></h2>
</p>
<p>
  <table width="50%" class="table1">
    <tr>
      <td>
        <h1>历史成绩</h1>
      </td>
    </tr>
    <tr><td>
<p>
<%
Dim conn, rs, sql

' 创建数据库连接对象
set conn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "D:\lab\lab3\score.mdb"

' SQL查询
sql = "SELECT s_name, s_score FROM score ORDER BY s_score DESC"

' 执行查询
rs.Open sql, conn

' 显示查询结果
Response.Write "<table border='1' width='100%' cellspacing='0' cellpadding='2'>"
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
<%


' 创建数据库连接对象
Set conn = Server.CreateObject("ADODB.Connection")
conn.Provider = "Microsoft.Jet.OLEDB.4.0"
conn.Open "D:\lab\lab3\score.mdb"

' DELETE语句
sql = "DELETE FROM score WHERE s_name = ' '"

' 执行DELETE语句
conn.Execute sql

' 关闭连接
conn.Close
Set conn = Nothing
%>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


</p></td></tr></table></p>
</body>
</html>