<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset="utf-8"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
 tmpname=request.form("txtsname")
 tmpsex=request.form("txtssex")
 tmpdept=request.form("txtdept")

 

 'set cn=server.createobject("adodb.connection")
 'cn.open "driver={sql server};server=jcb;database=reg;uid=sa;pwd=123"
 
 set cn=server.createobject("adodb.connection")
 strconn="driver={microsoft access driver (*.mdb)};dbq="&server.mappath("reg.mdb")
 cn.open strconn

 
 set rs=server.createobject("adodb.recordset")
 selstr="select * from useinfo where sname='" & tmpname & "'"
 rs.open selstr,cn

 if not (rs.eof) then
   	response.write "<script language=JavaScript>" & "alert('该用户名已被占用，请重新选择用户！');" & "history.back();" & "</script>" 
    rs.close
    set rs=nothing
    cn.close
    set cn=nothing
 else
   'set cn=server.createobject("adodb.connection")
   'cn.open "driver={sql server};server=jcb;database=reg;uid=sa;pwd=123"
   
   set cn=server.createobject("adodb.connection")
   strconn="driver={microsoft access driver (*.mdb)};dbq="&server.mappath("reg.mdb")
   cn.open strconn
   
   strin= "insert into useinfo(sname,sex,dept,email)"
   strin=strin&"values('"&tmpname&"','"&tmpsex&"','"&tmpdept&"','"&tmpemail&"')"
   cn.execute  strin
   response.write "<script language=JavaScript>" & "alert('注册成功！');" & "history.back();" & "</script>" 
 end if
%>