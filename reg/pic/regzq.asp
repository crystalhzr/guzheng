<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset="utf-8"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
 tmpno=request.form("txtsno")
 tmpname=request.form("txtsname")
 tmpsex=request.form("txtssex")
 tmpdept=request.form("txtdept")
 tmpclass=request.form("txtclass")
 'tmplove=request.form("txtlove")
 
 tmplove=""
 For I=1 to Request.form("txtlove").Count
	tmplove=tmplove&Request.form("txtlove")(I)&"," 
 Next
 
 m=len(tmplove)
 tmplove=Left(tmplove,m-1) 
 tmpemail=request.form("txtemail")

 'set cn=server.createobject("adodb.connection")
 'cn.open "driver={sql server};server=jcb;database=reg;uid=sa;pwd=123"
 
 set cn=server.createobject("adodb.connection")
 strconn="driver={microsoft access driver (*.mdb)};dbq="&server.mappath("reg.mdb")
 cn.open strconn

 
 set rs=server.createobject("adodb.recordset")
 selstr="select * from useinfo where sno='" & tmpno & "'"
 rs.open selstr,cn

 if not (rs.eof) then
   	response.write "<script language=JavaScript>" & "alert('该账号已被占用，请重新选择账号！');" & "history.back();" & "</script>" 
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
   
   strin= "insert into useinfo(sno,sname,sex,dept,class,love,email)"
   strin=strin&"values('"&tmpno&"','"&tmpname&"','"&tmpsex&"','"&tmpdept&"','"&tmpclass&"','"&tmplove&"','"&tmpemail&"')"
   cn.execute  strin
   response.write "<script language=JavaScript>" & "alert('注册成功！');" & "history.back();" & "</script>" 
 end if
%>