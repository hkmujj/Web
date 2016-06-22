<%@ Page Language="C#" ContentType="text/html" ResponseEncoding="gb2312" Debug="true" %>
<%@ Import Namespace="System.IO" %>
<%@ import namespace="System.Diagnostics" %>
<%@ Import Namespace="Microsoft.Win32" %>
<%@ Import Namespace="System.Collections" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Threading" %>
<%@ Import Namespace="System.Net.Sockets" %>
<%@ Import Namespace="System.Diagnostics" %>


<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Linq" %>
<%@ Import Namespace="System.Web" %>
<%@ Import Namespace="System.Web.Security" %>
<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="System.Web.UI.WebControls" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Xml.Linq" %>

 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>33333</title>
</head>
<body> 
<form id="Form4" runat="server">
 
</form>

<script runat="server">
//加载
protected void Page_Load(object sender, EventArgs e){

}
</script>
<%
	string str = "server='(local)';database='WebData';uid='sa';pwd='sa'";		//本地
	//string str = "server='qds157513275.my3w.com,1433';database='qds157513275_db';uid='qds157513275';pwd='313801120'"; 
	SqlConnection con = new SqlConnection(str); 
	con.Open();                                                                                               //打开连接
	string strsql = "select * from admin";                                                        //SQL查询语句 
	SqlCommand cmd = new SqlCommand(strsql, con);                    			//初始化Command对象 
	SqlDataReader rd = cmd.ExecuteReader();                                      //初始化DataReader对象

	string Astr="";
	while (rd.Read()){ 	
		Astr="<a href='?act=show&id="+ rd[0] +"'>"+ rd["userName"].ToString() +"</a>";
		Response.Write(rd[0].ToString() + "、" + Astr + "<hr>");                                         //通过索引获取列
	}
	string id=Request.QueryString["id"];
	Response.Write("id=" + id);
%>
</body>
</html>
