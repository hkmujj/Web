<%@ Page Language="C#" AutoEventWireup="true" CodeFile="1.cs" Inherits="_Default" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
</head>
<body>
<center>
<h2><font face="宋体">访问数据库的通用代码实例</font>
</h2>
</center>
<body>
    <form id="form1" runat="server">
    <div>

    <font face="宋体">
<p align="center">1.请输入相应数据库连接字符串</p>
<p align="center">
<asp:TextBox id="ConnStrTextBox" runat="server" Width="600"></asp:TextBox>
</p>
<p align="center">2.请输入相应SQL查询命令语句</p>
<p align="center">
<asp:TextBox id="SqlTextTextBox" runat="server" Width="600"></asp:TextBox>
</p>
<p align="center">3.请选择所连接的数据库类型</p>
<p align="center">
    <asp:DropDownList ID="DBDropDownList" runat="server" Width="204px">
        <asp:ListItem Selected="True">Access</asp:ListItem>
        <asp:ListItem>SQLServer</asp:ListItem>
        <asp:ListItem>Oracle</asp:ListItem>
        <asp:ListItem>DB2</asp:ListItem>
    </asp:DropDownList>
</p>
<p align="center">

<asp:Button ID="Button1" runat="server" onclick="Button1_Click"  Text="通用数据库连接代码测试" />

</p>
<p align="center">
<asp:Label id="lblMessage" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
</p>
    </form> 
    
<!--
1.请输入相应数据库连接字符串  输入这个是正确的
server='qds157513275.my3w.com,1433';database='qds157513275_db';uid='qds157513275';pwd='313801120'

server='127.0.0.1,1433';database='WebData';uid='sa';pwd='sa'

-->
