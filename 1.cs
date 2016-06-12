using System;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

using System.Text;		//储存文本的
using System.Data.SqlClient;
using System.Configuration;
using System.Collections.Generic;


/*
Access 数据库代码
Provider=Microsoft.Jet.OLEDB.4.0; Data source=E:\E盘\WEB网站\至前网站\Admin\Data\Data.mdb
SQLServer 代码
Data Source=qds157513275.my3w.com,1433;Initial Catalog=qds157513275_db;User ID=qds157513275;Password=313801120

DB2 是IBM公司的
*/

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
		//Response.Write("IsPostBack=" + IsPostBack + "<hr>");
        //通用数据库连接代码，这里以连接Access数据库为测试示例
        if (!IsPostBack)
        {
           ConnStrTextBox.Text = "Provider=Microsoft.Jet.OLEDB.4.0; Data source=" + Server.MapPath("/Admin/Data/oldData.mdb");
           SqlTextTextBox.Text = "Select COUNT(*) From admin";
           lblMessage.Text = "";
        }
    }
    protected void Button1_Click(object sender, EventArgs e)
    {

        //定义数据库连接字符串
        string MyConnectionString = this.ConnStrTextBox.Text;
        //定义查询操作的SQL语句
        string MySQL = this.SqlTextTextBox.Text;
        //定义所要连接的数据库类型为Access
        string MyType = this.DBDropDownList.SelectedValue;
        System.Data.IDbConnection MyConnection = null;
        // 根据数据库类型，创建相应的 Connection 对象
        switch (MyType)
        {
            //选择的数据库类型为“SQLServer”，创建SqlConnection类数据库连接对象
            case "SQLServer":
                MyConnection = new System.Data.SqlClient.SqlConnection(MyConnectionString);
                break;
            case "Oracle":
                //MyConnection = new System.Data.OracleClient.OracleConnection(MyConnectionString);
                break;
            //选择的数据库类型为“Access”，创建OleDbConnection类数据库连接对象
            case "Access":
                MyConnection = new System.Data.OleDb.OleDbConnection(MyConnectionString);
                break;
            //选择的数据库类型为“DB2”，创建OleDbConnection类数据库连接对象
            case "DB2":
                MyConnection = new System.Data.Odbc.OdbcConnection(MyConnectionString);
                break;
			//默认就是Access
            default:
                MyConnection = new System.Data.OleDb.OleDbConnection(MyConnectionString);
                break;
        }
        Execute(MyConnection, MySQL);
    }
    public void Execute(System.Data.IDbConnection MyConnection, string strquery)
    {
        //使用 CreateCommand() 方法生成 Command 对象
        System.Data.IDbCommand MyCommand = MyConnection.CreateCommand();
        //执行定义的SQL查询语句
        MyCommand.CommandText = strquery;
        try
        {
            //打开数据库连接
            MyConnection.Open();
            //定义查询的结果信息
            String MyInfo = "测试连接成功！符合查询要求的记录共有：" + MyCommand.ExecuteScalar().ToString() + "条！";
            //输出查询结果信息
            lblMessage.Text = MyInfo;
			
			DataSet ds=new DataSet();//这里是你的数据，我就不写了
 
			
			string str = "server='qds157513275.my3w.com,1433';database='qds157513275_db';uid='qds157513275';pwd='313801120'"; 
            SqlConnection con = new SqlConnection(str); 
            con.Open();                                                                                               //打开连接
            string strsql = "select * from admin";                                                        //SQL查询语句 
            SqlCommand cmd = new SqlCommand(strsql, con);                    //初始化Command对象 
            SqlDataReader rd = cmd.ExecuteReader();                                      //初始化DataReader对象 
            while (rd.Read()) 
            { 
                Response.Write(rd[0].ToString() + "、" + rd["username"].ToString() + "<hr>");                                         //通过索引获取列 
            }
			
			

			Response.Write("打开数据库成功！" + ds.Tables["product1sdfasd"]);
			
			
        	strsql = "select count(1) from xy_webdomain";
			SqlCommand cmd = new SqlCommand(strsql, con);
        	Response.Write("共" + cmd.ExecuteScalar() + "条");
			
        }
        catch (Exception ex)
        {
            //输出错误异常
            Response.Write(ex.ToString());
        }
        finally
        {
            //关闭数据库连接
            MyConnection.Close();
        }
    }
}

