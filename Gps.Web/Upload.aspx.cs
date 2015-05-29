using System;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

namespace Gps.Web
{
    public class Upload
    {
        protected void visualizzaTab(string nomefile)
        {
            
            //string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Server.MapPath("..\\UploadACU\\" + fuExcel.FileName) + ";Extended Properties=Excel 8.0;";
            //System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(strConn);
            //string sql = "SELECT Users, Serial_AcuConnect, LTRIM(Code_AcuConnect) as Code_AcuConnect, LTRIM(Key_AcuConnect) as Key_AcuConnect, Serial_RunTime, LTRIM(Code_RunTime) as Code_RunTime, LTRIM(Key_RunTime) as Key_RunTime,LTRIM(Upgrade) as Upgrade FROM [Sheet1$]";
            //System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(sql, conn);
            //try
            //{
            //    conn.Open();
            //    System.Data.OleDb.OleDbDataReader rd = cmd.ExecuteReader();
            //    dgExcel.Dispose();
            //    dgExcel.DataSource = rd;
            //    dgExcel.DataBind();
            //    dgExcel.Visible = true;
            //    btnAnnulla.Visible = true;
            //    btnConferma.Visible = true;
            //    conn.Close();
            //}
            //catch (Exception eccezione)
            //{
            //    lblErrori.Text = Functions.ProcessError(eccezione);
            //    conn.Close();
            //}
        }
    }
}