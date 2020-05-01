using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using log4net;

namespace MortfolioTimeSheetReportWeb
{
    public class App
    {
        //Database Connection String
        public string MflodbConString = ConfigurationManager.AppSettings["MFLO_CON_STRING"];
        public string IzinAppConString = ConfigurationManager.AppSettings["IZINAPP_CON_STRING"];

        //Getting current application path
        public readonly string MainFilePath = HttpContext.Current.Request.PhysicalApplicationPath;

        //Log4net main variable
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public void ExportToExcel(HtmlGenericControl tableBody, HttpResponse response)
        {
            response.Clear();
            response.Buffer = true;
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("content-disposition", "attachment;filename=OPTiiMTimesheetReport_" + DateTime.Now.ToString("yyyy-MM-ddhh:mm:ss") + ".xls");
            response.Write("<style type=\"text/css\"> td, tr, tbody {border: 1px solid black; }");
            response.Write("</style>");
            var sw = new StringWriter();
            response.Charset = "";
            response.ContentEncoding = Encoding.GetEncoding("windows-1254");
            var htw = new HtmlTextWriter(sw);
            tableBody.RenderControl(htw);
            response.Write(" <meta http-equiv='Content-Type' content='text/html; charset=windows-1254' />" + sw);
            response.End();
        }

        public string GetUserMail(string username)
        {
            string value = null;
            try
            {
                using (var con = new SqlConnection(MflodbConString))
                {
                    
                    using (var cmd = new SqlCommand("select EMAIL from USERACCOUNT where USER_NAME = '"+username+"'", con))
                    {
                        cmd.CommandType = CommandType.Text;
                        con.Open();
                        value = Convert.ToString(cmd.ExecuteScalar());
                        con.Close();
                    }
                }
            }
            catch (SqlException ex)
            {
                Log.Error(ex);
            }
            return value;
        }

        public string CheckLeaveIsExist(string leaveStartDate, string userEmail)
        {
            string value = null;
            try
            {
                using (var con = new SqlConnection(IzinAppConString))
                {
                    
                    using (var cmd = new SqlCommand("select count(*) from LeaveRequests lr join Users u on u.UserId = lr.UserId where Day >= 5 and Status = 1 and '"+leaveStartDate+"' between lr.StartDate and lr.EndDate and DATEADD(DAY, 4, '"+leaveStartDate+"') <= EndDate and u.Email = '"+userEmail+"'", con))
                    {
                        cmd.CommandType = CommandType.Text;
                        con.Open();
                        value = Convert.ToString(cmd.ExecuteScalar());
                        con.Close();
                    }
                }
            }
            catch (SqlException ex)
            {
                Log.Error(ex);
            }

            return value;
        }
    }
}