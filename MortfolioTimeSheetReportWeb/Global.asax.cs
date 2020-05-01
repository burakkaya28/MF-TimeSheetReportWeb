using System;
using System.Web;
using System.Web.Routing;
using log4net.Config;

namespace MortfolioTimeSheetReportWeb
{
    public class Global : HttpApplication
    {
        protected void Application_Start(object sender, EventArgs e)
        {
            Routing(RouteTable.Routes);
            XmlConfigurator.Configure();
        }

        private static void Routing(RouteCollection route)
        {
            route.MapPageRoute("Default", "", "~/Index.aspx");
        }
    }
}