using System;
using System.Web.UI;

namespace MortfolioTimeSheetReportWeb
{
    public partial class Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void ExcelBtn_OnClick(object sender, EventArgs e)
        {
            new App().ExportToExcel(tableBody, Response);
        }
    }
}