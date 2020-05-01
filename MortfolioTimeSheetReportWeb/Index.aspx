<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="MortfolioTimeSheetReportWeb.Default" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Diagnostics" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="MortfolioTimeSheetReportWeb" %>

<!DOCTYPE html>
<html lang="en">
<head>
	<title>OPTiiM Time Sheet Report</title>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
    <!--===============================================================================================-->
	<link rel="icon" type="image/png" href="images/icons/favicon.ico"/>
    <!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="vendor/bootstrap/css/bootstrap.min.css">
    <!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="fonts/font-awesome-4.7.0/css/font-awesome.min.css">
    <!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="vendor/animate/animate.css">
	<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="vendor/select2/select2.min.css">
	<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="vendor/perfect-scrollbar/perfect-scrollbar.css">
	<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="css/util.css">
	<link rel="stylesheet" type="text/css" href="css/main.css">
	<!--===============================================================================================-->
</head>
<body>
	<form runat="server">

	<div class="limiter" style="width: auto;" runat="server">
		<div>
			<h4 style="margin-left: 90px; padding-top: 18px; margin-bottom: 5px; width: 30%; float: left">OPTiiM Zaman Çizelgesi Raporu</h4>
            <asp:Button ID="ExcelBtn" runat="server" Text="Excele Çıkart" style="float: right; margin-right: 7.2%; margin-top: 8px; cursor: pointer; font-size: 14px;" class="btn btn-primary" OnClick="ExcelBtn_OnClick"/>
		</div>
		<div id="tableBody" class="container-table100" style="padding: 0; min-height: 0;" runat="server">
			<div class="wrap-table100">
				<div class="table100 ver1 m-b-110" style="height: 550px; margin-bottom: 0;">
					<div class="table100-head">
						<table>
							<thead>
								<tr class="row100 head">
									<th class="cell100 column1">Ad Soyad</th>
									<th class="cell100 column2">Takım</th>
									<th class="cell100 column3">Periyod</th>
									<th class="cell100 column4">Hafta Numarası</th>
									<th class="cell100 column5">Durum</th>
								</tr>
							</thead>
						</table>
					</div>

					<div class="table100-body js-pscroll" style="max-height: 488px;">
						<table>
							<tbody>
                                <%
                                    var script = File.ReadAllText(new App().MainFilePath + @"\timesheet.sql");

                                    var connection = new SqlConnection(new App().MflodbConString);
                                    connection.Open();
                                    var sqlDr = new SqlCommand(script, connection).ExecuteReader();
                                    while (sqlDr.Read())
                                    {
                                        var userName = sqlDr["USER_NAME"].ToString();
                                        var userEmail = new App().GetUserMail(userName);
                                        var fullName = sqlDr["FULL_NAME"].ToString();
                                        var team = sqlDr["TEAM"].ToString();
                                        var startDateStr = sqlDr["START_DATE"].ToString().Replace(" 12:00:00 AM", "").Replace("00:00:00", "");
                                        var startDate = DateTime.Parse(startDateStr).ToString("yyyy-MM-dd");
                                        var period = sqlDr["START_DATE"].ToString().Replace(" 12:00:00 AM", "").Replace("00:00:00", "") + " - " + sqlDr["END_DATE"].ToString().Replace(" 12:00:00 AM", "").Replace("00:00:00", "");
                                        var weekNumber = sqlDr["WEEK_NUMBER"].ToString();
                                        var state = sqlDr["STATE"].ToString();

                                        var leaveExist = new App().CheckLeaveIsExist(startDate, userEmail);
                                        
                                        //Log for Report
                                        Debug.WriteLine(userEmail + " => " + startDate + " => " + leaveExist + " => " + weekNumber);

                                        if (leaveExist == "0")
                                        {
                                            %>
							                    <tr class="row100 body">
							                        <td class="cell100 column1"><%=fullName %></td>
							                        <td class="cell100 column2"><%=team %></td>
							                        <td class="cell100 column3"><%=period %>  </td>
							                        <td class="cell100 column4"><%=weekNumber %></td>
							                        <td class="cell100 column5"><%=state %></td>
							                    </tr>
						                    <%
                                        }
                                    }
                                    sqlDr.Close();
                                    connection.Close();
                                %>
							</tbody>
						</table>
					</div>
				</div>
			</div>
		</div>
	</div>
	</form>


<!--===============================================================================================-->	
	<script src="vendor/jquery/jquery-3.2.1.min.js"></script>
<!--===============================================================================================-->
	<script src="vendor/bootstrap/js/popper.js"></script>
	<script src="vendor/bootstrap/js/bootstrap.min.js"></script>
<!--===============================================================================================-->
	<script src="vendor/select2/select2.min.js"></script>
<!--===============================================================================================-->
	<script src="vendor/perfect-scrollbar/perfect-scrollbar.min.js"></script>
	<script>
		$('.js-pscroll').each(function(){
			var ps = new PerfectScrollbar(this);

			$(window).on('resize', function(){
				ps.update();
			})
		});
			
		
	</script>
<!--===============================================================================================-->
	<script src="js/main.js"></script>

</body>
</html>