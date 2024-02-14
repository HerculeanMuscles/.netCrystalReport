using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Xml.Linq;

using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Web;
using System.Data.OleDb;

public partial class _Default : System.Web.UI.Page 
{

    private const string ExcelFilePath = "C:\\Users\\Fahri\\source\\repos\\Wazaran1\\Contracts2\\LayoffContract.xlsx";
    private const string ExcelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\";";
    private const string crystalreportFile1 = "LayoffReport.rpt";
    ReportDocument report = new ReportDocument();

    protected void Page_Init(object sender, EventArgs e)
    {
        this.Init += new System.EventHandler(this.Page_Init);


        if (!IsPostBack)
        {
            CrystalReportsLoad();
            CrystalReportViewer1.ReportSource = report;

        }
        else
        {
            ReportDocument doc = (ReportDocument)Session["ReportDocument"];
            CrystalReportViewer1.ReportSource = doc;
        }
    }

    protected void CrystalReportsLoad()
    {
        try
        {
            report.Load(Server.MapPath(crystalreportFile1));

            DataTable testData = GetTestDataFromExcel(); // retrieve data here

            report.SetDataSource(testData);

            CrystalReportViewer1.ReportSource = report;
            Session["ReportDocument"] = report;

        }
        catch (Exception ex) { throw ex; }
    }


    private DataTable GetTestDataFromExcel()
    {
        using (OleDbConnection connection = new OleDbConnection(ExcelConnectionString))
        {
            connection.Open();

            // Retrieve all data from the sheet
            string query = "SELECT * FROM [Sheet1$]";
            using (OleDbCommand cmd = new OleDbCommand(query, connection))
            {
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    connection.Close();
                    return dataTable;
                }
            }
        }
    }
}