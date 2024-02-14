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
using System.Diagnostics;

public partial class _Default : System.Web.UI.Page 
{

    private const string ExcelFilePath = "C:\\Users\\Fahri\\source\\repos\\Wazaran1\\Contracts2\\LayoffContract.xlsx";
    private const string ExcelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\";";
    private const string LayoffReport = "LayoffReport.rpt"; private const string OfferReport = "OfferReport.rpt";
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
            report.Load(Server.MapPath(LayoffReport));

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

    protected void btnGenerateReport_Click(object sender, EventArgs e)
    {
        // Create a DataSet
        DataSet dataSet = new DataSet();

        // Create a DataTable within the DataSet
        DataTable dataTable = new DataTable("Layoff");

        // Define columns in the DataTable (replace with your actual column names and types)
        dataTable.Columns.Add("SenderName", typeof(string));
        dataTable.Columns.Add("SenderAddress", typeof(string));
        dataTable.Columns.Add("SenderStateZip", typeof(string));
        dataTable.Columns.Add("txtEmail", typeof(string));
        dataTable.Columns.Add("DateToday", typeof(DateTime));
        dataTable.Columns.Add("LastEmploymentDate", typeof(DateTime));
        dataTable.Columns.Add("ReceiverName", typeof(string));
        dataTable.Columns.Add("Title", typeof(string));
        dataTable.Columns.Add("OrganizationName", typeof(string));
        dataTable.Columns.Add("ReceiverAddress", typeof(string));
        dataTable.Columns.Add("ReceiverStateZip", typeof(string));
        dataTable.Columns.Add("AidDetails", typeof(string));
        // Create a new row and fill it with form data
        DataRow dataRow = dataTable.NewRow();
        dataRow["SenderName"] = txtSenderName.Text;
        dataRow["SenderAddress"] = txtSenderAddress.Text;
        dataRow["SenderStateZip"] = txtSenderStateZip.Text;
        dataRow["txtEmail"] = txtEmail.Text;
        dataRow["DateToday"] = DateTime.Parse(txtDateToday.Text).Date; // Parse the date from the TextBox and remove the time
        dataRow["ReceiverName"] = txtReceiverName.Text;
        dataRow["Title"] = txtTitle.Text;
        dataRow["LastEmploymentDate"] = DateTime.Parse(txtLastEmploymentDate.Text).Date; // Parse the date from the TextBox and remove the time
        dataRow["OrganizationName"] = txtOrganizationName.Text;
        dataRow["ReceiverAddress"] = txtReceiverAddress.Text;
        dataRow["ReceiverStateZip"] = txtReceiverStateZip.Text;
        dataRow["AidDetails"] = txtAidDetails.Text;

        // Add the row to the DataTable
        dataTable.Rows.Add(dataRow);

        // Log DataTable contents to the console or other logging mechanism
        LogDataTableContents(dataTable);

        // Add the DataTable to the DataSet
        dataSet.Tables.Add(dataTable);

        // For example, bind the data to a Crystal Report Viewer
        report.Load(Server.MapPath(LayoffReport));
        report.SetDataSource(dataTable);
        CrystalReportViewer1.ReportSource = report;
        Session["ReportDocument"] = report;

        // Set the Crystal Report as the report source for the CrystalReportViewer
        CrystalReportViewer1.ReportSource = report;
    }

    // Helper method to log DataTable contents
    private void LogDataTableContents(DataTable dataTable)
    {
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (DataColumn col in dataTable.Columns)
            {
                string columnName = col.ColumnName;
                object value = row[col];
                // Log or print the column name and value
                Debug.WriteLine(string.Format("{0}: {1}", columnName, value));

            }
        }
    }

    protected void btnAdd2Excel_Click(object sender, EventArgs e)
    { }

    protected void test(object sender, EventArgs e)
    { }

    }