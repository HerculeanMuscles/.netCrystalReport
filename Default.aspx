<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="Default.aspx.cs" Inherits="_Default" %>
<%@ Register TagPrefix="CR" Namespace="CrystalDecisions.Web" Assembly="CrystalDecisions.Web, Version=13.0.4000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" %>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">

         <div id="formContainer ">
             <div id="inputFormSection">
                 <label for="txtSenderName">Sender Name:</label>
                 <asp:TextBox ID="txtSenderName" runat="server"></asp:TextBox>

                 <label for="txtSenderAddress">Sender Address:</label>
                 <asp:TextBox ID="txtSenderAddress" runat="server"></asp:TextBox><br />

                 <label for="txtSenderStateZip">SenderStateZip:</label>
                 <asp:TextBox ID="txtSenderStateZip" runat="server"></asp:TextBox><br />

                 <label for="txtEmail">Email:</label>
                 <asp:TextBox ID="txtEmail" runat="server"></asp:TextBox><br />
         
                 <label for="txtDateToday">Date Today:</label>
                 <asp:TextBox ID="txtDateToday" runat="server"></asp:TextBox><br />

                 <label for="txtReceiverName">Receiver Name:</label>
                 <asp:TextBox ID="txtReceiverName" runat="server"></asp:TextBox><br />

                 <label for="txtTitle">Title:</label>
                 <asp:TextBox ID="txtTitle" runat="server"></asp:TextBox><br />

                 <label for="txtOrganizationName">Organization Name:</label>
                 <asp:TextBox ID="txtOrganizationName" runat="server"></asp:TextBox><br />

                 <label for="txtReceiverAddress">Receiver Address:</label>
                 <asp:TextBox ID="txtReceiverAddress" runat="server"></asp:TextBox><br />

                 <label for="txtReceiverStateZip">ReceiverStateZip:</label>
                 <asp:TextBox ID="txtReceiverStateZip" runat="server"></asp:TextBox><br />

                 <label for="txtLastEmploymentDate">LastEmploymentDate:</label>
                 <asp:TextBox ID="txtLastEmploymentDate" runat="server"></asp:TextBox><br />

                 <label for="txtAidDetails">AidDetails:</label>
                 <asp:TextBox ID="txtAidDetails" runat="server"></asp:TextBox><br />

                 <asp:Button ID="btnGenerateReport" runat="server" Text="Generate Report" OnClick="btnGenerateReport_Click" />
                 <asp:Button ID="btnAdd2Excel" runat="server" Text="Add to Excel" OnClick="btnAdd2Excel_Click" />
                 <br />

                 <div id="FileUploadSection">
                     <asp:FileUpload ID="FileUpload1" runat="server" AllowMultiple="True" />
                        <br />

                     <asp:Button ID="Button1" runat="server" Text="Submit" OnClick="MultipleExcelFiles" />
                     <br />
                 </div>

             </div>
        </div>





    <div>
        <CR:CrystalReportViewer ID="CrystalReportViewer1" runat="server" AutoDataBind="True"
            Height="1039px" ReportSourceID="CrystalReportSource1" Width="901px" />
        <CR:CrystalReportSource ID="CrystalReportSource1" runat="server">
            <Report FileName="CrystalReport1.rpt">
            </Report>
        </CR:CrystalReportSource>
    
    </div>
    </form>

        <!-- jQuery and jQuery UI scripts -->
    <script src="https://code.jquery.com/jquery-3.6.4.min.js"></script>
    <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css" />
    <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

    <!-- Datepicker script -->
    <script>
        $(function () {
            $("#<%=txtLastEmploymentDate.ClientID%>").datepicker();
            $("#<%=txtDateToday.ClientID%>").datepicker();
        });
    </script>
</body>
</html>
