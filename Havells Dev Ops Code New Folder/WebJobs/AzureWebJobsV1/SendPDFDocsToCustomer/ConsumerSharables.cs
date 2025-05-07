using System;
using System.IO;
using Pechkin;
using Pechkin.Synchronized;
using System.Text;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Drawing.Printing;

namespace SendPDFDocsToCustomer
{
    public class ConsumerSharables
    {
        public void SendJobSheet(IOrganizationService service)
        {
            GlobalConfig gc = new GlobalConfig();
            gc.SetMargins(new Margins(0, 0, 0, 0))
                .SetDocumentTitle("Job Sheet")
                .SetPaperSize(PaperKind.A4);

            var oc = new ObjectConfig();
            oc.SetPrintBackground(true)
            .SetAllowLocalContent(true);
            IPechkin pechkin = new SynchronizedPechkin(gc);
            ObjectConfig configuration = new ObjectConfig();
            configuration
            .SetAllowLocalContent(true)
            .SetPrintBackground(true)
            .SetAllowLocalContent(true);
            ColumnSet selectedColounm = new ColumnSet(
                "msdyn_workorderid",
                "hil_productcategory",
                "hil_productcatsubcatmapping",
                "hil_mobilenumber",
                "msdyn_name",
                "hil_fulladdress",
                "hil_customercomplaintdescription",
                "hil_customerref",
                "hil_callingnumber",
                "hil_alternate",
                "hil_warrantystatus",
                "hil_purchasedfrom",
                "hil_countryclassification",
                "hil_actualcharges",
                "hil_quantity",
                "ownerid"
                );
            Entity _job = service.Retrieve("msdyn_workorder", new Guid("717d1b01-8992-eb11-b1ac-0022486e75ea"), selectedColounm);

            byte[] pdfContent = pechkin.Convert(configuration, htmlCode(_job, service));
            string _fileName = "TCR" + _job.GetAttributeValue<string>("msdyn_name") + ".pdf";
            Upload(_fileName, pdfContent, "devanduat");
        }
        public bool ByteArrayToFile(string _FileName, byte[] _ByteArray)
        {
            try
            {
                // Open file for reading
                FileStream _FileStream = new FileStream(_FileName, FileMode.Create, FileAccess.Write);
                // Writes a block of bytes to this stream using data from  a byte array.
                _FileStream.Write(_ByteArray, 0, _ByteArray.Length);

                // Close file stream
                _FileStream.Close();

                return true;
            }
            catch (Exception _Exception)
            {
                Console.WriteLine("Exception caught in process while trying to save : {0}", _Exception.ToString());
            }

            return false;
        }
        public String htmlCode(Entity _job, IOrganizationService service)
        {
            StringBuilder htmlWithStyles = new StringBuilder();
            try
            {
                String _product_sub_category = _job.Contains("hil_productcatsubcatmapping") ? _job.GetAttributeValue<EntityReference>("hil_productcatsubcatmapping").Name : String.Empty;
                String _registered_mobile_no = _job.Contains("hil_mobilenumber") ? _job.GetAttributeValue<string>("hil_mobilenumber") : String.Empty;
                String _quantity = _job.Contains("hil_quantity") ? _job.GetAttributeValue<int>("hil_quantity").ToString() : String.Empty;
                String _alternate_number = _job.Contains("hil_alternate") ? _job.GetAttributeValue<string>("hil_alternate") : String.Empty;
                String _customer_remarks = _job.Contains("hil_customercomplaintdescription") ? _job.GetAttributeValue<string>("hil_customercomplaintdescription") : String.Empty;
                String _complaint_number = _job.Contains("msdyn_name") ? _job.GetAttributeValue<string>("msdyn_name") : String.Empty;
                String _product_category = _job.Contains("hil_productcatsubcatmapping") ? _job.GetAttributeValue<EntityReference>("hil_productcategory").Name : String.Empty;
                String _name_of_the_customer = _job.Contains("hil_customerref") ? _job.GetAttributeValue<EntityReference>("hil_customerref").Name : String.Empty;
                String _address = _job.Contains("hil_fulladdress") ? _job.GetAttributeValue<String>("hil_fulladdress") : String.Empty;
                String _pin_code = _job.Contains("hil_fulladdress") ? _address.Substring(_address.Length - 6) : String.Empty;
                String _purchase_from = _job.Contains("hil_purchasedfrom") ? _job.GetAttributeValue<String>("hil_purchasedfrom") : String.Empty;
                String _totalCharge = _job.Contains("hil_actualcharges") ? Math.Round(_job.GetAttributeValue<Money>("hil_actualcharges").Value, 2, MidpointRounding.ToEven).ToString() : String.Empty;
                String _Warranty = String.Empty;
                String _Classification = String.Empty;
                if (_job.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 1)
                {
                    _Warranty = "In Warranty";
                }
                else if (_job.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 2)
                {
                    _Warranty = "Out Warranty";
                }
                else if (_job.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 4)
                {
                    _Warranty = "NA for Warranty";
                }
                else if (_job.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 3)
                {
                    _Warranty = "Warranty Void";
                }

                if (_job.GetAttributeValue<OptionSetValue>("hil_countryclassification").Value == 1)
                {
                    _Classification = "Local";
                }
                else if (_job.GetAttributeValue<OptionSetValue>("hil_countryclassification").Value == 2)
                {
                    _Classification = "Other District";
                }

                String _jobOwner = _job.Contains("ownerid") ? _job.GetAttributeValue<EntityReference>("ownerid").Name : String.Empty;

                #region fetch
                string _fetch = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
                                  "<entity name=\"msdyn_workorderincident\">" +
                                    "<attribute name=\"msdyn_workorderincidentid\" />" +
                                    "<attribute name=\"hil_observation\" />" +
                                    "<attribute name=\"hil_modelname\" />" +
                                    "<attribute name=\"hil_modelcode\" />" +
                                    "<attribute name=\"msdyn_customerasset\" />" +
                                    "<attribute name=\"msdyn_incidenttype\" />" +
                                    "<order attribute=\"hil_observation\" descending=\"false\" />" +
                                    "<filter type=\"and\">" +
                                      "<condition attribute=\"statecode\" operator=\"eq\" value=\"0\" />" +
                                      "<condition attribute=\"msdyn_workorder\" operator=\"eq\" uiname=\"0601217812077\" uitype=\"msdyn_workorder\" value=\"" + _job.Id + "\" />" +
                                    "</filter>" +
                                    "<link-entity name=\"msdyn_customerasset\" from=\"msdyn_customerassetid\" to=\"msdyn_customerasset\" visible=\"false\" link-type=\"outer\" alias=\"a_8852360171ebe811a96c000d3af05828\">" +
                                      "<attribute name=\"hil_invoiceno\" />" +
                                      "<attribute name=\"hil_invoicedate\" />" +
                                    "</link-entity>" +
                                    "<link-entity name=\"product\" from=\"productid\" to=\"hil_modelcode\" visible=\"false\" link-type=\"outer\" alias=\"a_75307c90ea04e911a94d000d3af06c56\">" +
                                      "<attribute name=\"description\" />" +
                                    "</link-entity>" +
                                  "</entity>" +
                                "</fetch>";
                #endregion
                StringBuilder CUSTOMER_PRODUCT_DETAILS__incident = new StringBuilder();
                EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetch));
                foreach (Entity _jobIncident in entCol.Entities)
                {
                    String _Model_Code = _jobIncident.Contains("hil_modelcode") ? _jobIncident.GetAttributeValue<EntityReference>("hil_modelcode").Name : String.Empty;
                    String _Model_Name = _jobIncident.Contains("Model.description") ? _jobIncident.GetAttributeValue<AliasedValue>("Model.description").Value.ToString() : String.Empty; ;
                    String _Inv_No = _jobIncident.Contains("a_customerasse.hil_invoiceno") ? _jobIncident.GetAttributeValue<AliasedValue>("a_customerasse.hil_invoiceno").Value.ToString() : String.Empty;
                    String _Inv_Date = _jobIncident.Contains("a_customerasse.hil_invoicedate") ? _jobIncident.GetAttributeValue<AliasedValue>("a_customerasse.hil_invoicedate").Value.ToString() : String.Empty;
                    String _Serial_Number = _jobIncident.Contains("msdyn_customerasset") ? _jobIncident.GetAttributeValue<EntityReference>("msdyn_customerasset").Name : String.Empty;
                    String _Observation = _jobIncident.Contains("hil_observation") ? _jobIncident.GetAttributeValue<EntityReference>("hil_observation").Name : String.Empty;
                    String _Cause = _jobIncident.Contains("msdyn_incidenttype") ? _jobIncident.GetAttributeValue<EntityReference>("msdyn_incidenttype").Name : String.Empty;


                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                <tr>");
                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                                <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px;\">" + _Warranty + "</td>");
                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                                <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Model_Code + " </td>");
                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                                <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Model_Name + "</td>");
                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                                <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Inv_No + "</td>");
                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                                <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Inv_Date + "</td>");
                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                                <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Serial_Number + "</td>");
                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                                <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Observation + "</td>");
                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                                <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Cause + "</td>");
                    CUSTOMER_PRODUCT_DETAILS__incident.Append("                                                                                                                                </tr>");

                }
                #region
                _fetch = "<fetch version=\"1.0\" output-format=\"xml - platform\" mapping=\"logical\" distinct=\"false\">"
                           + "<entity name=\"msdyn_workorderproduct\">"
                          + "<attribute name=\"msdyn_workorderproductid\" />"
                          + "<attribute name=\"msdyn_totalamount\" />"
                          + "<attribute name=\"hil_warrantystatus\" />"
                          + "<attribute name=\"hil_replacedpart\" />"
                          + "<order attribute=\"msdyn_totalamount\" descending=\"false\" />"
                          + "<filter type=\"and\">"
                          + "<condition attribute=\"statecode\" operator=\"eq\" value=\"0\" />"
                          + "</filter>"
                          + "<link-entity name=\"msdyn_workorder\" from=\"msdyn_workorderid\" to=\"msdyn_workorder\" link-type=\"inner\" alias=\"ad\">"
                          + "<filter type=\"and\">"
                          + "<condition attribute=\"msdyn_workorderid\" operator=\"eq\" uiname=\"0601217812077\" uitype=\"msdyn_workorder\" value=\"" + _job.Id + "\" />"
                          + "</filter>"
                          + "</link-entity>"
                          + "</entity>"
                          + "</fetch>";
                #endregion
                StringBuilder SPARE_PARTS = new StringBuilder();
                entCol = service.RetrieveMultiple(new FetchExpression(_fetch));
                foreach (Entity _jobProduct in entCol.Entities)
                {
                    String _Warranty_Status = String.Empty;
                    if (_jobProduct.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 1)
                    {
                        _Warranty_Status = "In Warranty";
                    }
                    else if (_jobProduct.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 2)
                    {
                        _Warranty_Status = "Out Warranty";
                    }
                    else if (_jobProduct.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 4)
                    {
                        _Warranty_Status = "NA for Warranty";
                    }
                    else if (_jobProduct.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 3)
                    {
                        _Warranty_Status = "Warranty Void";
                    }
                    String _Description = _jobProduct.Contains("hil_replacedpart") ? _jobProduct.GetAttributeValue<EntityReference>("hil_replacedpart").Name : String.Empty;
                    String _Charge = _jobProduct.Contains("msdyn_totalamount") ? Math.Round(_jobProduct.GetAttributeValue<Money>("msdyn_totalamount").Value, 2, MidpointRounding.ToEven).ToString() : String.Empty;

                    SPARE_PARTS.Append("                                                                                                                          <tr>");
                    SPARE_PARTS.Append("                                                                                                                                          <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px 4px 15px; border-left:1px dotted #333;\">" + _Warranty_Status + "</td>");
                    SPARE_PARTS.Append("                                                                                                                                          <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Description + "</td>");
                    SPARE_PARTS.Append("                                                                                                                                          <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Charge + "</td>");
                    SPARE_PARTS.Append("                                                                                                                          </tr>");
                }
                _fetch = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
                              "<entity name=\"msdyn_workorderservice\">" +
                                "<attribute name=\"msdyn_workorderserviceid\" />" +
                                "<attribute name=\"msdyn_service\" />" +
                                "<attribute name=\"hil_charge\" />" +
                                "<order attribute=\"msdyn_service\" descending=\"false\" />" +
                                "<filter type=\"and\">" +
                                  "<condition attribute=\"msdyn_linestatus\" operator=\"eq\" value=\"690970001\" />" +
                                "</filter>" +
                                "<link-entity name=\"msdyn_workorder\" from=\"msdyn_workorderid\" to=\"msdyn_workorder\" link-type=\"inner\" alias=\"ai\">" +
                                  "<filter type=\"and\">" +
                                    "<condition attribute=\"msdyn_workorderid\" operator=\"eq\" uiname=\"0601217812077\" uitype=\"msdyn_workorder\" value=\"{6E637280-DB4F-EB11-A812-000D3AF05F57}\" />" +
                                  "</filter>" +
                                "</link-entity>" +
                              "</entity>" +
                            "</fetch>";
                entCol = service.RetrieveMultiple(new FetchExpression(_fetch));
                foreach (Entity _jobService in entCol.Entities)
                {
                    String _Warranty_Status = String.Empty;

                    String _Description = _jobService.Contains("msdyn_service") ? _jobService.GetAttributeValue<EntityReference>("msdyn_service").Name : String.Empty;
                    String _Charge = _jobService.Contains("hil_charge") ? Math.Round(_jobService.GetAttributeValue<Decimal>("hil_charge"), 2, MidpointRounding.ToEven).ToString() : String.Empty;

                    SPARE_PARTS.Append("                                                                                                                          <tr>");
                    SPARE_PARTS.Append("                                                                                                                                          <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px 4px 15px; border-left:1px dotted #333;\"></td>");
                    SPARE_PARTS.Append("                                                                                                                                          <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Description + "</td>");
                    SPARE_PARTS.Append("                                                                                                                                          <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#8c8b8b; padding:4px 4px; border-left:1px dotted #333;\">" + _Charge + "</td>");
                    SPARE_PARTS.Append("                                                                                                                          </tr>");
                }
                #region Html...
                htmlWithStyles.Append("<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">");
                htmlWithStyles.Append("<html xmlns=\"http://www.w3.org/1999/xhtml\">");
                htmlWithStyles.Append("<head>");
                htmlWithStyles.Append("    <meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />");
                htmlWithStyles.Append("    <title>Havells Emailer</title>");
                htmlWithStyles.Append("</head>");
                htmlWithStyles.Append("<body>");
                #region Header...
                htmlWithStyles.Append("    <table border=\"0\" width=\"900px\" cellpadding=\"0\" cellspacing=\"0\" bgcolor=\"#e31e24\" align=\"center\" style=\"background-color:#e31e24; width:100%; max-width:900px;\">");
                htmlWithStyles.Append("        <tr>");
                htmlWithStyles.Append("            <td style=\"padding:20px 15px 0; background-color:#e31e24;\">");
                htmlWithStyles.Append("                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" align=\"center\">");
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                                                                                            <td rowspan=\"2\" width=\"20%\" valign=\"top\"><img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFkAAAA/CAYAAACLpmToAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAA+VpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMC1jMDYwIDYxLjEzNDc3NywgMjAxMC8wMi8xMi0xNzozMjowMCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdFJlZj0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlUmVmIyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ1M1IFdpbmRvd3MiIHhtcDpDcmVhdGVEYXRlPSIyMDIxLTAxLTE0VDE0OjI5OjUxKzA1OjMwIiB4bXA6TW9kaWZ5RGF0ZT0iMjAyMS0wMS0xNFQxNDo0Njo0OCswNTozMCIgeG1wOk1ldGFkYXRhRGF0ZT0iMjAyMS0wMS0xNFQxNDo0Njo0OCswNTozMCIgZGM6Zm9ybWF0PSJpbWFnZS9wbmciIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6M0E5NEI3RUU1NjQ5MTFFQjk4NTRCMzZDNjcwQzNGRkYiIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6M0E5NEI3RUY1NjQ5MTFFQjk4NTRCMzZDNjcwQzNGRkYiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5z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\"></td>");
                htmlWithStyles.Append("                                                                                            <td rowspan=\"2\" width=\"40%\" align=\"center\" valign=\"top\" style=\"color:#fff; font-family:Arial, Helvetica, sans-serif; text-transform:uppercase; font-size:18px; padding:0 50px 0 0; font-weight:700;\">");
                htmlWithStyles.Append("                                                                                            HAVELLS INDIA LIMITED <br />Trichy</td>");
                htmlWithStyles.Append("                                                                                            <td width=\"20%\" valign=\"top\" style=\"color:#fff; font-family:Arial, Helvetica, sans-serif; font-size:14px; padding:0; font-weight:600;\">DSE Name: </td>");
                htmlWithStyles.Append("                                                                                            <td width=\"20%\" valign=\"top\" style=\"color:#fff; font-family:Arial, Helvetica, sans-serif; font-size:14px; padding:0 0px 0 0;\">" + _jobOwner + "</td>");
                htmlWithStyles.Append("                                                                            </tr>");
                htmlWithStyles.Append("                                                            </table>");
                htmlWithStyles.Append("                                            </td>");
                htmlWithStyles.Append("        </tr>");
                htmlWithStyles.Append("        <tr>");
                htmlWithStyles.Append("            <td style=\"padding:15px;\">");
                htmlWithStyles.Append("                <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" bgcolor=\"#ffffff\" align=\"center\" style=\"background-color:#fff;\">");
                htmlWithStyles.Append("                    <tr>");
                htmlWithStyles.Append("                        <td style=\"text-align:center; padding:15px 0;\">");
                htmlWithStyles.Append("                            <h1 style=\"color:#333; font-family:Arial, Helvetica, sans-serif; text-transform:uppercase; font-size:16px; margin:0; font-weight:800;\">Job Sheet</h1>");
                htmlWithStyles.Append("                        </td>");
                htmlWithStyles.Append("                    </tr>");
                #endregion
                #region Customer Details...
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                                                                                            <td style=\"padding:0px;\">");
                htmlWithStyles.Append("                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 15px;  border-top:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"45%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Complaint Number :");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px;\">" + _complaint_number + "</td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 15px;  border-top:1px dotted #000; border-left:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"50%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Product Category: ");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px;\">" + _product_category + "</td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 15px;  border-top:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"45%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Name of the Customer:");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px;\">" + _name_of_the_customer + "</td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 15px;  border-top:1px dotted #000; border-left:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"50%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Product Sub-Category: ");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px;\">" + _product_sub_category + "</td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 15px;  border-top:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"45%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Address:");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px;\">");
                htmlWithStyles.Append("                                                                                                                                                                                                            5a 86 Yalambur Road Near Raja <br />Thearter, PERAMBALUR-TN, <br />" + _address);
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 0px;  border-top:1px dotted #000; border-left:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"50%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase; padding:5px 15px;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Quantity: ");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px; padding:5px 0px;\">" + _quantity + "</td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"50%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase; padding:5px 15px;  border-top:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Purchase From: ");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px; padding:5px 0px;  border-top:1px dotted #000;\">" + _purchase_from + "</td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"50%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase; padding:5px 15px;  border-top:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Warranty/Out Warranty:  ");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px; padding:5px 0px;  border-top:1px dotted #000;\">" + _Warranty + "</td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 15px;  border-top:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"45%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Pin code: ");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px;\">" + _pin_code + "</td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 15px;  border-top:1px dotted #000; border-left:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"50%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Classification: ");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px;\">" + _Classification + " </td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 15px;  border-top:1px dotted #000; border-bottom:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"45%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Registered Mobile No: ");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px;\">" + _registered_mobile_no + " </td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" width=\"50%\" style=\"padding:5px 15px;  border-top:1px dotted #000; border-left:1px dotted #000; border-bottom:1px dotted #000;\">");
                htmlWithStyles.Append("                                                                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td width=\"50%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; text-transform:uppercase;\">");
                htmlWithStyles.Append("                                                                                                                                                                                            Alternate Number: ");
                htmlWithStyles.Append("                                                                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; color:#8c8b8b; font-size:12px;\">" + _alternate_number + " </td>");
                htmlWithStyles.Append("                                                                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                            </td>");
                htmlWithStyles.Append("                                                                            </tr>");
                #endregion
                #region CUSTOMER PRODUCT DETAILS
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                        <td style=\"text-align:center; padding:15px 0;\">");
                htmlWithStyles.Append("                            <h1 style=\"color:#333; font-family:Arial, Helvetica, sans-serif; text-transform:uppercase; font-size:16px; margin:0; font-weight:800;\">CUSTOMER PRODUCT DETAILS</h1>");
                htmlWithStyles.Append("                        </td>");
                htmlWithStyles.Append("                    </tr>");
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                                                                                            <td style=\"padding:0 0px 15px\">");
                htmlWithStyles.Append("                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\"  style=\"border-top:1px dotted #333; border-bottom:1px dotted #333; text-align:left;\">");
                htmlWithStyles.Append("                                                                                                                            <thead>");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                #region CUSTOMER PRODUCT DETAILS header
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 4px 4px 15px; border-bottom:1px dotted #333;\">Warranty Status</th>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 4px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Model Code </th>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 4px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Model Name</th>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 4px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Inv No  </th>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 4px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Inv Date</th>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 4px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Serial Number</th>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 4px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Observation</th>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 15px 4px 4px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Cause</th>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                            </thead>");
                #endregion
                htmlWithStyles.Append("                                                                                                                            <tbody>");
                htmlWithStyles.Append("                                                                                                                            </tbody>");
                htmlWithStyles.Append(CUSTOMER_PRODUCT_DETAILS__incident.ToString());
                htmlWithStyles.Append("                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                            </td>");
                htmlWithStyles.Append("                                                                            </tr>");
                htmlWithStyles.Append("                                                                            <tr>");
                #endregion CUSTOMER PRODUCT DETAILS
                #region SPARE PARTS
                htmlWithStyles.Append("                        <td style=\"text-align:center; padding:15px 0;\">");
                htmlWithStyles.Append("                            <h1 style=\"color:#333; font-family:Arial, Helvetica, sans-serif; text-transform:uppercase; font-size:16px; margin:0; font-weight:800;\">SPARE PARTS</h1>");
                htmlWithStyles.Append("                        </td>");
                htmlWithStyles.Append("                    </tr>");
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                                                                                            <td style=\"padding:0 0px\">");
                htmlWithStyles.Append("                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\"  style=\"border-top:1px dotted #333; border-bottom:1px dotted #333; text-align:left;\">");
                #region SPARE PARTS header
                htmlWithStyles.Append("                                                                                                                            <thead>");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 4px 4px 15px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Warranty Status</th>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 4px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Description</th>");
                htmlWithStyles.Append("                                                                                                                                            <th style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; color:#585858; padding:4px 15px 4px 4px; border-bottom:1px dotted #333; border-left:1px dotted #333;\">Charge</th>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                            </thead>");
                #endregion
                htmlWithStyles.Append("                                                                                                                            <tbody>");
                htmlWithStyles.Append(SPARE_PARTS.ToString());
                htmlWithStyles.Append("                                                                                                                            </tbody>");
                htmlWithStyles.Append("                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                            </td>");
                htmlWithStyles.Append("                                                                            </tr>");
                #endregion
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                                                                                            <td style=\"padding:0; border-bottom:1px dotted #333;\">");
                htmlWithStyles.Append("                                                                                                            <table width=\"/100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                <tr>");
                htmlWithStyles.Append("                                    <td width=\"21%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:16px;  font-weight:700; text-transform:uppercase; color:#585858; padding:4px 4px;\">Total Charges :</td>");
                htmlWithStyles.Append("                                    <td width=\"24%\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:16px; color:#8c8b8b; padding:4px 4px;\">" + _totalCharge + "</td>");
                htmlWithStyles.Append("                                </tr>");
                htmlWithStyles.Append("                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                            </td>");
                htmlWithStyles.Append("                                                                            </tr>");
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                        <td style=\"padding:20px 15px 15px; border-bottom:1px dotted #333;\">");
                htmlWithStyles.Append("                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:600; color:#585858; text-transform:uppercase; width:200px;\">Customer Remarks:</td>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; color:#8c8b8b; text-transform:uppercase;\">" + _customer_remarks + "</td>");
                htmlWithStyles.Append("                                                                                                                                            <td valign=\"top\" style=\"text-align:center; width:200px;\">");
                htmlWithStyles.Append("                                                                                                                                                            <p style=\"margin:0 0 5px;\"><img src=\"https://d365storagesa.blob.core.windows.net/d365-workorder/1604219152186.jpg\"></p>");
                htmlWithStyles.Append("                                                                                                                                                            <h3 style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; margin:0 0 5px; color:#585858; text-transform:uppercase;\">Customer Signature</h3>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                            </td>");
                htmlWithStyles.Append("                                                                            </tr>");
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                        <td style=\"padding:20px 15px 15px;\">");
                htmlWithStyles.Append("                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:600; color:#585858; text-transform:uppercase; width:200px;\">Job Sheet No:</td>");
                htmlWithStyles.Append("                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; color:#8c8b8b; text-transform:uppercase; width:160px;\">" + _complaint_number + " </td>");
                htmlWithStyles.Append("                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:600; color:#585858; text-transform:uppercase;\">Temporary Cash /  Cheque Receipt </td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                                            <p style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; margin:20px 0 5px; color:#8c8b8b; text-transform:uppercase;\">RECEIVED WITH THANKS THE SUM OF RS () TOWARDS SERVICE CHARGES. </p>");
                htmlWithStyles.Append("                                                                                                            <p style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; margin:0; color:#8c8b8b; text-transform:uppercase;\">THIS IS A COMPUTER GENERATED DOCUMENT AND SIGNATURe Not Required</p>");
                htmlWithStyles.Append("                                                                                            </td>");
                htmlWithStyles.Append("                                                                            </tr>");
                htmlWithStyles.Append("                    <tr>");
                htmlWithStyles.Append("                        <td style=\"padding:50px 15px 15px;\">");
                htmlWithStyles.Append("                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; font-weight:600; color:#585858; text-transform:uppercase;\">Engineer Signature </td>");
                htmlWithStyles.Append("                                                                                                                                            <td style=\"text-align:center; width:200px;\">");
                htmlWithStyles.Append("                                                                                                                                                            <h3 style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; margin:0 0 5px; color:#585858; text-transform:uppercase;\">For Havells India Ltd.</h3>");
                htmlWithStyles.Append("                                                                                                                                                            <p style=\"font-family:Arial, Helvetica, sans-serif; font-size:13px; margin:0; color:#8c8b8b; text-transform:uppercase;\">T02 CHILL TECH ENGINEERS</p>");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                            </td>");
                htmlWithStyles.Append("                    </tr>");
                htmlWithStyles.Append("                    <tr>");
                htmlWithStyles.Append("                        <td style=\"padding:0px; border-top:1px solid #333\"> ");
                htmlWithStyles.Append("                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"background-color:#ebebeb; text-align:center;\">");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td align=\"center\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:18px; padding:10px 0 0px; color:#000; font-weight:800; text-transform:uppercase;\">Havells India Limited</td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td align=\"center\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; padding:0px 0 0px; color:#000; font-weight:600; text-transform:uppercase;\">Corporate Office</td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td style=\"font-family:Arial, Helvetica, sans-serif; font-size:14px; padding:0px 0 15px; color:#000; font-weight:600; color:#585858;\">QRG Towers, 2D, Sec- 126, Expressway Noida - 201304 U.P. ( India)</td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                            </td>");
                htmlWithStyles.Append("                    </tr>");
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                        <td style=\"padding:20px 15px; background:#4f4f4f;\"> ");
                htmlWithStyles.Append("                                                                                                            <table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"text-align:center;\">");
                htmlWithStyles.Append("                                                                                                                            <tr>");
                htmlWithStyles.Append("                                                                                                                                            <td align=\"center\" style=\"font-family:Arial, Helvetica, sans-serif; font-size:12px; line-height:18px; padding:px 0 0px; color:#fff;\">");
                htmlWithStyles.Append("                                                                                                                                                            For Havells Products : Customer Care No.: 1800 103 1313, 1800 11 0303, Email: customercare@havells.com <br />");
                htmlWithStyles.Append("                                                                                                                                                            For LLoyd Products : Customer Care No.: 1800 102 0666, 1800 4251 5666, Email: perfectservice@lloydmail.com <br />");
                htmlWithStyles.Append("                                                                                                                                                            For Sale Service : www.consumerconnect.havells.com");
                htmlWithStyles.Append("                                                                                                                                            </td>");
                htmlWithStyles.Append("                                                                                                                            </tr>");
                htmlWithStyles.Append("                                                                                                            </table>");
                htmlWithStyles.Append("                                                                                            </td>");
                htmlWithStyles.Append("                    </tr>");
                htmlWithStyles.Append("                                                                            <tr>");
                htmlWithStyles.Append("                                                                                            <td style=\"padding:10px 0; text-align:center;\"> <img src=\"group-logo.png\"></td>");
                htmlWithStyles.Append("                                                                            </tr>");
                htmlWithStyles.Append("                </table>");
                htmlWithStyles.Append("            </td>");
                htmlWithStyles.Append("        </tr>");
                htmlWithStyles.Append("    </table>");
                htmlWithStyles.Append("</body>");
                htmlWithStyles.Append("</html>");
                #endregion HTML

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return htmlWithStyles.ToString();
        }
        static string ToURLSlug(string s)
        {
            return Regex.Replace(s, @"[^a-z0-9]+", "-", RegexOptions.IgnoreCase)
                .Trim(new char[] { '-' })
                .ToLower();
        }
        static string Upload(string fileName, byte[] fileContent, string containerName)
        {
            string _blobURI = string.Empty;
            try
            {
                //byte[] fileContent = Convert.FromBase64String(noteBody);
                string ConnectionSting = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";
                // create object of storage account
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(ConnectionSting);

                // create client of storage account
                CloudBlobClient client = storageAccount.CreateCloudBlobClient();

                // create the reference of your storage account
                CloudBlobContainer container = client.GetContainerReference(ToURLSlug(containerName));

                // check if the container exists or not in your account
                var isCreated = container.CreateIfNotExists();

                // set the permission to blob type
                container.SetPermissionsAsync(new BlobContainerPermissions
                { PublicAccess = BlobContainerPublicAccessType.Blob });

                // create the memory steam which will be uploaded
                using (MemoryStream memoryStream = new MemoryStream(fileContent))
                {
                    // set the memory stream position to starting
                    memoryStream.Position = 0;

                    // create the object of blob which will be created
                    // Test-log.txt is the name of the blob, pass your desired name
                    CloudBlockBlob blob = container.GetBlockBlobReference(fileName);

                    // get the mime type of the file
                    string mimeType = "application/unknown";
                    string ext = (fileName.Contains(".")) ?
                                System.IO.Path.GetExtension(fileName).ToLower() : "." + fileName;
                    Microsoft.Win32.RegistryKey regKey =
                                Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
                    if (regKey != null && regKey.GetValue("Content Type") != null)
                        mimeType = regKey.GetValue("Content Type").ToString();

                    // set the memory stream position to zero again
                    // this is one of the important stage, If you miss this step, 
                    // your blob will be created but size will be 0 byte
                    memoryStream.ToArray();
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    // set the mime type
                    blob.Properties.ContentType = mimeType;

                    // upload the stream to blob
                    blob.UploadFromStream(memoryStream);
                    _blobURI = blob.Uri.AbsoluteUri;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Havells_Plugin.Annotations.PreCreate.Upload: " + ex.Message);
            }
            return _blobURI;
        }
    }
}
