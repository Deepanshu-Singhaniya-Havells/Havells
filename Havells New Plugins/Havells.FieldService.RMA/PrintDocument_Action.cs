using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Util;
using System.Xml.Linq;

namespace Havells.FieldService.RMA
{
    public class PrintDocument_Action : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            try
            {
                string entityId = context.InputParameters["EntityId"].ToString();
                string entityName = context.InputParameters["EntityName"].ToString();
                string DocumentTypeID = context.InputParameters["DocumentTypeID"].ToString();
                string Challanhtml = "";
                string DocumentType = "";

                entityName = "msdyn_rmaproduct";

                Entity DocumetnTemplate = service.Retrieve("hil_inventorydocumenttype", new Guid(DocumentTypeID), new ColumnSet("hil_name", "hil_html"));
               
                Challanhtml = DocumetnTemplate.Contains("hil_html") ? DocumetnTemplate.GetAttributeValue<string>("hil_html").ToString() : throw new InvalidPluginExecutionException("Template Is Not Defin.");
                
                DocumentType = DocumetnTemplate.Contains("hil_name") ? DocumetnTemplate.GetAttributeValue<string>("hil_name").ToString() : throw new InvalidPluginExecutionException("Template Name Is Not Defin.");

                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='{entityName}'>
                            <attribute name='createdon' />
                            <attribute name='msdyn_rma' />
                            <attribute name='msdyn_quantitytoreturn' />
                            <attribute name='msdyn_qtyreceived' />
                            <attribute name='msdyn_qtyprocessed' />
                            <attribute name='msdyn_product' />
                            <attribute name='msdyn_itemstatus' />
                            <attribute name='msdyn_rmaproductid' />
                            <attribute name='msdyn_description' />
                            <order attribute='msdyn_rma' descending='true' />
                            <filter type='and'>
                              <condition attribute='msdyn_rma' operator='eq' value='{entityId}' />
                            </filter>
                            <link-entity name='msdyn_rma' from='msdyn_rmaid' to='msdyn_rma' visible='false' link-type='outer' alias='rma'>
                              <attribute name='msdyn_serviceaccount' />
                              <attribute name='createdon' />
                            </link-entity>
                            <link-entity name='msdyn_workorderproduct' from='msdyn_workorderproductid' to='msdyn_woproduct' visible='false' link-type='outer' alias='wop'>
                              <attribute name='msdyn_workorder' />
                            </link-entity>
                          </entity>
                        </fetch>";

                EntityCollection rmaproductCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (rmaproductCollection.Entities.Count > 0)
                {
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Product", typeof(string));
                    dt.Columns.Add("Quantity", typeof(double));
                    dt.Columns.Add("ChallanNo", typeof(string));
                    dt.Columns.Add("CreatedOn", typeof(string));
                    dt.Columns.Add("Description", typeof(string));
                    dt.Columns.Add("ServiceAccount", typeof(string));
                    dt.Columns.Add("JobId", typeof(string));

                    DataRow cdr = null;
                    foreach (Entity row in rmaproductCollection.Entities)
                    {
                        cdr = dt.NewRow();
                        cdr["Product"] = row.Contains("msdyn_product") ? row.GetAttributeValue<EntityReference>("msdyn_product").Name.ToString() : "";
                        cdr["Quantity"] = row.Contains("msdyn_quantitytoreturn") ? row.GetAttributeValue<double>("msdyn_quantitytoreturn") : 0;
                        cdr["ChallanNo"] = row.Contains("msdyn_rma") ? row.GetAttributeValue<EntityReference>("msdyn_rma").Name.ToString() : "";
                        cdr["CreatedOn"] = row.Contains("rma.createdon") ? ((DateTime)row.GetAttributeValue<AliasedValue>("rma.createdon").Value).AddMinutes(330).ToString() : "";
                        cdr["Description"] = row.Contains("msdyn_description") ? row.GetAttributeValue<string>("msdyn_description") : "";
                        cdr["ServiceAccount"] = row.Contains("rma.msdyn_serviceaccount") ? ((EntityReference)row.GetAttributeValue<AliasedValue>("rma.msdyn_serviceaccount").Value).Name.ToString() : "";
                        cdr["JobId"] = row.Contains("wop.msdyn_workorder") ? ((EntityReference)row.GetAttributeValue<AliasedValue>("wop.msdyn_workorder").Value).Name.ToString() : "";

                        dt.Rows.Add(cdr);
                    }

                    Challanhtml = Challanhtml.Replace("{ServiceAccount}", dt.Rows[0]["ServiceAccount"].ToString())
                           .Replace("{ChallanNo}", dt.Rows[0]["ChallanNo"].ToString())
                           .Replace("{ChallanDate}", dt.Rows[0]["CreatedOn"].ToString().Split(' ')[0]);

                    if (DocumentType == "Defective Challan")
                    {
                        string DeliveryChallanData = "";
                        DataTable dTable = dt.AsEnumerable()
                                    .GroupBy(r => new { Product = r["Product"], Description = r["Description"] })
                                    .Select(g =>
                                    {
                                        var row = dt.NewRow();

                                        row["Quantity"] = g.Sum(r => r.Field<double>("Quantity"));
                                        row["Product"] = g.Key.Product;
                                        row["Description"] = g.Key.Description;
                                        return row;
                                    }).CopyToDataTable();

                        for (int i = 0; i < dTable.Rows.Count; i++)
                        {
                            DeliveryChallanData = DeliveryChallanData + $"<tr><td>{dTable.Rows[i]["Product"]}</td><td class='text-left'>{dTable.Rows[i]["Description"]}</td><td class='text-center'>{dTable.Rows[i]["Quantity"]}</td></tr>";
                        }
                        Challanhtml = Challanhtml.Replace("{DeliveryChallanData}", DeliveryChallanData);
                    }
                    else if (DocumentType == "Defective Challan Annexure")
                    {
                        string DeliveryChallanAnnexure = "";
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DeliveryChallanAnnexure = DeliveryChallanAnnexure + $"<tr><td>{i + 1}.</td><td>{dt.Rows[i]["Product"]}</td><td class='text-left'>{dt.Rows[i]["Description"]}</td><td class='text-center'>{dt.Rows[i]["Quantity"]}</td><td class='text-left'>{dt.Rows[i]["JobId"]}</td></tr>";
                        }
                        Challanhtml = Challanhtml.Replace("{AnnexureData}", DeliveryChallanAnnexure);
                    }
                }

                context.OutputParameters["HtmlDiv"] = Challanhtml;
                context.OutputParameters["ErrorOccured"] = false;

            }
            catch (Exception ex)
            {
                context.OutputParameters["ErrorOccured"] = true;
                context.OutputParameters["Error"] = ex.Message;
            }
        }
    }
}
