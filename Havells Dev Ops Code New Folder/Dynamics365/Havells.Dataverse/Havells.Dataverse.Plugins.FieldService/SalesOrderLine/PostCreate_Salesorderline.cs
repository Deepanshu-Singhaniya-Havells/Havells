using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.SalesOrderLine
{
    public class PostCreate_Salesorderline : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Plugin Config
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && (context.MessageName.ToUpper() == "CREATE") && context.PrimaryEntityName.ToLower() == "salesorderdetail" && context.Stage == 40)
            {
                try
                {
                    tracingService.Trace("Execution");

                    Entity orderline = (Entity)context.InputParameters["Target"];
                    orderline = service.Retrieve("salesorderdetail", orderline.Id, new ColumnSet("priceperunit", "extendedamount", "productid", "salesorderid", "hil_job"));

                    if (orderline != null)
                    {
                        Entity orderso = service.Retrieve(orderline.GetAttributeValue<EntityReference>("salesorderid").LogicalName, orderline.GetAttributeValue<EntityReference>("salesorderid").Id, new ColumnSet("customerid", "hil_serviceaddress", "requestdeliveryby", "hil_preferreddaytime", "hil_ordertype", "hil_modeofpayment"));
                        if (orderso != null)
                        {
                            if (orderso.Contains("hil_ordertype") && orderso.GetAttributeValue<EntityReference>("hil_ordertype").Id.ToString() == "019f761c-1669-ef11-a670-000d3a3e636d") //// Ordertype == AlaCarte on Order Entity
                            {
                                Entity contact = service.Retrieve("contact", orderso.GetAttributeValue<EntityReference>("customerid").Id, new ColumnSet("contactid", "firstname", "mobilephone", "address1_telephone2", "emailaddress1"));
                                Entity product = service.Retrieve(orderline.GetAttributeValue<EntityReference>("productid").LogicalName, orderline.GetAttributeValue<EntityReference>("productid").Id, new ColumnSet("hil_division", "hil_materialgroup"));

                                Entity Workorder = new Entity("msdyn_workorder");
                                Workorder["hil_customerref"] = new EntityReference("contact", orderso.GetAttributeValue<EntityReference>("customerid").Id);
                                Workorder["hil_customername"] = contact.Contains("firstname") ? contact.GetAttributeValue<string>("firstname") : null;
                                Workorder["hil_mobilenumber"] = contact.Contains("mobilephone") ? contact.GetAttributeValue<string>("mobilephone") : null;
                                Workorder["hil_callingnumber"] = contact.Contains("address1_telephone2") ? contact.GetAttributeValue<string>("address1_telephone2") : null;
                                Workorder["hil_alternate"] = contact.Contains("address1_telephone3") ? contact.GetAttributeValue<string>("address1_telephone3") : null;
                                Workorder["hil_email"] = contact.Contains("emailaddress1") ? contact.GetAttributeValue<string>("emailaddress1") : null;
                                Workorder["hil_preferredtime"] = orderso.Contains("hil_preferreddaytime") ? orderso.GetAttributeValue<OptionSetValue>("hil_preferreddaytime") : null;
                                Workorder["hil_preferreddate"] = orderso.Contains("requestdeliveryby") ? orderso.GetAttributeValue<DateTime>("requestdeliveryby") : (DateTime?)null;
                                Workorder["hil_address"] = orderso.Contains("hil_serviceaddress") ? new EntityReference("product", orderso.GetAttributeValue<EntityReference>("hil_serviceaddress").Id) : null;
                                Workorder["hil_productcategory"] = product.Contains("hil_division") ? new EntityReference("product", product.GetAttributeValue<EntityReference>("hil_division").Id) : null;
                                #region Condition check for Productcategory 
                                QueryExpression query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                                query.ColumnSet.AddColumn("hil_name");
                                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, product.GetAttributeValue<EntityReference>("hil_materialgroup").Name.ToString());

                                EntityCollection productcon = service.RetrieveMultiple(query);

                                if (productcon.Entities.Count > 0)
                                {
                                    Workorder["hil_productcatsubcatmapping"] = new EntityReference("hil_stagingdivisonmaterialgroupmapping", productcon[0].Id);
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException($"Product SubCategory Not Found while creating job");
                                }
                                #endregion
                                Workorder["hil_consumertype"] = new EntityReference("hil_consumertype", new Guid("484897de-2abd-e911-a957-000d3af0677f"));
                                Workorder["hil_consumercategory"] = new EntityReference("hil_consumercategory", new Guid("baa52f5e-2bbd-e911-a957-000d3af0677f"));
                                Workorder["hil_quantity"] = 1;
                                Workorder["hil_callertype"] = new OptionSetValue(910590001);
                                Workorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
                                Workorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
                                Workorder["hil_sourceofjob"] = new OptionSetValue(24);
                                Workorder["hil_actualcharges"] = orderline.GetAttributeValue<Money>("priceperunit");    //Requested by Mobile App Team
                                Workorder["hil_receiptamount"] = orderline.GetAttributeValue<Money>("extendedamount");    ////Requested by Mobile App Team
                                #region condition check for natureofcomplaint & callsubtype

                                var query1 = new QueryExpression("hil_natureofcomplaint");
                                query1.TopCount = 1;
                                query1.ColumnSet.AddColumns("hil_callsubtype", "hil_productcode", "hil_applicationservice", "hil_relatedproduct");
                                query1.Criteria.AddCondition("hil_applicationservice", ConditionOperator.Equal, orderline.GetAttributeValue<EntityReference>("productid").Id); // overrided discussed with kuldder sir 

                                EntityCollection noccoll = service.RetrieveMultiple(query1);

                                if (noccoll.Entities.Count > 0)
                                {

                                    Workorder["hil_natureofcomplaint"] = new EntityReference("hil_natureofcomplaint", noccoll[0].Id);
                                    Workorder["hil_callsubtype"] = new EntityReference("hil_callsubtype", noccoll[0].GetAttributeValue<EntityReference>("hil_callsubtype").Id);
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("Nature of complaint not found related to Product on order line");
                                }

                                #endregion
                                #region Check mode of payment in salesorder
                                if (orderso.GetAttributeValue<OptionSetValue>("hil_modeofpayment").Value != 1)
                                {
                                    Workorder["hil_modeofpayment"] = new OptionSetValue(2);    // Cash									
                                }
                                Workorder["hil_paymentstatus"] = new OptionSetValue(3);
                                #endregion
                                Guid workorderid = service.Create(Workorder);

                                if (workorderid != null)
                                {
                                    orderline["hil_job"] = new EntityReference("msdyn_workorder", workorderid);
                                    orderline["salesorderdetailid"] = orderline.Id;

                                    service.Update(orderline);
                                    tracingService.Trace("Job Created and updated");
                                }
                                if (workorderid != null && orderso.GetAttributeValue<OptionSetValue>("hil_modeofpayment").Value != 1)
                                {
                                    Paymentstatuscreation(service, workorderid);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException("HavellsNewPlugin.SalesOrderlinePostCreate.Execute Error :  " + ex.Message);
                }
            }
        }

        private void Paymentstatuscreation(IOrganizationService service, Guid workorderid)
        {
            Entity wo = service.Retrieve("msdyn_workorder", workorderid, new ColumnSet("msdyn_name"));
            Entity paymentstatus = new Entity("hil_paymentstatus");
            paymentstatus["hil_name"] = "D365" + wo.GetAttributeValue<string>("msdyn_name");
            paymentstatus["hil_job"] = new EntityReference("msdyn_workorder", wo.Id);
            service.Create(paymentstatus);
        }

    }
}


