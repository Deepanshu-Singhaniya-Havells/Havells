using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{
    public class ProductCodeRequision : IPlugin
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
                if (context.InputParameters.Contains("ProductId") && context.InputParameters["ProductId"] is string && context.Depth == 1)
                {
                    tracingService.Trace("1");
                    var ProductId = context.InputParameters["ProductId"].ToString();
                    string[] ArrProductId = ProductId.Split(';');
                    Guid[] guids = new Guid[ArrProductId.Length];
                    StringBuilder str = new StringBuilder();
                    for (int p = 0; p < ArrProductId.Length; p++)
                    {
                        guids[p] = new Guid(ArrProductId[p]);
                    }
                    QueryExpression query = new QueryExpression("hil_tenderproduct");
                    query.ColumnSet = new ColumnSet("hil_productdescription", "hil_quantity", "hil_selectproduct", "hil_tenderid");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("hil_tenderproductid", ConditionOperator.In, guids));
                    query.Criteria.AddCondition(new ConditionExpression("hil_product", ConditionOperator.Null));
                    EntityCollection entColl = service.RetrieveMultiple(query);

                    if (entColl.Entities.Count > 0)
                    {
                        tracingService.Trace("2");
                        StringBuilder sbtopBody = new StringBuilder();
                        StringBuilder sbtable = new StringBuilder();
                        StringBuilder sbbelowBody = new StringBuilder();

                        string TenderLogicalName = entColl.Entities[0].GetAttributeValue<EntityReference>("hil_tenderid").LogicalName;
                        Entity tender = service.Retrieve(entColl.Entities[0].GetAttributeValue<EntityReference>("hil_tenderid").LogicalName, entColl.Entities[0].GetAttributeValue<EntityReference>("hil_tenderid").Id, new ColumnSet("hil_designteam", "hil_rm", "ownerid", "hil_zonalhead", "hil_salesoffice", "hil_name", "hil_customername", "hil_customerprojectname", "hil_department"));
                        tracingService.Trace("3");
                        sbtopBody.Append("<Div style='width:700pt;margin-left:-.15pt;font-weight:bold'>Dear Design Team,</Div>");
                        sbtable.Append("<Div><p>Item code for the below new items is requested by Sales office <b>" + tender.GetAttributeValue<EntityReference>("hil_salesoffice").Name + "</b> against tender No <b>" + tender.GetAttributeValue<string>("hil_name") + "</b> of customer <b>" + tender.GetAttributeValue<EntityReference>("hil_customername").Name + "</b> for project <b>" + tender.GetAttributeValue<string>("hil_customerprojectname") + "</p></Div>");
                        sbtable.Append("<Div>");
                        sbtable.Append("<table border=1 cellspacing=0 cellpadding=0 width=0 style='width:400pt;margin-left:-.15pt;border-collapse:collapse'>");
                        sbtable.Append("<tr style='height:24.0pt;font-weight:bold;background-color: burlywood;'>" +
                                    "<th>Sno</th><th> Product Description </th></tr>");

                        int i = 1;
                        foreach (Entity ent in entColl.Entities)
                        {
                            bool selectProduct = ent.GetAttributeValue<bool>("hil_selectproduct");
                            if (selectProduct == true)
                            {
                                sbtable.Append("<tr style='height:24.0pt;text-align:center'>" + "<td>" + i + "</td><td>" + ent.GetAttributeValue<string>("hil_productdescription") + "</td>" + "</tr>");
                            }
                            i++;
                        }
                        sbtable.Append("</table></Div>");
                        sbtable.Append("<Div></Div>");
                        sbtable.Append("<Div style='font-weight:bold;margin-top:2%;'><p>You are requested to kindly provide the same ASAP, as to book the OCL and OA for further execution of order well within the delivery schedule of PO, to avoid any loss on account of Liquidated Damages.</p></Div>");

                        sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>Kind Regards </p></Div>");
                        sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>EMS </p></Div>");


                        string mailBody = sbtopBody.ToString() + sbtable.ToString() + sbbelowBody.ToString();
                        string subject = "Product Code Requisition";


                        tracingService.Trace("4");
                        EntityReference Owner = tender.GetAttributeValue<EntityReference>("ownerid");
                        EntityReference RM = tender.GetAttributeValue<EntityReference>("hil_rm");
                        EntityReference ZonalHead = tender.GetAttributeValue<EntityReference>("hil_zonalhead");
                        EntityReference DesignTeam = tender.GetAttributeValue<EntityReference>("hil_designteam");

                        Entity designManager = service.Retrieve("systemuser", DesignTeam.Id, new ColumnSet("parentsystemuserid"));
                        EntityReference designManagerId = designManager.GetAttributeValue<EntityReference>("parentsystemuserid");

                        EntityCollection entTOList = new EntityCollection();
                        Entity entTO = new Entity("activityparty");
                        entTO["partyid"] = DesignTeam;
                        entTOList.Entities.Add(entTO);

                        EntityCollection entCCList = new EntityCollection();
                        Entity entCC = new Entity("activityparty");
                        entCC["partyid"] = Owner;
                        entCCList.Entities.Add(entCC);
                        entCC = new Entity("activityparty");
                        entCC["partyid"] = RM;
                        entCCList.Entities.Add(entCC);

                        entCC = new Entity("activityparty");
                        entCC["partyid"] = ZonalHead;
                        entCCList.Entities.Add(entCC);

                        if (designManagerId != null)
                        {
                            entCC = new Entity("activityparty");
                            entCC["partyid"] = designManagerId;
                            entCCList.Entities.Add(entCC);
                        }

                        sendEmail(entTOList, entCCList, tender.ToEntityReference(), mailBody, subject, service);
                        // createTATLine("Product Code Requisition TAT", tender.ToEntityReference(), tender.GetAttributeValue<EntityReference>("hil_department"), service);

                        Entity tend = new Entity(tender.LogicalName);
                        tend.Id = tender.Id;
                        tend["hil_productrequisition"] = true;
                        tend["hil_stakeholder"] = new OptionSetValue(3);
                        service.Update(tend);
                        tracingService.Trace("stakeholder updated");
                        context.OutputParameters["Message"] = "Rquisition Made Successfully";
                    }

                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.TenderProduct.ProductCodeRequisition.Execute Error " + ex.Message);
            }
        }
        static void createTATLine(string Tatname, EntityReference regardingRef, EntityReference department, IOrganizationService service)
        {
            try
            {
                tracingService.Trace("Tatname " + Tatname);
                tracingService.Trace("regardingRef " + regardingRef.Id);
                tracingService.Trace("department " + department.Id);
                QueryExpression query = new QueryExpression("hil_salestatmaster");
                query.ColumnSet = new ColumnSet("hil_durationmin", "hil_name", "hil_department");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_department", ConditionOperator.Equal, department.Id);
                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Tatname);
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection Found = service.RetrieveMultiple(query);
                tracingService.Trace("Found.Entities.Count " + Found.Entities.Count);
                if (Found.Entities.Count > 0)
                {
                    tracingService.Trace("1.1");
                    int durationmin = Found[0].Contains("hil_durationmin") ? Found[0].GetAttributeValue<int>("hil_durationmin") : throw new InvalidPluginExecutionException("***** Duration Not Defin for this type *****"); ;
                    tracingService.Trace("durationmin 1.1 " + durationmin);
                    DateTime startTime = DateTime.Now.ToUniversalTime();
                    tracingService.Trace("1.1 " + startTime);
                    DateTime endTime = DateTime.Now.ToUniversalTime().AddMinutes(durationmin);
                    tracingService.Trace("endTime 1.1 " + endTime);
                    Entity entity = new Entity("hil_bdtatownership");
                    entity["subject"] = Found[0].Contains("hil_name") ? Found[0].GetAttributeValue<string>("hil_name") : throw new InvalidPluginExecutionException("***** name Not Defin for this type *****"); ;
                    tracingService.Trace("1.2 ");
                    entity["scheduleddurationminutes"] = durationmin;
                    tracingService.Trace("1.3 ");
                    entity["scheduledstart"] = startTime;
                    tracingService.Trace("1.3");
                    entity["scheduledend"] = endTime;
                    tracingService.Trace("1.4");
                    entity["hil_salestatmaster"] = Found[0].ToEntityReference();
                    tracingService.Trace("1.5");
                    entity["actualstart"] = startTime;
                    tracingService.Trace("1.6");
                    entity["regardingobjectid"] = regardingRef;
                    tracingService.Trace("1.7");
                    entity["hil_department"] = Found[0].GetAttributeValue<EntityReference>("hil_department");
                    tracingService.Trace("1.8");
                    Guid guid = service.Create(entity);
                    tracingService.Trace("1.9");
                }
                else
                {
                    //throw new InvalidPluginExecutionException("***** Master Not Defin for this type *****");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error on Create TAT Line " + ex.Message);
            }
        }
        public static void sendEmail(EntityCollection too, EntityCollection copyto, EntityReference regarding, string mailbody, string subject, IOrganizationService service)
        {
            try
            {
                Entity entEmail = new Entity("email");
                tracingService.Trace("9");
                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
                Entity[] entFromList = { entFrom };
                entEmail["from"] = entFromList;


                Entity toActivityParty = new Entity("activityparty");
                toActivityParty["partyid"] = too;
                entEmail["to"] = too;

                tracingService.Trace("10");
                Entity ccActivityParty = new Entity("activityparty");
                ccActivityParty["partyid"] = copyto;
                entEmail["cc"] = copyto;

                entEmail["subject"] = subject;
                entEmail["description"] = mailbody;

                entEmail["regardingobjectid"] = regarding;

                Guid emailId = service.Create(entEmail);

                tracingService.Trace("11");
                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
        }
    }
}
