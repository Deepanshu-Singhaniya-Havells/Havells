using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OAProduct
{
    public class SendMailOnInspectionStatusChangeAction : IPlugin
    {
        public static ITracingService tracingService = null;
        static public EntityReference sender = new EntityReference();
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
                if (context.InputParameters.Contains("oalineids") && context.InputParameters["oalineids"] is string && context.Depth == 1)
                {
                    tracingService.Trace("1");
                    var oalineids = context.InputParameters["oalineids"].ToString();
                    var inspectionstatus = Convert.ToInt32(context.InputParameters["inspectionstatus"].ToString());
                    var inspectiondate = context.InputParameters["inspectiondate"].ToString();
                    var remarks = context.InputParameters["remarks"].ToString();
                    string[] Arroalineids = oalineids.Split(';');
                    foreach (string guid in Arroalineids)
                    {
                        Entity prd = new Entity("hil_oaproduct");
                        prd.Id = new Guid(guid);
                        prd["hil_inspectioncallstatus"] = new OptionSetValue(inspectionstatus);
                        if (inspectionstatus == 1)
                            prd["hil_inspectioncallscheduledon"] = Convert.ToDateTime(inspectiondate);
                        service.Update(prd);
                    }

                    QueryExpression queryExp = new QueryExpression("hil_oaproduct");
                    queryExp.ColumnSet = new ColumnSet("ownerid", "hil_oaheader", "hil_productcode", "hil_quantity", "hil_inspectioncallstatus",
                        "hil_inspectioncallscheduledon", "hil_omsemailersetup");
                    queryExp.Distinct = true;
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);

                    Object[] guids = new object[Arroalineids.Length]; 
                    int i = 0;
                    foreach (string guid in Arroalineids)
                    {
                        guids[i] = new Guid(guid);
                        i++;
                    }

                    queryExp.Criteria.AddCondition("hil_oaproductid", ConditionOperator.In, guids);


                    EntityCollection Coll = service.RetrieveMultiple(queryExp);
                    if (Coll.Entities.Count > 0)
                    {

                        QueryExpression qahead = new QueryExpression("hil_oaheader");
                        qahead.ColumnSet = new ColumnSet(false);
                        qahead.Criteria = new FilterExpression(LogicalOperator.And);
                        qahead.Criteria.AddCondition("hil_oaheaderid", ConditionOperator.Equal, Coll.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader").Id);
                        LinkEntity EntityA = new LinkEntity("hil_oaheader", "hil_orderchecklist", "hil_orderchecklistid", "hil_orderchecklistid", JoinOperator.Inner);
                        EntityA.Columns = new ColumnSet("hil_rm");
                        EntityA.EntityAlias = "PEnq";
                        qahead.LinkEntities.Add(EntityA);
                        EntityCollection entOCLRM = service.RetrieveMultiple(qahead);
                        EntityReference RM = ((EntityReference)((AliasedValue)entOCLRM.Entities[0]["PEnq.hil_rm"]).Value);
                        tracingService.Trace("RM" + RM.Id);

                        string Istatus = string.Empty;
                        if (inspectionstatus == 1)
                        {
                            Istatus = "scheduled";
                        }
                        else if (inspectionstatus == 2)
                        {
                            Istatus = "waived";
                        }
                        else if (inspectionstatus == 3)
                        {
                            Istatus = "completed";
                        }

                        tracingService.Trace("2");
                        EntityReference owner = Coll.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                        string mailBody = BodyText(Coll, Istatus, remarks);
                        tracingService.Trace("3");

                        EntityCollection entTOList = new EntityCollection();
                        Entity entTO = new Entity("activityparty");
                        entTO["partyid"] = new EntityReference("systemuser", new Guid("C1E1342F-0233-EC11-B6E6-002248D4CC92"));//PAWAN2.KUMAR@HAVELLS.COM //""
                        entTOList.Entities.Add(entTO);
                        entTO = new Entity("activityparty");
                        entTO["partyid"] = new EntityReference("systemuser", new Guid("05F225D7-D671-EA11-A811-000D3AF0543F")); //DEEPSHIKHA.AGARWAL@HAVELLS.COM
                        entTOList.Entities.Add(entTO);
                        entTO = new Entity("activityparty");
                        entTO["partyid"] = new EntityReference("systemuser", new Guid("B546EE9D-FCEE-E811-A949-000D3AF03089")); //NARENDRA.YADAV@HAVELLS.COM
                        entTOList.Entities.Add(entTO);

                        tracingService.Trace("4");

                        EntityCollection entCCList = new EntityCollection();
                        Entity entCC = new Entity("activityparty");
                        entCC["partyid"] = owner;
                        entCCList.Entities.Add(entCC);
                        entCC = new Entity("activityparty");
                        entCC["partyid"] = RM;
                        entCCList.Entities.Add(entCC);
                        entCC = new Entity("activityparty");
                        entCC["partyid"] = new EntityReference("account", new Guid("723e2208-ae7d-ec11-8d21-002248d48617"));//ppc team
                        entCCList.Entities.Add(entCC);

                        tracingService.Trace("5");

                        tracingService.Trace("6");
                        string subject = @"Inspection has been " + Istatus + " against " + Coll.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader").Name;
                        tracingService.Trace("7");
                        EntityReference oaheaderregarding = Coll.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader");
                        tracingService.Trace("8");
                        sendEmal(entTOList, entCCList, oaheaderregarding, mailBody, subject, service);
                    }

                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);
            }
        }
        public string BodyText(EntityCollection entcoll, string Istatus, string remarks)
        {

            StringBuilder sbtopBody = new StringBuilder();
            StringBuilder sbtable = new StringBuilder();
            StringBuilder sbbelowBody = new StringBuilder();
            string remarkText = string.Empty;
            if (Istatus == "scheduled")
            {
                remarkText = "Remarks : " + remarks;
            }
            else
            {
                remarkText = "";
            }
            sbtopBody.Append("<Div style='width:700pt;margin-left:-.15pt;font-weight:bold'>Dear CRI Team/PPC Team,</Div>");
            sbtable.Append("<Div><p>The Inspection against OA# " + entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader").Name + " has been " + Istatus + " by client. of below item. This is for your kind information for further needful </p></Div>");
            sbtable.Append("<Div>");
            sbtable.Append("<table border=1 cellspacing=0 cellpadding=0 width=0 style='width:700pt;margin-left:-.15pt;border-collapse:collapse'>");
            sbtable.Append("<tr style='height:24.0pt;font-weight:bold;'>" +
                        "<th> OA No </th><th> Product Code </th><th> Qty </th><th> Inspection Status </th><th> Inspection Date </th></tr>");

            foreach (Entity entCol in entcoll.Entities)
            {

                sbtable.Append("<tr style='height:24.0pt;text-align:center'>" +
                        "<td>" + entCol.GetAttributeValue<EntityReference>("hil_oaheader").Name + "</td>" +
                        "<td>" + entCol.GetAttributeValue<EntityReference>("hil_productcode").Name + " </td>" +
                          "<td>" + String.Format("{0:0.00}", entCol.GetAttributeValue<decimal>("hil_quantity")) + " </td>" +
                          "<td>" + entCol.FormattedValues["hil_inspectioncallstatus"].ToString() + " </td>" +
                           "<td>" + (entCol.Contains("hil_inspectioncallscheduledon") ? entCol.GetAttributeValue<DateTime>("hil_inspectioncallscheduledon").AddMinutes(330).ToString("dd/MMM/yyyy") : "") + " </td>" +
                        "</tr>");
            }
            sbtable.Append("</table></Div>");
            sbtable.Append("<Div style='font-weight:bold;margin-top:2%;'>" + remarkText + "</Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>Kind Regards </p></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>SMS </p></Div>");
            return sbtopBody.ToString() + sbtable.ToString() + sbbelowBody.ToString();
        }
        public static void sendEmal(EntityCollection too, EntityCollection copyto, EntityReference regarding, string mailbody, string subject, IOrganizationService service)
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
