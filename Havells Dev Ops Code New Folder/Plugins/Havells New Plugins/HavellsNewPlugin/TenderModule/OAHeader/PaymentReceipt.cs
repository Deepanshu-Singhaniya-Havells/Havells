using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OAHeader
{
    public class PaymentReceipt : IPlugin
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
                if (context.InputParameters.Contains("PaymentId") && context.InputParameters["PaymentId"] is string && context.Depth == 1)
                {
                    tracingService.Trace("1");
                    var PaymentIds = context.InputParameters["PaymentId"].ToString();
                    string[] ArrPymentId = PaymentIds.Split(';');
                    QueryExpression queryExp = new QueryExpression("hil_tenderpaymentdetail");
                    queryExp.ColumnSet = new ColumnSet("hil_name", "hil_paymenttype", "hil_paymentamount", "hil_voucherno", "hil_oaheader", "ownerid");
                    queryExp.Distinct = true;
                    queryExp.Criteria = new FilterExpression(LogicalOperator.Or);
                    foreach (string guid in ArrPymentId)
                    {
                        queryExp.Criteria.AddCondition("hil_tenderpaymentdetailid", ConditionOperator.Equal, new Guid(guid));

                    }
                    EntityCollection payCol = service.RetrieveMultiple(queryExp);
                    tracingService.Trace("2");
                    if (payCol.Entities.Count > 0)
                    {
                        tracingService.Trace("2");
                        EntityReference owner = payCol.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                        EntityReference tomailerref = service.Retrieve(payCol.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader").LogicalName, payCol.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader").Id, new ColumnSet("hil_omsemailersetup")).GetAttributeValue<EntityReference>("hil_omsemailersetup");
                        Entity toemailer = mailCommonConfig.getTeamUser(tomailerref, service);
                        string mailBody = mailBodtText(payCol);
                        Entity userConfiguration = mailCommonConfig.getUserConfiguartion(owner, service);
                        tracingService.Trace("3");
                        EntityCollection entCCList = new EntityCollection();
                        Entity entCC = new Entity("activityparty");
                        entCC["partyid"] = userConfiguration.GetAttributeValue<EntityReference>("hil_zonalhead");
                        entCCList.Entities.Add(entCC);
                        entCC = new Entity("activityparty");
                        entCC["partyid"] = owner;
                        entCCList.Entities.Add(entCC);
                        tracingService.Trace("4");
                        string subject = @"Payment Receipt";
                        EntityReference oaheaderregarding = payCol.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader");
                        mailCommonConfig.sendEmal(toemailer.GetAttributeValue<EntityReference>("hil_cmt"), entCCList, oaheaderregarding, mailBody, subject, service);
                    }
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);
            }
        }
        
        public string mailBodtText(EntityCollection entcoll)
        {
           
                StringBuilder sbtopBody = new StringBuilder();
                StringBuilder sbtable = new StringBuilder();
                StringBuilder sbbelowBody = new StringBuilder();
                sbtopBody.Append("<Div style='width:700pt;margin-left:-.15pt;font-weight:bold'>Dear CMT Team,</Div>");
                sbtable.Append("<Div><p>Please confirm recipt of below payment against billing took place vide OA# " + entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_oaheader").Name + " </p></Div>");
                sbtable.Append("<Div>");
                sbtable.Append("<table border=1 cellspacing=0 cellpadding=0 width=0 style='width:700pt;margin-left:-.15pt;border-collapse:collapse'>");
                sbtable.Append("<tr style='height:24.0pt;font-weight:bold;'>" +
                            "<th> Cheque No </th><th> Payment Type </th><th> Amount Received </th><th> Voucher No </th></tr>");

                foreach (Entity entCol in entcoll.Entities)
                {

                    sbtable.Append("<tr style='height:24.0pt;text-align:center'>" +
                            "<td>" + entCol.GetAttributeValue<string>("hil_name") + "</td>" +
                            "<td>" + entCol.FormattedValues["hil_paymenttype"] + " </td>" +
                              "<td>" + entCol.GetAttributeValue<Money>("hil_paymentamount").Value + " </td>" +
                              "<td>" + entCol.GetAttributeValue<string>("hil_voucherno") + " </td>" +
                            "</tr>");
                }
                sbtable.Append("</table></Div>");
                sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>Kind Regards </p></Div>");
                sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>OMS </p></Div>");
                return sbtopBody.ToString() + sbtable.ToString() + sbbelowBody.ToString();

            
        }

    }
}
