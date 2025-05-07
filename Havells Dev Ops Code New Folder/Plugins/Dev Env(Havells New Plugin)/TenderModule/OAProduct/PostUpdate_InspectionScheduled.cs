using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OAProduct
{
    public class PostUpdate_InspectionScheduled : IPlugin
    {
        public static ITracingService tracingService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            //  throw new InvalidPluginExecutionException("Test");
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace(entity.LogicalName);
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    if (entity.Contains("hil_inspectioncallstatus"))
                    {
                        int _inspectionStatus = entity.GetAttributeValue<OptionSetValue>("hil_inspectioncallstatus").Value;
                        tracingService.Trace("insp :- " + _inspectionStatus);

                        if (_inspectionStatus == 1)
                        {
                            EntityReference _team = (EntityReference)entity["hil_omsemailersetup"];
                            EntityReference _OAheader = (EntityReference)entity["hil_oaheader"];

                            DateTime _deliveryDate = entity.GetAttributeValue<DateTime>("hil_deliverydate");
                            EntityReference _OAHeader = (EntityReference)entity["hil_oaheader"];
                            DateTime _inspectionDate = entity.GetAttributeValue<DateTime>("hil_inspectioncallscheduledon");
                            QueryExpression queryExp = new QueryExpression("hil_oaproduct");
                            queryExp.ColumnSet = new ColumnSet(true);
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition("hil_deliverydate", ConditionOperator.On, _deliveryDate);
                            queryExp.Criteria.AddCondition("hil_oaheader", ConditionOperator.Equal, _OAHeader.Id);
                            LinkEntity EntityA = new LinkEntity("hil_oaproduct", "hil_oaheader", "hil_oaheader", "hil_oaheaderid", JoinOperator.Inner);
                            EntityA.Columns = new ColumnSet("hil_tenderid");
                            EntityA.EntityAlias = "OAHeader";
                            queryExp.LinkEntities.Add(EntityA);
                            EntityCollection OAProductCol = service.RetrieveMultiple(queryExp);
                            bool _sendEmail = true;
                            tracingService.Trace("product Count " + OAProductCol.Entities.Count);

                            foreach (Entity _product in OAProductCol.Entities)
                            {
                                if (_product.Contains("hil_inspectioncallscheduledon"))
                                    _sendEmail = false;
                            }
                            tracingService.Trace("send Email  " + _sendEmail);

                            if (_sendEmail)
                            {
                                Entity emailr = service.Retrieve(_team.LogicalName, _team.Id, new ColumnSet("hil_ppc"));
                                EntityReference sender = getSender("EMS", service);
                                tracingService.Trace("sender " + sender.Id);
                                String URL = "";
                                Integration integration = PluginHelper.IntegrationConfiguration(service, "TenderAppURL");
                                URL = integration.uri;
                                String subject = "Inspection has Scheduled on Date " + _inspectionDate;
                                EntityCollection entTOList = new EntityCollection();
                                Entity entTO = new Entity("activityparty");
                                entTO["partyid"] = emailr.GetAttributeValue<EntityReference>("hil_ppc");
                                entTOList.Entities.Add(entTO);
                                sendEmail(emailBody(OAProductCol, _inspectionDate, URL + _OAheader.LogicalName + "&id=" + _OAheader.Id), subject, _OAheader, entTOList, null, sender, service);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
        }
        static string emailBody(EntityCollection entColPPC, DateTime inspectionScheduleOn, string URL)
        {
            StringBuilder sbtopBody = new StringBuilder();
            sbtopBody.Append("<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:0cm;line-height:normal;font-size:15px;font-family:'Calibri',sans-serif;'><span style='font-size: 18px; font-family: 'Times New Roman', Times, serif; color: black;'>Dear CRI/PPC Team,</span></p>");
            sbtopBody.Append(@"<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:0cm;line-height:normal;font-size:15px;font-family:'Calibri',sans-serif;'><span style='font-family: 'Times New Roman', Times, serif;'><br></span></p>
    <p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:0cm;line-height:normal;font-size:15px;font-family:'Calibri',sans-serif;'><span style='font-family: 'Times New Roman', Times, serif;'><span style='font-size: 18px; color: black;'>
The Inspection against below order has been scheduled by client for dated <strong>" + inspectionScheduleOn.ToString("dd/MMM/yyyy") + @"</strong>.&thinsp;You are requested to kindly book inspection slot and ensure the readiness of complete material to avoid any penalty.</span></span><span style='font-family: 'Times New Roman', Times, serif;'><br></span></p>");
            sbtopBody.Append("<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:0cm;line-height:normal;font-size:15px;font-family:'Calibri',sans-serif;'><br></p>");

            sbtopBody.Append(@"<table style='width: 100%;border:1px solid black;'>
        <tbody>
            <tr style='border:1px solid black;'>
                <td style='width: 20.0000%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><strong><span style='font-size: 16px; line-height: 107%; color: black;'>OA Number</span></strong><br></span></td>
                <td style='width: 18.7877%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><strong><span style='font-size: 16px; line-height: 107%; color: black;'>Enquiry No</span></strong><br></span></td>
                <td style='width: 20.9713%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><strong><span style='font-size: 16px; line-height: 107%; color: black;'>Product Description</span></strong><br></span></td>
                <td style='width: 20.0000%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><strong><span style='font-size: 16px; line-height: 107%; color: black;'>Readiness Date</span></strong><br></span></td>
                <td style='width: 20.0000%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><strong><span style='font-size: 16px; line-height: 107%; color: black;'>Delivery Qty.</span></strong><br></span></td>
            </tr>");
            foreach (Entity entCol in entColPPC.Entities)
            {
                sbtopBody.Append(@"<tr>");
                string _oaNumber = entCol.GetAttributeValue<EntityReference>("hil_oaheader").Name;
                string _enquiryNumber = ((EntityReference)entCol.GetAttributeValue<AliasedValue>("OAHeader.hil_tenderid").Value).Name;
                string _product = entCol.GetAttributeValue<string>("hil_productdescription");
                string _redinessDate = entCol.GetAttributeValue<DateTime>("hil_stockreadinessdateppc").ToString("dd/MMM/yyyy");
                string _deliveryQuantity = String.Format("{0:0.##}", entCol.GetAttributeValue<Decimal>("hil_quantity"));
                sbtopBody.Append(@"<td style='width: 20.0000%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><span style='font-size: 16px; line-height: 107%; color: black;'>" + _oaNumber + "</span><br></span></td>");
                sbtopBody.Append(@"<td style='width: 20.0000%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><span style='font-size: 16px; line-height: 107%; color: black;'>" + _enquiryNumber + "</span><br></span></td>");
                sbtopBody.Append(@"<td style='width: 20.0000%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><span style='font-size: 16px; line-height: 107%; color: black;'>" + _product + "</span><br></span></td>");
                sbtopBody.Append(@"<td style='width: 20.0000%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><span style='font-size: 16px; line-height: 107%; color: black;'>" + _redinessDate + "</span><br></span></td>");
                sbtopBody.Append(@"<td style='width: 20.0000%;border:1px solid black;'><span style='font-family: 'Times New Roman', Times, serif;'><span style='font-size: 16px; line-height: 107%; color: black;'>" + _deliveryQuantity + "</span><br></span></td>");
                sbtopBody.Append(@"</tr>");
            }
            sbtopBody.Append(@" </tbody>
    </table>");
            sbtopBody.Append(@"<p><span style='font-family: 'Times New Roman', Times, serif;'><br></span></p>
    <p style='margin-top:0cm;margin-right:0cm;margin-bottom:8.0pt;margin-left:0cm;line-height:normal;font-size:15px;font-family:'Calibri',sans-serif;'><span style='font-family: 'Times New Roman', Times, serif;'><span style='font-size: 18px; color: black;'>Please <a href='" + URL);
            sbtopBody.Append(@"' target='_blank'><span style='color:blue;'>click here</span></a> for more details.</span></span></p>");
            sbtopBody.Append(@"<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:0cm;line-height:normal;font-size:15px;font-family:'Calibri',sans-serif;'><span style='font-family: 'Times New Roman', Times, serif;'><span style='font-size: 18px; color: black;'>&nbsp;</span></span></p>
    <p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:0cm;line-height:normal;font-size:15px;font-family:'Calibri',sans-serif;'><span style='font-family: 'Times New Roman', Times, serif;'><span style='font-size: 18px; color: black;'>Kind Regards,</span></span></p>
    <p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:0cm;line-height:normal;font-size:15px;font-family:'Calibri',sans-serif;'><span style='font-size: 18px; font-family: 'Times New Roman', Times, serif; color: black;'><strong>OMS</strong></span></p>");


            return sbtopBody.ToString();

        }
        public static EntityReference getSender(string queName, IOrganizationService service)
        {

            EntityReference sender = new EntityReference();
            QueryExpression _queQuery = new QueryExpression("queue");
            _queQuery.ColumnSet = new ColumnSet(false);
            _queQuery.Criteria = new FilterExpression(LogicalOperator.And);
            _queQuery.Criteria.AddCondition("name", ConditionOperator.Equal, queName);
            EntityCollection queueColl = service.RetrieveMultiple(_queQuery);
            if (queueColl.Entities.Count == 1)
            {
                sender = queueColl[0].ToEntityReference();
                tracingService.Trace("sender Logical Name " + sender.LogicalName);
            }
            else
                throw new InvalidPluginExecutionException("Sender Not Found");
            return sender;
        }
        public static void sendEmail(string mailBody, string subject, EntityReference regarding, EntityCollection to, EntityCollection cc, EntityReference senderID, IOrganizationService service)
        {
            tracingService.Trace("Email Sending Function Started");
            try
            {
                Entity entEmail = new Entity("email");
                entEmail["subject"] = subject;
                tracingService.Trace("Subject " + subject);
                entEmail["description"] = mailBody;
                tracingService.Trace("mailBody " + mailBody);
                entEmail["to"] = to;
                if (cc != null)
                    if (cc.Entities.Count > 0)
                        entEmail["cc"] = cc;

                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
                Entity[] entFromList = { entFrom };
                entEmail["from"] = entFromList;

                if (regarding.Id != Guid.Empty)
                    entEmail["regardingobjectid"] = regarding;

                Guid emailId = service.Create(entEmail);
                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }

}
