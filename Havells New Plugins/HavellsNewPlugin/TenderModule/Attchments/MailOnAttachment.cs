using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.Attchments
{
    public class MailOnAttachment : IPlugin
    {
        public static ITracingService tracingService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            String URL = @"https://havells.crm8.dynamics.com/main.aspx?appid=675ffa44-b6d0-ea11-a813-000d3af05d7b&forceUCI=1&pagetype=entityrecord&etn=";
            //hil_tender&id=71cab514-03d0-eb11-bacc-6045bd72e7e1";
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                   && context.PrimaryEntityName.ToLower() == "hil_attachment" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace(entity.LogicalName);
                    if (entity.Contains("regardingobjectid") || entity.Contains("hil_regardinguser"))
                    {
                        EntityReference regarding = entity.Contains("regardingobjectid") ? entity.GetAttributeValue<EntityReference>("regardingobjectid") : entity.GetAttributeValue<EntityReference>("hil_regardinguser");
                        tracingService.Trace("Regarding Logical Name  " + regarding.LogicalName);
                        if (regarding.LogicalName == "hil_tender" || regarding.LogicalName == "hil_tenderbankguarantee")
                        {
                            URL = URL + regarding.LogicalName + "&id=" + regarding.Id;
                            tracingService.Trace("URL " + URL);
                            Entity regardingENtity = service.Retrieve(regarding.LogicalName, regarding.Id, new ColumnSet("hil_customername",
                                "hil_rm", "hil_zonalhead", "hil_buhead", "hil_designteam", "ownerid", "hil_customerprojectname", "hil_salesoffice"));


                            tracingService.Trace("Regarding ENtity ");
                            EntityReference docType = entity.GetAttributeValue<EntityReference>("hil_documenttype");
                            tracingService.Trace("docType Name  " + docType.Name);
                            Entity _docuTypeentity = service.Retrieve(docType.LogicalName, docType.Id, new ColumnSet("hil_mailto"));
                            if (_docuTypeentity.Contains("hil_mailto"))
                            {
                                tracingService.Trace("Mail TO  " + _docuTypeentity.GetAttributeValue<string>("hil_mailto"));
                                String[] mailto = _docuTypeentity.GetAttributeValue<string>("hil_mailto").Split(',');
                                string subject = "Doc: '" + docType.Name + @"' Uploaded!!";
                                tracingService.Trace("Subject " + subject);
                                String emailBody = @"Hi,<br><br>" + docType.Name + @" is uploaded for " + regarding.Name + @" for:- <br>Customer Name : " + regardingENtity.GetAttributeValue<EntityReference>("hil_customername").Name + "<br>Project Name: " + regardingENtity.GetAttributeValue<string>("hil_customerprojectname") + " <br>Sales Office Name : " + regardingENtity.GetAttributeValue<EntityReference>("hil_salesoffice").Name +
                                    "<br><br>For more information please <a href=\"" + URL + "\"  target=\"_blank\">click here.</a><br><br>Regards,<br>System";
                                tracingService.Trace("mailbody  " + emailBody);

                                Entity entEmail = new Entity("email");
                                entEmail["subject"] = subject;
                                entEmail["description"] = emailBody;
                                Entity entTo = new Entity("activityparty");
                                tracingService.Trace("mailbody  ");
                                EntityCollection entToList = new EntityCollection();
                                entToList.EntityName = "systemuser";
                                foreach (string to in mailto)
                                {
                                    tracingService.Trace("to  ::::  " + to);
                                    if (to == "Branch Product Head")
                                    {
                                        entTo = new Entity("activityparty");
                                        entTo["partyid"] = regardingENtity.GetAttributeValue<EntityReference>("hil_rm");

                                    }
                                    else if (to == "Design Team")
                                    {
                                        entTo = new Entity("activityparty");
                                        entTo["partyid"] = regardingENtity.GetAttributeValue<EntityReference>("hil_designteam");
                                    }
                                    else if (to == "Enquiry Creator")
                                    {
                                        entTo = new Entity("activityparty");
                                        entTo["partyid"] = regardingENtity.GetAttributeValue<EntityReference>("ownerid");
                                    }
                                    else if (to == "Zonal Head")
                                    {
                                        entTo = new Entity("activityparty");
                                        entTo["partyid"] = regardingENtity.GetAttributeValue<EntityReference>("hil_zonalhead");
                                    }
                                    else if (to == "BU Head")
                                    {
                                        entTo = new Entity("activityparty");
                                        entTo["partyid"] = regardingENtity.GetAttributeValue<EntityReference>("hil_buhead");
                                    }
                                    else if (to == "Design Head")
                                    {
                                        entTo = new Entity("activityparty");
                                        entTo["partyid"] = service.Retrieve("systemuser", (regardingENtity.GetAttributeValue<EntityReference>("hil_designteam")).Id, new ColumnSet("parentsystemuserid")).GetAttributeValue<EntityReference>("parentsystemuserid");
                                    }
                                    else if (to == "Sr. Manager Finance")
                                    {
                                        QueryExpression query = new QueryExpression("systemuser");
                                        query.ColumnSet = new ColumnSet(true);
                                        query.Criteria = new FilterExpression(LogicalOperator.And);
                                        query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("6532dc30-40a6-eb11-9442-6045bd72b6fd"));
                                        EntityCollection _entitys = service.RetrieveMultiple(query);

                                        entTo = new Entity("activityparty");
                                        entTo["partyid"] = _entitys[0].ToEntityReference();
                                    }
                                    else if (to == "Sr. Manager Treasury")
                                    {
                                        QueryExpression query = new QueryExpression("user");
                                        query.ColumnSet = new ColumnSet(true);
                                        query.Criteria = new FilterExpression(LogicalOperator.And);
                                        query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("{45C29172-3CD0-EB11-BACC-6045BD72E9C2}"));
                                        EntityCollection _entitys = service.RetrieveMultiple(query);

                                        entTo = new Entity("activityparty");
                                        entTo["partyid"] = _entitys[0].ToEntityReference();
                                    }
                                    entToList.Entities.Add(entTo);
                                }
                                tracingService.Trace("entToList.Entities.Count): " + entToList.Entities.Count);

                                entEmail["to"] = entToList;
                                tracingService.Trace("to entity ");

                                Entity entFrom = new Entity("activityparty");
                                entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
                                Entity[] entFromList = { entFrom };
                                entEmail["from"] = entFromList;
                                entEmail["regardingobjectid"] = regardingENtity.ToEntityReference();
                                tracingService.Trace("From Name ");
                                Guid emailId = service.Create(entEmail);
                                tracingService.Trace("Email Guid " + emailId.ToString());
                                //Send email
                                SendEmailRequest sendEmailReq = new SendEmailRequest()
                                {
                                    EmailId = emailId,
                                    IssueSend = true
                                };
                                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);
                                tracingService.Trace("Email Sent..");
                            }
                            else
                            {
                                return;
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                    else { return; }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
    }
}
