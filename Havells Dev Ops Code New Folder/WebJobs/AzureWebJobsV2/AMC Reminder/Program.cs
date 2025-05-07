using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Configuration;

namespace AMC_Reminder
{
    class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();//"AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8="
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();//"https://havells.crm8.dynamics.com"
            IOrganizationService _service = ConnectToCRM(string.Format(connStr, CrmURL));
            sendEmail(_service);
        }
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        static void sendEmail(IOrganizationService service)
        {
            Entity entity = service.Retrieve("hil_production", new Guid("7401ee07-aed1-ec11-a7b6-6045bdad2d7e"), new ColumnSet(true));
            // if (context.MessageName.ToUpper() == "CREATE")
            {
                string SchemaName = entity.GetAttributeValue<string>("hil_schemaname").ToString();
                string RecordType = entity.FormattedValues["hil_recordtype"].ToString();
                string newvalue = entity.GetAttributeValue<string>("hil_newvalue");
                EntityReference entityrecord = null;
                if (entity.Contains("hil_tenderno"))
                {
                    entityrecord = entity.GetAttributeValue<EntityReference>("hil_tenderno");
                }
                else if (entity.Contains("hil_orderchecklist"))
                {
                    entityrecord = entity.GetAttributeValue<EntityReference>("hil_orderchecklist");
                }

                Console.WriteLine("RecordType " + RecordType);
                Console.WriteLine("SchemaName " + SchemaName);
                //SetApprover(service, entity);
                RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest
                {
                    EntityLogicalName = entityrecord.LogicalName,
                    LogicalName = SchemaName,
                    RetrieveAsIfPublished = false
                };
                RetrieveAttributeResponse attributeResponse =    (RetrieveAttributeResponse)service.Execute(attributeRequest);

                Console.WriteLine("Retrieved the attribute {0}.",
                    attributeResponse.AttributeMetadata.SchemaName);
                var attri = attributeResponse.AttributeMetadata;
                if (attri.AttributeType == AttributeTypeCode.Picklist)
                {
                    var option = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)attri).OptionSet.Options;
                    string optionLable = string.Empty;
                    int value = 0;
                    for (int j = 0; j < option.Count; j++)
                    {
                        Console.WriteLine("Convert.ToInt32(option[j].Value); " + Convert.ToInt32(option[j].Value));
                        optionLable += option[j].Label.LocalizedLabels[0].Label + " ,";
                        if (optionLable.Contains(newvalue))
                        {
                            if (option[j].Label.LocalizedLabels[0].Label.ToString().ToLower() == newvalue.ToLower())
                            {
                                value = Convert.ToInt32(option[j].Value);
                                Console.WriteLine("newvalue " + newvalue + " value " + value);
                            }
                        }
                    }
                    Console.WriteLine("optionLable " + optionLable);
                    Console.WriteLine("value " + value);
                    Console.WriteLine("newvalue " + newvalue);

                    if (optionLable.Contains(newvalue))
                    {
                        //    tender[SchemaName] = new OptionSetValue(value);
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Please enter correct value out of these " + optionLable);
                    }
                    Console.WriteLine("Completed");
                }
                else if (attri.AttributeType == AttributeTypeCode.Lookup)
                {
                    if (attri.SchemaName.ToLower() == SchemaName.ToLower())
                    {
                        var TargetEntity = ((Microsoft.Xrm.Sdk.Metadata.LookupAttributeMetadata)attri).Targets;
                        Console.WriteLine("TargetEntity" + TargetEntity[0].ToString());
                        Console.WriteLine("newvalue" + newvalue);
                        QueryExpression targetLook = new QueryExpression(TargetEntity[0].ToString());
                        targetLook.ColumnSet = new ColumnSet(false);
                        targetLook.Criteria = new FilterExpression(LogicalOperator.And);
                        targetLook.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, newvalue);
                        EntityCollection entity1 = service.RetrieveMultiple(targetLook);

                        if (entity1.Entities.Count > 0)
                        {
                            Console.WriteLine("Customerid " + entity1.Entities[0].Id);

                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("Please enter correct Customer Code  ");
                        }

                    }
                }
                else if (attri.AttributeType == AttributeTypeCode.String)
                {
                    if (attri.SchemaName.ToLower() == SchemaName.ToLower())
                    {
                        Console.WriteLine("customerprojectname " + newvalue);
                    }

                }

            }
        }

        public static EntityReference getSender(string queName, IOrganizationService service)
        {
            EntityReference sender = new EntityReference();
            QueryExpression _queQuery = new QueryExpression("queue");
            _queQuery.ColumnSet = new ColumnSet("name");
            _queQuery.Criteria = new FilterExpression(LogicalOperator.And);
            _queQuery.Criteria.AddCondition("name", ConditionOperator.Equal, queName);
            EntityCollection queueColl = service.RetrieveMultiple(_queQuery);
            if (queueColl.Entities.Count == 1)
            {
                sender = queueColl[0].ToEntityReference();
            }
            return sender;
        }

        public static void sendEmal(EntityReference to, EntityReference regarding, string mailbody, string subject, IOrganizationService service)
        {
            try
            {
                Entity entEmail = new Entity("email");

                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = getSender("Havells Connect", service);
                Entity[] entFromList = { entFrom };
                entEmail["from"] = entFromList;


                Entity toActivityParty = new Entity("activityparty");
                toActivityParty["partyid"] = to;
                entEmail["to"] = new Entity[] { toActivityParty };


                entEmail["subject"] = subject;
                entEmail["description"] = mailbody;

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
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
        }
        public static void sendSMS()
        {

        }
    }
}
