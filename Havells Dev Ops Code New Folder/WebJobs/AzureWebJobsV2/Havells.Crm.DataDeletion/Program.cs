using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.Crm.DataDeletion
{
    class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = createConnection(finalString);
            string filter = $@" <filter type=""and"">
                                      <condition attribute=""createdon"" operator=""on-or-before"" value=""2023-03-25"" />
                                    </filter>
                                    <link-entity name=""hil_oaheader"" from=""hil_oaheaderid"" to=""regardingobjectid"" link-type=""outer"" alias=""bs"" />
                                    <link-entity name=""hil_tenderbankguarantee"" from=""hil_tenderbankguaranteeid"" to=""regardingobjectid"" link-type=""outer"" alias=""bt"" />
                                    <link-entity name=""hil_orderchecklist"" from=""hil_orderchecklistid"" to=""regardingobjectid"" link-type=""outer"" alias=""bu"" />
                                    <link-entity name=""hil_tender"" from=""hil_tenderid"" to=""regardingobjectid"" link-type=""outer"" alias=""bv"" />
                                    <filter type=""and"">
                                      <condition entityname=""bs"" attribute=""hil_oaheaderid"" operator=""null"" />
                                    </filter>
                                    <filter type=""and"">
                                      <condition entityname=""bt"" attribute=""hil_tenderbankguaranteeid"" operator=""null"" />
                                    </filter>
                                    <filter type=""and"">
                                      <condition entityname=""bu"" attribute=""hil_orderchecklistid"" operator=""null"" />
                                    </filter>
                                    <filter type=""and"">
                                      <condition entityname=""bv"" attribute=""hil_tenderid"" operator=""null"" />
                                    </filter>";
            //DeleteJObs(service, "email", "activityid", filter);
            filter = @"<filter type='and'>
                      <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2021-04-17' />
                    </filter>";
            DeleteJObs(service, "msdyn_workorder", "msdyn_workorderid", filter);

        }
        static void DeleteJObs(IOrganizationService service, string entityName, string primiryKey, string condition)
        {
            Console.WriteLine($"Entity {entityName} Deletion Started");
            try
            {
                int i = 0;
                int j = 0;

                while (true)
                {
                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='{entityName}'>
                                                <attribute name='{primiryKey}' />
                                                <order attribute='createdon' descending='false' />
                                                  {condition}
                                              </entity>
                                            </fetch>";

                    EntityCollection EntityList = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0)
                    {
                        break;
                    }
                    j = j + EntityList.Entities.Count;
                    Parallel.ForEach(EntityList.Entities, record =>
                     {
                         try
                         {
                             service.Delete(record.LogicalName, record.Id);
                             Console.WriteLine("Success");
                         }
                         catch (Exception ex)
                         {
                             Console.WriteLine("^^^^^ " + ex.Message);
                         }
                         i += 1;
                         Console.WriteLine(record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + ": " + record.GetAttributeValue<string>("name") + ": " + i.ToString() + "/" + j.ToString());
                     });
                }
                Console.WriteLine("Success");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        static void changeProcessLogField(IOrganizationService service)
        {
            //<condition attribute='primaryentity' operator='not-in'>
            //                            <value>10144</value>
            //                            <value>10196</value>
            //                            <value>10202</value>
            //                            <value>10200</value>
            //                            <value>10199</value>
            //                          </condition>
            string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='workflow'>
                                    <attribute name='workflowid' />
                                    <attribute name='name' />
                                    <attribute name='category' />
                                    <attribute name='primaryentity' />
                                    <attribute name='statecode' />
                                    <attribute name='createdon' />
                                    <attribute name='ownerid' />
                                    <attribute name='owningbusinessunit' />
                                    <attribute name='type' />
                                    <attribute name='syncworkflowlogonfailure' />
                                    <attribute name='asyncautodelete' />
                                    <order attribute='name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='syncworkflowlogonfailure' operator='eq' value='1' />
                                      <condition attribute='statecode' operator='eq' value='1' />
                                      <condition attribute='category' operator='eq' value='0' />
                                    </filter>
                                  </entity>
                                </fetch>";
            EntityCollection collection = service.RetrieveMultiple(new FetchExpression(fetch));

            Console.WriteLine("Totla Count " + collection.Entities.Count);

            int i = 0;

            foreach (Entity process in collection.Entities)
            {

                Guid ProcessID = process.Id;
                try
                {
                    Entity entityProcess = new Entity("workflow");
                    entityProcess["syncworkflowlogonfailure"] = false;
                    entityProcess["asyncautodelete"] = true;
                    entityProcess.Id = ProcessID;
                    service.Update(entityProcess);
                    i++;
                    Console.WriteLine("Success " + i);

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error " + ex.Message);
                    if (ex.Message == "Cannot update a published workflow definition.")
                    {
                        try
                        {
                            //Deactive Record
                            SetStateRequest deactivateRequest = new SetStateRequest
                            {
                                EntityMoniker = new EntityReference("workflow", ProcessID),
                                State = new OptionSetValue(0),
                                Status = new OptionSetValue(1)
                            };
                            service.Execute(deactivateRequest);
                            Console.WriteLine("Deactivated ");
                            // Update record
                            Entity entityProcess = new Entity("workflow");
                            entityProcess["asyncautodelete"] = true;
                            entityProcess["syncworkflowlogonfailure"] = false;
                            entityProcess.Id = ProcessID;
                            service.Update(entityProcess);
                            Console.WriteLine("Updated ");
                            //Activate Record
                            SetStateRequest activateRequest = new SetStateRequest
                            {
                                EntityMoniker = new EntityReference("workflow", ProcessID),
                                State = new OptionSetValue(1),
                                Status = new OptionSetValue(2)
                            };
                            service.Execute(activateRequest);
                            i++;
                            Console.WriteLine("Success " + i);
                        }
                        catch (Exception exc)
                        {
                            Console.WriteLine("Error " + exc.Message);
                        }
                    }
                    else
                    {
                        i++;
                        Console.WriteLine("Skip " + i);

                    }
                }
            }
        }
        static void DeleteProcessSessions(IOrganizationService service)
        {
            try
            {
                int i = 0;
                int j = 0;

                while (true)
                {
                    string fetchXML = @"<fetch top='5000'>
                    <entity name='processsession'>
                    <attribute name='processsessionid' />
                    <attribute name='createdon' />
                    <attribute name='name' />
                    <order attribute='createdon' descending='true' />
                    <filter type='and'>
                        <condition attribute='createdon' operator='on' value='2022-08-01' />
                    </filter>
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;
                    foreach (var record in EntityList.Entities)
                    {
                        try
                        {
                            service.Delete("processsession", record.Id);
                            Console.WriteLine("Success");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        i += 1;
                        Console.WriteLine(record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + ": " + record.GetAttributeValue<string>("name") + ": " + i.ToString() + "/" + j.ToString());
                    }
                }
                Console.WriteLine("Success");
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }
        static void DeleteJobProduct(IOrganizationService service, string fromDate, string toDate, string entityName)
        {
            Console.WriteLine("EntityName " + entityName + " || FromDate " + fromDate + " || ToDate " + toDate);
            int i = 0;
            int j = 0;
            while (true)
            {
                string fetchXML = @"<fetch top='5000'>
                              <entity name='" + entityName + @"'>
                                <attribute name='createdon' />
                                <order attribute='createdon' descending='false' />
                                <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='ad'>
                                  <filter type='and'>
                                    <condition attribute='msdyn_substatus' operator='in'>
                                      <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{1527FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                                      <value uiname='Closed' uitype='msdyn_workordersubstatus'>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                                      <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{6C8F2123-5106-EA11-A811-000D3AF057DD}</value>
                                    </condition>
                                    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + fromDate + @"' />
                                    <condition attribute='msdyn_timeclosed' operator='on-or-before' value='" + toDate + @"' />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";
                EntityCollection EntityList = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (EntityList.Entities.Count == 0) { break; }
                j = j + EntityList.Entities.Count;
                foreach (var record in EntityList.Entities)
                {
                    try
                    {
                        service.Delete(record.LogicalName, record.Id);
                        Console.WriteLine("Success");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    i += 1;
                    Console.WriteLine(record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + ": " + record.GetAttributeValue<string>("name") + ": " + i.ToString() + "/" + j.ToString());
                }
            }
        }
        public static IOrganizationService createConnection(string connectionString)
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
    }
}
