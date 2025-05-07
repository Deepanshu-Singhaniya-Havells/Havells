using Microsoft.Xrm.Sdk;
using System;
using Microsoft.Xrm.Sdk.Query;

namespace Loyalty_Program_WebJobs
{
    public class LoyaltyProgramJobs
    {
        public static void SendReminderToTechnicianForHighPriorotyJobs(IOrganizationService service)
        {

            Guid jobStatus = new Guid("2727FA6C-FA0F-E911-A94E-000D3AF060A1");
            string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <attribute name='hil_customerref' />
                    <attribute name='hil_callsubtype' />
                    <attribute name='msdyn_workorderid' />
                    <attribute name='hil_productcategory' />
                    <attribute name='hil_priorityindicator' />
                    <attribute name='hil_mobilenumber' />
                    <attribute name='ownerid' />
                    <order attribute='msdyn_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='msdyn_substatus' operator='eq' value='{jobStatus}' />
                      <condition attribute='createdon' operator='last-x-hours' value='100' />
                      <condition attribute='hil_typeofassignee' operator='in'>
                        <value uiname='DSE' uitype='position'>{{7D1ECBAB-1208-E911-A94D-000D3AF0694E}}</value>
                        <value uiname='Franchise' uitype='position'>{{4A1AA189-1208-E911-A94D-000D3AF0694E}}</value>
                        <value uiname='Franchise Technician' uitype='position'>{{0197EA9B-1208-E911-A94D-000D3AF0694E}}</value>
                      </condition>
                      <condition attribute='hil_loyaltyprogramtier' operator='not-null' />
                      <condition attribute='hil_jobpriority' operator='eq' value='0' />
                    </filter>
                    <link-entity name='hil_address' from='hil_addressid' to='hil_address' visible='false' link-type='outer' alias='add'>
                      <attribute name='hil_fulladdress' />
                    </link-entity>
                    <link-entity name='systemuser' from='systemuserid' to='owninguser' visible='false' link-type='outer' alias='asp'>
                      <attribute name='mobilephone' />
                    </link-entity>
                    <link-entity name='hil_jobsextension' from='hil_jobsextensionid' to='hil_jobextension' link-type='inner' alias='ab'>
                        <filter type='and'>
                        <condition attribute='hil_jobpendingremindersent' operator='ne' value='1' />
                        </filter>
                    </link-entity>
                  </entity>
                </fetch>";

            EntityCollection entityWO = service.RetrieveMultiple(new FetchExpression(fetchQuery));
            if (entityWO.Entities.Count > 0)
            {
                Console.WriteLine($"Total Record found for processing: {entityWO.Entities.Count.ToString()}");
                int i = 1;
                foreach (var wo in entityWO.Entities)
                {
                    string jobid = wo.GetAttributeValue<string>("msdyn_name");
                    Console.WriteLine($"Processing: JobId: {jobid} Row Count: {i++.ToString()}");
                    string jobProcessDateTime = wo.GetAttributeValue<string>("hil_priorityindicator");
                    string customerRef = wo.GetAttributeValue<EntityReference>("hil_customerref").Name + "- " + wo.GetAttributeValue<string>("hil_mobilenumber");
                    string productCategory = wo.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                    string _callSubType = wo.GetAttributeValue<EntityReference>("hil_callsubtype").Name;
                    string custAddress = wo.Contains("add.hil_fulladdress") ? wo.GetAttributeValue<AliasedValue>("add.hil_fulladdress").Value.ToString() : null;
                    string technicianMobile = wo.Contains("asp.mobilephone") ? wo.GetAttributeValue<AliasedValue>("asp.mobilephone").Value.ToString() : null;
                    if (custAddress != null)
                    {
                        custAddress = custAddress.Substring(0, 90);
                    }
                    if (technicianMobile != null && custAddress != null)
                    {
                        EntityReference entTemplate = new EntityReference("hil_smstemplates", new Guid("1f445a77-0118-ef11-840a-6045bdc64f75"));
                        EntityReference entJob = new EntityReference("msdyn_workorder", wo.Id);

                        Entity smsEntity = new Entity("hil_smsconfiguration");
                        smsEntity["hil_smstemplate"] = entTemplate;
                        smsEntity["subject"] = "High Priority job reminder " + jobid;
                        smsEntity["hil_message"] = $"REMINDER, HIGH PRIORITY job pending, close job no. {jobid} by {jobProcessDateTime}. This job is for {productCategory} for {_callSubType}, Customer Details- {customerRef}, Address- {custAddress} - Havells";
                        smsEntity["hil_mobilenumber"] = technicianMobile;
                        smsEntity["hil_direction"] = new OptionSetValue(2);
                        smsEntity["regardingobjectid"] = entJob;
                        try
                        {
                            service.Create(smsEntity);

                            string _fetchUpdate = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_jobsextension'>
                                <attribute name='hil_jobpendingremindersent' />
                                <filter type='and'>
                                    <condition attribute='hil_jobs' operator='eq' value='{wo.Id}' />
                                </filter>
                                </entity>
                                </fetch>";
                            EntityCollection entityWOExt = service.RetrieveMultiple(new FetchExpression(_fetchUpdate));
                            if (entityWOExt.Entities.Count > 0)
                            {
                                Entity _entUpdate = new Entity(entityWOExt.Entities[0].LogicalName, entityWOExt.Entities[0].Id);
                                _entUpdate["hil_jobpendingremindersent"] = true;
                                service.Update(_entUpdate);
                            }
                        }
                        catch (Exception ex)
                        {

                            Console.WriteLine(ex.Message);
                        }
                    }
                }
            }
            else {
                Console.WriteLine("No Record Found.");
            }
        }
    }
}

