using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells_Plugin.WorkOrder
{
    public class PostUpdate_SendHighPriorityJobSMS : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && context.MessageName.ToUpper() == "UPDATE" && context.Depth == 1)
                {
                    tracingService.Trace("Step-1");
                    string _fetchXml = string.Empty;
                    string jobId = string.Empty;

                    Entity entity = (Entity)context.InputParameters["Target"];
                    _fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='msdyn_workorder'>
                            <attribute name='msdyn_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_mobilenumber' />
                            <attribute name='hil_typeofassignee' />
                            <attribute name='hil_jobpriority' />
                            <attribute name='hil_priorityindicator' />
                            <attribute name='hil_customerref' />
                            <attribute name='hil_callsubtype' />
                            <attribute name='msdyn_workorderid' />
                            <attribute name='ownerid' />
                            <attribute name='hil_customerref' />
                            <attribute name='hil_productcategory' />
                            <attribute name='hil_fulladdress' />
                            <filter type='and'>
                                <condition attribute='msdyn_workorderid' operator='eq' value='{entity.Id}' />
                                <condition attribute='msdyn_substatus' operator='eq' value='{{2727FA6C-FA0F-E911-A94E-000D3AF060A1}}'/>   
                                <condition attribute='hil_typeofassignee' operator='in'>
                                    <value uiname='DSE' uitype='position'>{{7D1ECBAB-1208-E911-A94D-000D3AF0694E}}</value>
                                    <value uiname='Franchise' uitype='position'>{{4A1AA189-1208-E911-A94D-000D3AF0694E}}</value>
                                    <value uiname='Franchise Technician' uitype='position'>{{0197EA9B-1208-E911-A94D-000D3AF0694E}}</value>
                                </condition>
                                <condition attribute='hil_loyaltyprogramtier' operator='not-null' />
                                <condition attribute='hil_jobpriority' operator='eq' value='0' />
                            </filter>
                        </entity>
                    </fetch>";

                    EntityCollection entityWO = service.RetrieveMultiple(new FetchExpression(_fetchXml));
                    if (entityWO.Entities.Count > 0)
                    {
                        tracingService.Trace(entityWO.Entities.Count.ToString());
                        foreach (var workorder in entityWO.Entities)
                        {
                            jobId = workorder.GetAttributeValue<string>("msdyn_name");
                            Guid jobGuid = workorder.Id;
                            EntityReference typeofAssignee = workorder.Contains("hil_typeofassignee") ? workorder.GetAttributeValue<EntityReference>("hil_typeofassignee") : null;
                            EntityReference ownerId = workorder.Contains("ownerid") ? workorder.GetAttributeValue<EntityReference>("ownerid") : null;
                            string franchiseeMobNumber = service.Retrieve("systemuser", ownerId.Id, new ColumnSet("mobilephone")).GetAttributeValue<string>("mobilephone");

                            if (!string.IsNullOrEmpty(franchiseeMobNumber))
                            {
                                tracingService.Trace("Franchisee Mobile Number which is present on job " + franchiseeMobNumber);
                                DateTime createdOn = workorder.GetAttributeValue<DateTime>("createdon");
                                string priorityIndicator = workorder.GetAttributeValue<string>("hil_priorityindicator");
                                EntityReference customerRef = workorder.Contains("hil_customerref") ? workorder.GetAttributeValue<EntityReference>("hil_customerref") : null;
                                string customerMobNumber = workorder.GetAttributeValue<string>("hil_mobilenumber");
                                EntityReference productCategoryRef = workorder.Contains("hil_productcategory") ? workorder.GetAttributeValue<EntityReference>("hil_productcategory") : null;
                                EntityReference callsubtypeRef = workorder.Contains("hil_callsubtype") ? workorder.GetAttributeValue<EntityReference>("hil_callsubtype") : null;
                                string customerAddress = workorder.Contains("hil_fulladdress") ? workorder.GetAttributeValue<string>("hil_fulladdress") : "";
                                if (customerAddress != "")
                                    customerAddress = AddressLength(customerAddress);

                                EntityReference jobIdRef = new EntityReference("msdyn_workorder", jobGuid);
                                EntityReference smsTemplate = null;
                                if (typeofAssignee.Name == "Franchise" || typeofAssignee.Name == "DSE")
                                {
                                    smsTemplate = new EntityReference("hil_smstemplates", new Guid("a4ed3328-0118-ef11-840a-6045bdc64f75"));
                                    // Send SMS to Franchisee
                                    Entity smsEntity = new Entity("hil_smsconfiguration");
                                    smsEntity["hil_smstemplate"] = smsTemplate;
                                    smsEntity["subject"] = $"HIGH PRIORITY JOB#{jobId} SEND SMS TO Franchisee Technician on Job Creation";
                                    smsEntity["hil_message"] = $"HIGH PRIORITY Job created, Assign Technician to close job no. {jobId} by {priorityIndicator}. This job is for {productCategoryRef.Name} for {callsubtypeRef.Name}, Customer Details- {customerRef.Name + " " + customerMobNumber} Address- {customerAddress} - Havells";
                                    smsEntity["hil_mobilenumber"] = franchiseeMobNumber;//"9621961190";
                                    smsEntity["hil_direction"] = new OptionSetValue(2); // Outgoing
                                    smsEntity["regardingobjectid"] = jobIdRef;
                                    service.Create(smsEntity);
                                }

                                else if (typeofAssignee.Name == "Franchise Technician" || typeofAssignee.Name == "DSE")
                                {
                                    smsTemplate = new EntityReference("hil_smstemplates", new Guid("46a00953-0118-ef11-840a-6045bdc64f75"));
                                    // Send SMS to Technician Franchisee
                                    Entity smsEntity = new Entity("hil_smsconfiguration");
                                    smsEntity["hil_smstemplate"] = smsTemplate;
                                    smsEntity["subject"] = $"HIGH PRIORITY JOB#{jobId} SEND SMS TO Franchisee Technician on Job Assignment";
                                    smsEntity["hil_message"] = $"HIGH PRIORITY Job Assigned, close job no. {jobId} by {priorityIndicator}. This job is for {productCategoryRef.Name} for {callsubtypeRef.Name}, Customer Details- {customerRef.Name + " " + customerMobNumber} Address- {customerAddress} - Havells";
                                    smsEntity["hil_mobilenumber"] = franchiseeMobNumber;//"9621961190";
                                    smsEntity["hil_direction"] = new OptionSetValue(2); // Outgoing
                                    smsEntity["regardingobjectid"] = jobIdRef;
                                    service.Create(smsEntity);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Workorder.SendHighPriorityJobSMS.PostUpdate.Execute" + ex.Message);
            }
        }
        // Method to check Customer Addres Length 
        private static string AddressLength(string customerAddress)
        {
            string FinalAddress = customerAddress;
            if (customerAddress.Length > 90)
            {
                FinalAddress = customerAddress.Substring(0, 90);
            }
            return FinalAddress;
        }
    }
}
