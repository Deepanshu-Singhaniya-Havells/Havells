using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Job_Request
{
    public class PostCreateJobRequest : IPlugin
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
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && (context.MessageName.ToUpper() == "CREATE"))
            {
                Entity workOrderRequest = (Entity)context.InputParameters["Target"];
                try
                {
                    Callstaging(service, tracingService, workOrderRequest);
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
        private static void Callstaging(IOrganizationService service, ITracingService tracingService, Entity workOrderRequest)
        {
            try
            {
                #region Customer_Creation
                //string _registeredNumber = workOrderRequest.GetAttributeValue<string>("hil_registerednumber");
                //string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //    <entity name='contact'>
                //    <attribute name='contactid' />
                //    <filter type='and'>
                //        <condition attribute='mobilephone' operator='eq' value='{_registeredNumber}' />
                //    </filter>
                //    </entity>
                //</fetch>";

                Entity entConsumer = new Entity("contact");
                entConsumer["mobilephone"] = workOrderRequest.GetAttributeValue<string>("hil_registerednumber");
                entConsumer["address1_telephone2"] = workOrderRequest.GetAttributeValue<string>("hil_name");
                entConsumer["firstname"] = workOrderRequest.GetAttributeValue<string>("hil_firstname");
                entConsumer["lastname"] = workOrderRequest.Contains("hil_lastname") ? workOrderRequest.GetAttributeValue<string>("hil_lastname") : null;
                entConsumer["emailaddress1"] = workOrderRequest.Contains("hil_email") ? workOrderRequest.GetAttributeValue<string>("hil_email") : null;
                entConsumer["address1_telephone3"] = workOrderRequest.Contains("hil_alternativenumber") ? workOrderRequest.GetAttributeValue<string>("hil_alternativenumber") : null;
                if (workOrderRequest.Contains("hil_gender"))
                {
                    int gendervalue = workOrderRequest.GetAttributeValue<OptionSetValue>("hil_gender").Value;
                    tracingService.Trace("Gender value " + gendervalue);
                    if (gendervalue == 1)
                    {
                        entConsumer["hil_gender"] = true;
                    }
                    else if (gendervalue == 2)
                    {
                        entConsumer["hil_gender"] = false;
                    }
                    else if (gendervalue == 3)
                    {
                        entConsumer["hil_gender"] = null;
                    }
                }
                entConsumer["hil_consumersource"] = new OptionSetValue(1);
                entConsumer["hil_subscribeformessagingservice"] = true;
                //entConsumer["hil_preferredlanguageforcommunication"] = new EntityReference("hil_preferredlanguageforcommunication", new Guid("d825675f-37db-ec11-a7b5-6045bdad294c"));

                Guid contactId = service.Create(entConsumer);
                tracingService.Trace("Consumer_ID" + contactId);
                #endregion

                #region Address_Creation
                Entity address = new Entity("hil_address");
                address["hil_customer"] = new EntityReference("contact", contactId);
                address["hil_street1"] = workOrderRequest.GetAttributeValue<string>("hil_addressline1");
                address["hil_street2"] = workOrderRequest.GetAttributeValue<string>("hil_addressline2");
                address["hil_addresstype"] = new OptionSetValue(1);
                address["hil_businessgeo"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_pincode");

                Guid Addressid = service.Create(address);

                tracingService.Trace("Address_ID" + Addressid);
                #endregion

                #region Workorder_Creation

                //Entity contactdata = service.Retrieve("contact", contactId, new ColumnSet("mobilephone", "address1_telephone2", "address1_telephone3", "emailaddress1"));

                Entity Workorder = new Entity("msdyn_workorder");
                Workorder["hil_customerref"] = new EntityReference("contact", contactId);
                Workorder["hil_customername"] = workOrderRequest.GetAttributeValue<string>("hil_firstname") + " " + workOrderRequest.GetAttributeValue<string>("hil_lastname");
                Workorder["hil_mobilenumber"] = workOrderRequest.GetAttributeValue<string>("hil_registerednumber");
                Workorder["hil_callingnumber"] = workOrderRequest.GetAttributeValue<string>("hil_name");
                Workorder["hil_alternate"] = workOrderRequest.Contains("hil_alternativenumber") ? workOrderRequest.GetAttributeValue<string>("hil_alternativenumber") : null;
                Workorder["hil_email"] = workOrderRequest.Contains("hil_email") ? workOrderRequest.GetAttributeValue<string>("hil_email") : null;
                Workorder["hil_preferredtime"] = workOrderRequest.GetAttributeValue<OptionSetValue>("hil_preferreddaytime");
                Workorder["hil_preferreddate"] = workOrderRequest.GetAttributeValue<DateTime>("hil_preferreddate");
                Workorder["hil_address"] = new EntityReference("hil_address", Addressid);
                Workorder["hil_productcategory"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_productcategory");
                Workorder["hil_productcatsubcatmapping"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_productsubcategory");
                Workorder["hil_consumertype"] = new EntityReference("hil_consumertype", new Guid("484897de-2abd-e911-a957-000d3af0677f"));
                Workorder["hil_consumercategory"] = new EntityReference("hil_consumercategory", new Guid("baa52f5e-2bbd-e911-a957-000d3af0677f"));
                Workorder["hil_natureofcomplaint"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_natureofcomplaint");
                Workorder["hil_callsubtype"] = workOrderRequest.GetAttributeValue<EntityReference>("hil_callsubtype");
                Workorder["hil_quantity"] = workOrderRequest.GetAttributeValue<int>("hil_quantity"); ;
                Workorder["hil_callertype"] = workOrderRequest.GetAttributeValue<OptionSetValue>("hil_callertype");
                Workorder["hil_newserialnumber"] = workOrderRequest.GetAttributeValue<string>("hil_newserialnumber");
                Workorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
                Workorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
                Workorder["hil_customercomplaintdescription"] = workOrderRequest.Contains("hil_customerdescriptionremarks") ? workOrderRequest.GetAttributeValue<string>("hil_customerdescriptionremarks") : null;
                Workorder["hil_emergencyremarks"] = workOrderRequest.Contains("hil_emergencyremarks") ? workOrderRequest.GetAttributeValue<string>("hil_emergencyremarks") : null;
                Workorder["hil_purchasedate"] = workOrderRequest.Contains("hil_purchasedate") ? workOrderRequest.GetAttributeValue<DateTime>("hil_purchasedate") : (DateTime?)null;
                Workorder["hil_emergencycall"] = workOrderRequest.Contains("hil_emergencycall") ? workOrderRequest.GetAttributeValue<OptionSetValue>("hil_emergencycall") : null;

                Guid workorderid = service.Create(Workorder);
                tracingService.Trace("Job_ID" + workorderid);
                #endregion

                if (workorderid != Guid.Empty)
                {
                    Entity jobRequest = new Entity(workOrderRequest.LogicalName, workOrderRequest.Id);
                    jobRequest["hil_jobid"] = new EntityReference("msdyn_workorder", workorderid);
                    service.Update(jobRequest);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
