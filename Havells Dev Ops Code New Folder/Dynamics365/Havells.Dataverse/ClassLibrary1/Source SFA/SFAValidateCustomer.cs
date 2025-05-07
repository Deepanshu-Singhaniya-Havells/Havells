using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;
namespace Havells.Dataverse.CustomConnector.Source_SFA

{
    public class SFAValidateCustomer : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            Guid customerGuid = Guid.Empty;
            SFA_ValidateCustomer objValidateCustomer;
            SFA_AddressBook objAddress;
            List<SFA_AddressBook> lstAddressBook = new List<SFA_AddressBook>();
            EntityCollection entcoll;
            QueryExpression Query;
            string JsonResponse = "";

            try
            {

                string CustomerMobileNo = Convert.ToString(context.InputParameters["CustomerMobileNo"]);

                if (service != null)
                {

                    if (string.IsNullOrWhiteSpace(CustomerMobileNo))
                    {
                        objValidateCustomer = new SFA_ValidateCustomer { ResultStatus = "false", ResultMessage = "Customer Mobile No. is required.", AddressBook = new List<SFA_AddressBook>() };
                        JsonResponse = JsonSerializer.Serialize(objValidateCustomer);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    else if (!string.IsNullOrWhiteSpace(CustomerMobileNo))
                    {
                        if (!APValidate.IsValidMobileNumber(CustomerMobileNo))
                        {
                            objValidateCustomer = new SFA_ValidateCustomer { ResultStatus = "false", ResultMessage = "Customer Mobile Number is not valid.", AddressBook = new List<SFA_AddressBook>() };
                            JsonResponse = JsonSerializer.Serialize(objValidateCustomer);
                            context.OutputParameters["data"] = JsonResponse;
                            return;
                        }
                    }

                    Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("fullname", "emailaddress1");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, CustomerMobileNo);
                    entcoll = service.RetrieveMultiple(Query);

                    if (entcoll.Entities.Count == 0)
                    {
                        objValidateCustomer = new SFA_ValidateCustomer { ResultStatus = "false", ResultMessage = "Customer Mobile No. does not exist.", AddressBook = new List<SFA_AddressBook>() };
                        JsonResponse = JsonSerializer.Serialize(objValidateCustomer);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    else
                    {
                        objValidateCustomer = new SFA_ValidateCustomer();
                        objValidateCustomer.CustomerMobileNo = CustomerMobileNo;
                        objValidateCustomer.CustomerName = entcoll.Entities[0].GetAttributeValue<string>("fullname");
                        if (entcoll.Entities[0].Attributes.Contains("emailaddress1"))
                        {
                            objValidateCustomer.CustomerEmailId = entcoll.Entities[0].GetAttributeValue<string>("emailaddress1");
                        }
                        objValidateCustomer.CustomerGuid = entcoll.Entities[0].Id;
                        objValidateCustomer.ResultStatus = "true";
                        objValidateCustomer.ResultMessage = "SUCCESS";
                        objValidateCustomer.AddressBook = new List<SFA_AddressBook>();

                        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                <entity name='hil_address'>
                                                    <attribute name='hil_addressid' />
                                                    <attribute name='hil_pincode' />
                                                    <attribute name='hil_area' />
                                                    <attribute name='hil_fulladdress' />
                                                    <order attribute='hil_fulladdress' descending='false' />
                                                    <filter type='and'>
                                                        <condition attribute='hil_customer' operator='eq' value='{" + entcoll.Entities[0].Id + @"}' />
                                                    </filter>
                                                    <link-entity name='hil_area' from='hil_areaid' to='hil_area' visible='false' link-type='outer' alias='area'>
                                                        <attribute name='hil_areacode' />
                                                        <attribute name='hil_name' />
                                                    </link-entity>
                                                </entity>
                                                </fetch>";

                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (entcoll.Entities.Count > 0)
                        {
                            foreach (Entity ent in entcoll.Entities)
                            {
                                objAddress = new SFA_AddressBook();
                                objAddress.AddressGuid = ent.GetAttributeValue<Guid>("hil_addressid");
                                if (ent.Attributes.Contains("hil_fulladdress"))
                                {
                                    objAddress.Address = ent.GetAttributeValue<string>("hil_fulladdress");
                                }
                                if (ent.Attributes.Contains("hil_pincode"))
                                {
                                    objAddress.PINCode = ent.GetAttributeValue<EntityReference>("hil_pincode").Name;
                                }
                                if (ent.Attributes.Contains("area.hil_areacode"))
                                {
                                    objAddress.AreaCode = ent.GetAttributeValue<AliasedValue>("area.hil_areacode").Value.ToString();
                                }
                                objValidateCustomer.AddressBook.Add(objAddress);
                            }
                        }
                        JsonResponse = JsonSerializer.Serialize(objValidateCustomer);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                }
                else
                {
                    objValidateCustomer = new SFA_ValidateCustomer { ResultStatus = "false", ResultMessage = "D365 Service Unavailable", AddressBook = new List<SFA_AddressBook>() };
                    JsonResponse = JsonSerializer.Serialize(objValidateCustomer);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            catch (Exception ex)
            {
                objValidateCustomer = new SFA_ValidateCustomer { ResultStatus = "false", ResultMessage = "D365 Internal Server Error : " + ex.Message, AddressBook = new List<SFA_AddressBook>() };
                JsonResponse = JsonSerializer.Serialize(objValidateCustomer);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }
    }
    public class SFA_AddressBook
    {
        public string Address { get; set; }
        public string PINCode { get; set; }
        public string AreaCode { get; set; }
        public string AreaName { get; set; }
        public Guid AddressGuid { get; set; }
    }
    public class SFA_ValidateCustomer
    {
        public string CustomerMobileNo { get; set; }
        public string CustomerName { get; set; }
        public string CustomerEmailId { get; set; }
        public Guid CustomerGuid { get; set; }
        public string ResultStatus { get; set; }
        public string ResultMessage { get; set; }
        public List<SFA_AddressBook> AddressBook { get; set; }
    }
}


