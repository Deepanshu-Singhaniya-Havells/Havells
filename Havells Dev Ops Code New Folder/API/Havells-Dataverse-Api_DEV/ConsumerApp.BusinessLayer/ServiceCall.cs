using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ServiceCall
    {
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public Guid NOCGuid { get; set; }
        [DataMember]
        public string ChiefComplaint { get; set; }
        [DataMember]
        public Guid AddressGuid { get; set; }
        [DataMember]
        public string CustomerMobleNo { get; set; }
        [DataMember]
        public Guid JobGuid { get; set; }
        [DataMember]
        public string JobId { get; set; }
        [DataMember]
        public int SourceOfJob { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }

        public ServiceCall CreateServiceCall(ServiceCall serviceCalldata)
        {
            ServiceCall objServiceCall;
            Guid customerGuid = Guid.Empty;
            Guid callSubTypeGuid = Guid.Empty;
            Guid serviceCallGuid = Guid.Empty;
            Entity lookupObj = null;
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (serviceCalldata.CustomerMobleNo.Trim().Length == 0)
                {
                    return new ServiceCall { ResultStatus = false, ResultMessage = "Mobile No. is required." };
                }
                if (serviceCalldata.NOCGuid == Guid.Empty)
                {
                    return new ServiceCall { ResultStatus = false, ResultMessage = "Nature of Complaint is required." };
                }
                if (serviceCalldata.AddressGuid == Guid.Empty)
                {
                    return new ServiceCall { ResultStatus = false, ResultMessage = "Consumer Address is required." };
                }
                if (serviceCalldata.SerialNumber.Trim().Length == 0)
                {
                    return new ServiceCall { ResultStatus = false, ResultMessage = "Product Serial Number is required." };
                }
                lookupObj = service.Retrieve("hil_address", serviceCalldata.AddressGuid, new ColumnSet("hil_name"));
                if (lookupObj == null)
                {
                    return new ServiceCall { ResultStatus = false, ResultMessage = "Consumer Address does not exist." };
                }
                lookupObj = service.Retrieve("hil_natureofcomplaint", serviceCalldata.NOCGuid, new ColumnSet("hil_callsubtype"));
                if (lookupObj == null)
                {
                    return new ServiceCall { ResultStatus = false, ResultMessage = "Nature of Complaint does not exist." };
                }
                else
                {
                    if (lookupObj.Attributes.Contains("hil_callsubtype"))
                    {
                        callSubTypeGuid = lookupObj.GetAttributeValue<EntityReference>("hil_callsubtype").Id;
                    }
                }
                Query = new QueryExpression("contact");
                Query.ColumnSet = new ColumnSet("fullname");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, serviceCalldata.CustomerMobleNo);
                entcoll = service.RetrieveMultiple(Query);
                if (entcoll.Entities.Count == 0)
                {
                    return new ServiceCall { ResultStatus = false, ResultMessage = "Mobile No. does not exist." };
                }
                else {
                    customerGuid = entcoll.Entities[0].Id;
                    string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='msdyn_customerasset'>
                        <attribute name='msdyn_name' />
                        <attribute name='hil_customer' />
                        <attribute name='hil_productsubcategorymapping' />
                        <attribute name='hil_productcategory' />
                        <attribute name='msdyn_customerassetid' />
                        <order attribute='msdyn_name' descending='false' />
	                    <filter type='and'>
                            <condition attribute='hil_customer' operator='eq' value='{" + customerGuid.ToString() + @"}' />
                            <condition attribute='msdyn_name' operator='eq' value='" + serviceCalldata.SerialNumber + @"' />
                        </filter>
                        <link-entity name='contact' from='contactid' to='hil_customer' visible='false' link-type='inner' alias='con'>
                            <attribute name='fullname' />
                            <attribute name='emailaddress1' />
                        </link-entity>
                    </entity>
                    </fetch>";
                    entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entcoll.Entities.Count == 0)
                    {
                        return new ServiceCall { ResultStatus = false, ResultMessage = "Product Serial Number is not registered with Mobile No." };
                    }
                    else
                    {
                        Int32 warrantyStatus = 2;
                        objServiceCall = new ServiceCall();
                        objServiceCall = serviceCalldata;
                        Entity enWorkorder = new Entity("msdyn_workorder");

                        if (entcoll.Entities[0].Attributes.Contains("hil_customer"))
                        {
                            enWorkorder["hil_customerref"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_customer");
                        }
                        if (entcoll.Entities[0].Attributes.Contains("con.fullname"))
                        {
                            enWorkorder["hil_customername"] = entcoll.Entities[0].GetAttributeValue<AliasedValue>("con.fullname").Value.ToString();
                        }
                        if (entcoll.Entities[0].Attributes.Contains("hil_customer"))
                        {
                            enWorkorder["hil_mobilenumber"] = serviceCalldata.CustomerMobleNo;
                        }
                        if (entcoll.Entities[0].Attributes.Contains("con.emailaddress1"))
                        {
                            enWorkorder["hil_email"] = entcoll.Entities[0].GetAttributeValue<AliasedValue>("con.emailaddress1").Value.ToString();
                        }

                        if (serviceCalldata.AddressGuid != Guid.Empty)
                        {
                            enWorkorder["hil_address"] = new EntityReference("hil_address", serviceCalldata.AddressGuid);
                        }

                        if (entcoll.Entities[0].Attributes.Contains("msdyn_customerassetid"))
                        {
                            enWorkorder["msdyn_customerasset"] = new EntityReference("msdyn_customerasset", entcoll.Entities[0].GetAttributeValue<Guid>("msdyn_customerassetid"));
                        }
                        if (entcoll.Entities[0].Attributes.Contains("hil_invoicedate"))
                        {
                            enWorkorder["hil_purchasedate"] = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate");
                        }
                        if (entcoll.Entities[0].Attributes.Contains("hil_modelname"))
                        {
                            enWorkorder["hil_modelname"] = entcoll.Entities[0].GetAttributeValue<string>("hil_modelname");
                        }

                        if (entcoll.Entities[0].Attributes.Contains("hil_productcategory"))
                        {
                            enWorkorder["hil_productcategory"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productcategory");
                        }
                        if (entcoll.Entities[0].Attributes.Contains("hil_productsubcategory"))
                        {
                            enWorkorder["hil_productsubcategory"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategory");
                        }
                        if (entcoll.Entities[0].Attributes.Contains("hil_productsubcategorymapping"))
                        {
                            enWorkorder["hil_productcatsubcatmapping"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategorymapping");
                        }

                        if (entcoll.Entities[0].Attributes.Contains("hil_productsubcategory") && entcoll.Entities[0].Attributes.Contains("hil_invoicedate"))
                        {
                            DateTime invoiceDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate");
                            Query = new QueryExpression("hil_warrantytemplate");
                            Query.ColumnSet = new ColumnSet("hil_warrantyperiod");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_product", ConditionOperator.Equal, entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategory").Id);
                            Query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1);
                            EntityCollection ec = service.RetrieveMultiple(Query);
                            if (ec.Entities.Count > 0)
                            {
                                invoiceDate = invoiceDate.AddMonths(ec.Entities[0].GetAttributeValue<Int32>("hil_warrantyperiod"));
                                if (invoiceDate.CompareTo(DateTime.Now) == 1)
                                {
                                    warrantyStatus = 1; //IN Warranty
                                }
                            }
                        }
                        EntityCollection entCol;
                        enWorkorder["hil_warrantystatus"] = new OptionSetValue(warrantyStatus);
                        Query = new QueryExpression("hil_consumertype");
                        Query.ColumnSet = new ColumnSet("hil_consumertypeid");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "B2C");
                        entCol = service.RetrieveMultiple(Query);
                        if (entCol.Entities.Count > 0)
                        {
                            enWorkorder["hil_consumertype"] = entCol.Entities[0].ToEntityReference();
                        }

                        Query = new QueryExpression("hil_consumercategory");
                        Query.ColumnSet = new ColumnSet("hil_consumercategoryid");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "End User");
                        entCol = service.RetrieveMultiple(Query);
                        if (entCol.Entities.Count > 0)
                        {
                            enWorkorder["hil_consumercategory"] = entCol.Entities[0].ToEntityReference();
                        }

                        enWorkorder["hil_natureofcomplaint"] = new EntityReference("hil_natureofcomplaint", serviceCalldata.NOCGuid);

                        if (callSubTypeGuid != Guid.Empty)
                        {
                            enWorkorder["hil_callsubtype"] = new EntityReference("hil_callsubtype", callSubTypeGuid);
                        }
                        enWorkorder["hil_quantity"] = 1;
                        enWorkorder["hil_customercomplaintdescription"] = serviceCalldata.ChiefComplaint;
                        enWorkorder["hil_callertype"] = new OptionSetValue(910590001);// {SourceofJob:"Customer"}

                        enWorkorder["hil_sourceofjob"] = new OptionSetValue(SourceOfJob); // {SourceofJob:"WhatsApp"}

                        enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("D4A39573-3099-E811-A961-000D3AF05828")); // {ServiceAccount:"Dummy Account"}
                        enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("D4A39573-3099-E811-A961-000D3AF05828")); // {BillingAccount:"Dummy Account"}
                        serviceCallGuid = service.Create(enWorkorder);
                        if (serviceCallGuid != Guid.Empty)
                        {
                            objServiceCall.JobGuid = serviceCallGuid;
                            objServiceCall.JobId = service.Retrieve("msdyn_workorder", serviceCallGuid, new ColumnSet("msdyn_name")).GetAttributeValue<string>("msdyn_name");
                            objServiceCall.ResultStatus = true;
                            objServiceCall.ResultMessage = "Service Call has been registered successfully.";
                        }
                        else
                        {
                            objServiceCall.ResultStatus = false;
                            objServiceCall.ResultMessage = "FAILURE !!! Something went wrong";
                        }
                    }
                    return objServiceCall;
                }
            }
            catch (Exception ex)
            {
                objServiceCall = new ServiceCall { ResultStatus = false, ResultMessage = "FAILURE !!! Something went wrong \n" + ex.Message };
                return objServiceCall;
            }
        }
    }
}
