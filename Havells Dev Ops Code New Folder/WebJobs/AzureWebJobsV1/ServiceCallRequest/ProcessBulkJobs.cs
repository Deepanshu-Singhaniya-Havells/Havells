using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace ServiceCallRequest
{
    public class ProcessBulkJobs 
    {
        public void Execute(IOrganizationService service)
        {
            Guid jobId = Guid.Empty;
            bool inputsValidated = true;
            string strInputsValidationsummary = string.Empty;
            string addressRemarks = string.Empty;

            DateTime? hil_expecteddeliverydate = null;
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_bulkjobsuploader'>
                    <all-attributes />
                    <order attribute='hil_customermobileno' descending='false' />
                    <filter type='and'>
                      <condition attribute='createdon' operator='on-or-after' value='2024-05-07' />
                      <condition attribute='hil_jobstatus' operator='eq' value='0' />
                      <condition attribute='hil_description' operator='null' />
                    </filter>
                  </entity>
                </fetch>";

                //string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //  <entity name='hil_bulkjobsuploader'>
                //    <all-attributes />
                //    <order attribute='hil_customermobileno' descending='false' />
                //    <filter type='and'>
                //        <condition attribute='createdon' operator='on-or-after' value='2024-05-07' />
                //        <condition attribute='hil_jobstatus' operator='eq' value='0' />
                //        <condition attribute='hil_description' operator='like' value='%Area does not exist in Area Master. Customer Address does not exist.%' />
                //    </filter>
                //  </entity>
                //</fetch>";

                EntityCollection entColJobs = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 0, totalCount = 0;
                if (entColJobs.Entities.Count > 0) { totalCount = entColJobs.Entities.Count; } else { return; }
                foreach (Entity entity in entColJobs.Entities)
                {
                    jobId = Guid.Empty;
                    inputsValidated = true;
                    strInputsValidationsummary = string.Empty;
                    addressRemarks = string.Empty;

                    Console.WriteLine("Processing .. " + i++ + "/" + totalCount);
                    Entity entityUpdate = new Entity(entity.LogicalName, entity.Id);
                    try
                    {
                        #region Getting Data from Excel Sheet Columns
                        string hil_addressline1 = entity.GetAttributeValue<string>("hil_addressline1");

                        string hil_addressline2 = string.Empty;
                        if (entity.Attributes.Contains("hil_addressline2"))
                        { hil_addressline2 = entity.GetAttributeValue<string>("hil_addressline2"); }

                        string hil_alternatenumber = string.Empty;
                        if (entity.Attributes.Contains("hil_alternatenumber"))
                        { hil_alternatenumber = entity.GetAttributeValue<string>("hil_alternatenumber"); }

                        string hil_area = string.Empty;
                        if (entity.Attributes.Contains("hil_area"))
                        { hil_area = entity.GetAttributeValue<string>("hil_area"); }

                        OptionSetValue hil_callertype = entity.GetAttributeValue<OptionSetValue>("hil_callertype");
                        string hil_callsubtype = entity.GetAttributeValue<string>("hil_callsubtype");
                        string hil_customerfirstname = entity.GetAttributeValue<string>("hil_customerfirstname");
                        string hil_customerlastname = string.Empty;

                        if (entity.Attributes.Contains("hil_customerlastname"))
                        { hil_customerlastname = entity.GetAttributeValue<string>("hil_customerlastname"); }
                        string hil_customermobileno = entity.GetAttributeValue<string>("hil_customermobileno");
                        string hil_dealercode = entity.GetAttributeValue<string>("hil_dealercode");

                        if (entity.Attributes.Contains("hil_expecteddeliverydate"))
                        { hil_expecteddeliverydate = entity.GetAttributeValue<DateTime>("hil_expecteddeliverydate"); }

                        string hil_landmark = string.Empty;
                        if (entity.Attributes.Contains("hil_landmark"))
                        { hil_landmark = entity.GetAttributeValue<string>("hil_landmark"); }

                        string hil_natureofcomplaint = entity.GetAttributeValue<string>("hil_natureofcomplaint");
                        string hil_pincode = entity.GetAttributeValue<string>("hil_pincode");
                        string hil_productsubcategory = entity.GetAttributeValue<string>("hil_productsubcategory");
                        string hil_salutation = string.Empty;
                        if (entity.Attributes.Contains("hil_salutation"))
                        { hil_salutation = entity.GetAttributeValue<string>("hil_salutation"); }
                        #endregion

                        #region Validating Excel file data
                        if (hil_customermobileno == string.Empty || hil_customermobileno == null) { strInputsValidationsummary += "\n Customer Mobile Number is required."; inputsValidated = false; }
                        if (hil_customerfirstname == string.Empty || hil_customerfirstname == null) { strInputsValidationsummary += "\n Customer First Name is required."; inputsValidated = false; }
                        if (hil_addressline1 == string.Empty || hil_addressline1 == null) { strInputsValidationsummary += "\n Address Line 1 is required."; inputsValidated = false; }
                        if (hil_pincode == string.Empty || hil_pincode == null) { strInputsValidationsummary += "\n PIN Code is required."; inputsValidated = false; }
                        if (hil_productsubcategory == string.Empty || hil_productsubcategory == null) { strInputsValidationsummary += "\n Product Sub Category is required."; inputsValidated = false; }
                        if (hil_natureofcomplaint == string.Empty || hil_natureofcomplaint == null) { strInputsValidationsummary += "\n Nature of Complaint is required."; inputsValidated = false; }
                        if (hil_callsubtype == string.Empty || hil_callsubtype == null) { strInputsValidationsummary += "\n Call SubType is required."; inputsValidated = false; }
                        if (hil_callertype == null) { strInputsValidationsummary += "\n Caller Type is required."; inputsValidated = false; }

                        if (hil_callsubtype.ToUpper() == "INSTALLATION")
                        {
                            if (hil_expecteddeliverydate == null) { hil_expecteddeliverydate = DateTime.Now; }
                        }
                        #endregion

                        #region Declaring Local Variables
                        OptionSetValue salutation;
                        Entity entTemp;
                        Entity entConsumer = null;

                        QueryExpression queryExp;
                        EntityCollection entCol;
                        EntityReference erArea = null;
                        EntityReference erPinCode = null;
                        EntityReference erBizGeoMapping = null;
                        EntityReference erProductSubCategoryStagging = null;
                        EntityReference erproductSubCategory = null;
                        EntityReference erproductCategory = null;
                        EntityReference erCallSubType = null;
                        EntityReference erAddress = null;
                        EntityReference erNatureOfComplaint = null;
                        EntityReference erConsumerCategory = null;
                        EntityReference erConsumerType = null;
                        Guid _customerGuid = Guid.Empty;
                        bool _duplicateWO = false;
                        string _duplicateWOIds = string.Empty;
                        #endregion

                        #region Creating Consumer, Address, and Work Order Records
                        if (inputsValidated)
                        {
                            entTemp = ExecuteScalar(service, "hil_consumercategory", "hil_name", "End User", new string[] { "hil_consumercategoryid" }, ref entityUpdate);
                            if (entTemp != null)
                            { erConsumerCategory = entTemp.ToEntityReference(); }
                            entTemp = ExecuteScalar(service, "hil_consumertype", "hil_name", "B2C", new string[] { "hil_consumertypeid" }, ref entityUpdate);
                            if (entTemp != null)
                            { erConsumerType = entTemp.ToEntityReference(); }
                            if (hil_salutation.ToUpper() == "MR." || hil_salutation.ToUpper() == "MR") { salutation = new OptionSetValue(2); }
                            else if (hil_salutation.ToUpper() == "MISS." || hil_salutation.ToUpper() == "MISS") { salutation = new OptionSetValue(1); }
                            else if (hil_salutation.ToUpper() == "MRS." || hil_salutation.ToUpper() == "MRS") { salutation = new OptionSetValue(3); }
                            else if (hil_salutation.ToUpper() == "DR." || hil_salutation.ToUpper() == "DR") { salutation = new OptionSetValue(4); }
                            else if (hil_salutation.ToUpper() == "M/S." || hil_salutation.ToUpper() == "M/S") { salutation = new OptionSetValue(5); }
                            else { salutation = new OptionSetValue(2); }

                            queryExp = new QueryExpression("contact");
                            queryExp.ColumnSet = new ColumnSet("fullname", "emailaddress1", "contactid");
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, hil_customermobileno));
                            entCol = service.RetrieveMultiple(queryExp);

                            if (entCol.Entities.Count == 1)
                            {
                                _customerGuid = entCol.Entities[0].Id;
                                entTemp = ExecuteScalar(service, "hil_stagingdivisonmaterialgroupmapping", "hil_name", hil_productsubcategory, new string[] { "hil_productcategorydivision", "hil_productsubcategorymg" }, ref entityUpdate);
                                if (entTemp != null)
                                {
                                    erproductCategory = entTemp.GetAttributeValue<EntityReference>("hil_productcategorydivision");
                                    #region Duplicate Work Order Check {Same Customer/Prod Category and Open Job in last 30 days}
                                    queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                                    queryExp.ColumnSet = new ColumnSet("msdyn_name");
                                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp.Criteria.AddCondition(new ConditionExpression("hil_customerref", ConditionOperator.Equal, _customerGuid));
                                    queryExp.Criteria.AddCondition(new ConditionExpression("hil_productcategory", ConditionOperator.Equal, erproductCategory.Id));
                                    queryExp.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.LastXDays, 30));
                                    queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_substatus", ConditionOperator.NotIn, new object[] { new Guid("1527FA6C-FA0F-E911-A94E-000D3AF060A1"), new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"), new Guid("6C8F2123-5106-EA11-A811-000D3AF057DD"), new Guid("2927FA6C-FA0F-E911-A94E-000D3AF060A1"), new Guid("7E85074C-9C54-E911-A951-000D3AF0677F") }));
                                    queryExp.AddOrder("createdon", OrderType.Descending);
                                    entCol = service.RetrieveMultiple(queryExp);
                                    if (entCol.Entities.Count > 0)
                                    {
                                        _duplicateWO = true;
                                        foreach (Entity entTempDup in entCol.Entities)
                                        {
                                            _duplicateWOIds += entTempDup.GetAttributeValue<string>("msdyn_name") + ",";
                                        }
                                    }
                                    #endregion
                                }
                            }
                            else
                            {
                                string[] fullName = hil_customerfirstname.Split(' ');
                                if (fullName.Length > 1)
                                {
                                    hil_customerfirstname = fullName[0];
                                    hil_customerlastname = fullName[1];
                                }

                                try
                                {
                                    Contact cnt = new Contact()
                                    {
                                        FirstName = hil_customerfirstname,
                                        LastName = hil_customerlastname,
                                        MobilePhone = hil_customermobileno,
                                        hil_Salutation = salutation
                                    };
                                    cnt["hil_consumersource"] = new OptionSetValue(9);//Excel Upload
                                    _customerGuid = service.Create(cnt);
                                }
                                catch (Exception ex)
                                {
                                    entityUpdate["hil_description"] = "Error While Creating Customer!!! " + ex.Message;
                                    entityUpdate["hil_jobstatus"] = false;
                                }
                            }
                            if (!_duplicateWO)
                            {
                                if (_customerGuid != Guid.Empty)
                                {
                                    entConsumer = service.Retrieve("contact", _customerGuid, new ColumnSet("fullname", "emailaddress1", "contactid"));
                                    Entity entPincode = ExecuteScalar(service, "hil_pincode", "hil_name", hil_pincode, new string[] { "hil_pincodeid" }, ref entityUpdate);
                                    if (entPincode != null)
                                    {
                                        erPinCode = entPincode.ToEntityReference();
                                    }
                                    erAddress = GetCustomerAddress(service, _customerGuid, erPinCode != null ? erPinCode.Id : Guid.Empty);

                                    if (erAddress == null)
                                    {
                                        //entTemp = ExecuteScalar(service, "hil_area", "hil_name", hil_area, new string[] { "hil_areaid" }, ref entityUpdate);
                                        //if (entTemp != null) { erArea = entTemp.ToEntityReference(); } else { addressRemarks = "Area does not exist in Area Master."; }
                                        //entTemp = ExecuteScalar(service, "hil_pincode", "hil_name", hil_pincode, new string[] { "hil_pincodeid" }, ref entity);
                                        if (entPincode != null)
                                        {
                                            //erPinCode = entTemp.ToEntityReference();
                                            queryExp = new QueryExpression("hil_businessmapping");
                                            queryExp.ColumnSet = new ColumnSet("hil_businessmappingid");
                                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, erPinCode.Id));
                                            queryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                                            if (erArea != null)
                                            {
                                                queryExp.Criteria.AddCondition(new ConditionExpression("hil_area", ConditionOperator.Equal, erArea.Id));
                                            }
                                            queryExp.TopCount = 1;
                                            entCol = service.RetrieveMultiple(queryExp);
                                            if (entCol.Entities.Count >= 1)
                                            {
                                                erBizGeoMapping = entCol.Entities[0].ToEntityReference();
                                                try
                                                {
                                                    hil_address addr = new hil_address()
                                                    {
                                                        hil_Street1 = hil_addressline1,
                                                        hil_Street2 = hil_addressline2,
                                                        hil_Street3 = hil_landmark,
                                                        hil_AddressType = new OptionSetValue(1),
                                                        hil_BusinessGeo = erBizGeoMapping,
                                                        hil_Customer = new EntityReference("contact", entConsumer.Id)
                                                    };
                                                    Guid _addressGuid = service.Create(addr);
                                                    erAddress = new EntityReference("hil_address", _addressGuid);
                                                }
                                                catch (Exception ex)
                                                {
                                                    entityUpdate["hil_description"] = "Error While Creating Consumer's Address." + ex.Message;
                                                    entityUpdate["hil_jobstatus"] = false;
                                                }
                                            }
                                            else
                                            {
                                                entityUpdate["hil_description"] = "No Business Geo Mapping found for PIN Code.";
                                                entityUpdate["hil_jobstatus"] = false;
                                            }
                                        }
                                        else
                                        {
                                            entityUpdate["hil_description"] = "PIN Code does not exist.";
                                            entityUpdate["hil_jobstatus"] = false;
                                        }
                                    }
                                }

                                if (erAddress == null)
                                {
                                    entTemp = ExecuteScalar(service, "hil_pincode", "hil_name", hil_pincode, new string[] { "hil_pincodeid" }, ref entityUpdate);
                                    if (entTemp != null)
                                    {
                                        erPinCode = entTemp.ToEntityReference();
                                        queryExp = new QueryExpression("hil_businessmapping");
                                        queryExp.ColumnSet = new ColumnSet("hil_businessmappingid");
                                        queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                        queryExp.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, erPinCode.Id));
                                        if (erArea != null)
                                        {
                                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_area", ConditionOperator.Equal, erArea.Id));
                                        }
                                        queryExp.TopCount = 1;
                                        entCol = service.RetrieveMultiple(queryExp);
                                        if (entCol.Entities.Count >= 1)
                                        {
                                            erBizGeoMapping = entCol.Entities[0].ToEntityReference();
                                            try
                                            {
                                                hil_address addr = new hil_address()
                                                {
                                                    hil_Street1 = hil_addressline1,
                                                    hil_Street2 = hil_addressline2,
                                                    hil_Street3 = hil_landmark,
                                                    hil_AddressType = new OptionSetValue(1),
                                                    hil_BusinessGeo = erBizGeoMapping,
                                                    hil_Customer = new EntityReference("contact", entConsumer.Id)
                                                };
                                                Guid _addressGuid = service.Create(addr);
                                                erAddress = new EntityReference("hil_address", _addressGuid);
                                            }
                                            catch (Exception ex)
                                            {
                                                entityUpdate["hil_description"] = "Error While Creating Consumer's Address." + ex.Message;
                                                entityUpdate["hil_jobstatus"] = false;
                                            }
                                        }
                                        else
                                        {
                                            entityUpdate["hil_description"] = "No Business Geo Mapping found for PIN Code.";
                                            entityUpdate["hil_jobstatus"] = false;
                                        }
                                    }
                                }
                                if (entConsumer != null && erAddress != null)
                                {
                                    entTemp = ExecuteScalar(service, "hil_stagingdivisonmaterialgroupmapping", "hil_name", hil_productsubcategory, new string[] { "hil_productcategorydivision", "hil_productsubcategorymg" }, ref entityUpdate);
                                    if (entTemp != null)
                                    {
                                        erProductSubCategoryStagging = entTemp.ToEntityReference();
                                        erNatureOfComplaint = GetNatureOfComplaint(service, hil_natureofcomplaint, entTemp.GetAttributeValue<EntityReference>("hil_productsubcategorymg").Id);
                                        erproductCategory = entTemp.GetAttributeValue<EntityReference>("hil_productcategorydivision");
                                        erproductSubCategory = entTemp.GetAttributeValue<EntityReference>("hil_productsubcategorymg");

                                        entTemp = ExecuteScalar(service, "hil_callsubtype", "hil_name", hil_callsubtype, new string[] { "hil_callsubtypeid" }, ref entityUpdate);
                                        if (entTemp != null)
                                        {
                                            erCallSubType = entTemp.ToEntityReference();
                                        }
                                        msdyn_workorder enPMSWorkorder = new msdyn_workorder();
                                        enPMSWorkorder.hil_CustomerRef = entConsumer.ToEntityReference();
                                        enPMSWorkorder.hil_customername = entConsumer.GetAttributeValue<string>("fullname");
                                        enPMSWorkorder.hil_mobilenumber = hil_customermobileno;
                                        enPMSWorkorder.hil_Alternate = hil_alternatenumber;
                                        enPMSWorkorder.hil_Address = erAddress;
                                        enPMSWorkorder.hil_Productcategory = erproductCategory;
                                        enPMSWorkorder.hil_ProductCatSubCatMapping = erProductSubCategoryStagging;
                                        if (erConsumerType != null)
                                        {
                                            enPMSWorkorder["hil_consumertype"] = erConsumerType;
                                        }
                                        if (erConsumerCategory != null)
                                        {
                                            enPMSWorkorder["hil_consumercategory"] = erConsumerCategory;
                                        }
                                        if (erNatureOfComplaint != null)
                                        {
                                            enPMSWorkorder.hil_natureofcomplaint = erNatureOfComplaint;
                                        }
                                        enPMSWorkorder["hil_newserialnumber"] = hil_dealercode;
                                        if (erCallSubType != null)
                                        {
                                            enPMSWorkorder.hil_CallSubType = erCallSubType;
                                        }
                                        enPMSWorkorder.hil_quantity = 1;
                                        enPMSWorkorder["hil_callertype"] = hil_callertype;
                                        if (hil_expecteddeliverydate != null)
                                        {
                                            enPMSWorkorder["hil_expecteddeliverydate"] = hil_expecteddeliverydate;
                                        }
                                        enPMSWorkorder.hil_SourceofJob = new OptionSetValue(14); // {SourceofJob:"Excel Upload"}
                                        enPMSWorkorder.msdyn_ServiceAccount = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                                        enPMSWorkorder.msdyn_BillingAccount = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
                                        jobId = service.Create(enPMSWorkorder);

                                        entityUpdate["hil_jobid"] = new EntityReference("msdyn_workorder", jobId);
                                        entityUpdate["hil_jobstatus"] = true;
                                        entityUpdate["hil_description"] = "Job created successfully.";
                                    }
                                    else
                                    {
                                        entityUpdate["hil_description"] = "Product Sub Category Configuration (MG4) is missing.";
                                        entityUpdate["hil_jobstatus"] = false;
                                    }
                                }
                                else
                                {
                                    if (entConsumer == null)
                                    {
                                        entityUpdate["hil_description"] = "Customer does not exist.";
                                        entityUpdate["hil_jobstatus"] = false;
                                    }
                                    else if (erAddress == null)
                                    {
                                        entityUpdate["hil_description"] = addressRemarks + " Customer Address does not exist.";
                                        entityUpdate["hil_jobstatus"] = false;
                                    }
                                }
                            }
                            else
                            {
                                entityUpdate["hil_description"] = "Duplicate Job. Old Job# " + _duplicateWOIds;
                                entityUpdate["hil_jobstatus"] = false;
                            }
                        }
                        else
                        {
                            entityUpdate["hil_description"] = strInputsValidationsummary;
                            entityUpdate["hil_jobstatus"] = false;
                        }
                        #endregion

                    }
                    catch (Exception ex)
                    {
                        entityUpdate["hil_description"] = ex.Message;
                        entityUpdate["hil_jobstatus"] = false;
                    }
                    finally { service.Update(entityUpdate); }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private EntityReference GetCustomerAddress(IOrganizationService service, Guid customerGuid, Guid _pincode)
        {
            EntityReference retValue = null;
            QueryExpression Query;
            EntityCollection enCol;
            try
            {
                LinkEntity _lnkEntity = new LinkEntity()
                {
                    LinkFromEntityName = "msdyn_workorder",
                    LinkToEntityName = "hil_address",
                    LinkFromAttributeName = "hil_address",
                    LinkToAttributeName = "hil_addressid",
                    Columns = new ColumnSet(false),
                };
                _lnkEntity.LinkCriteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, _pincode));

                Query = new QueryExpression("msdyn_workorder")
                {
                    ColumnSet = new ColumnSet("hil_address"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerref", ConditionOperator.Equal, customerGuid));
                Query.TopCount = 1;
                Query.AddOrder("createdon", OrderType.Descending);
                if (_pincode != Guid.Empty)
                {
                    Query.LinkEntities.Add(_lnkEntity);
                }
                enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    retValue = enCol.Entities[0].GetAttributeValue<EntityReference>("hil_address");
                }
                else
                {
                    Query = new QueryExpression("hil_address")
                    {
                        ColumnSet = new ColumnSet("hil_fulladdress"),
                        Criteria = new FilterExpression(LogicalOperator.And)
                    };
                    Query.Criteria.AddCondition(new ConditionExpression("hil_customer", ConditionOperator.Equal, customerGuid));
                    Query.TopCount = 1;
                    Query.AddOrder("createdon", OrderType.Descending);
                    enCol = service.RetrieveMultiple(Query);
                    if (enCol.Entities.Count > 0)
                    {
                        retValue = enCol.Entities[0].ToEntityReference();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            return retValue;
        }
        private EntityReference GetNatureOfComplaint(IOrganizationService service, string noc, Guid relatedProduct)
        {
            EntityReference retValue = null;
            QueryExpression Query;
            EntityCollection enCol;
            try
            {
                Query = new QueryExpression("hil_natureofcomplaint")
                {
                    ColumnSet = new ColumnSet("hil_natureofcomplaintid"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                Query.Criteria.AddCondition(new ConditionExpression("hil_relatedproduct", ConditionOperator.Equal, relatedProduct));
                Query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, noc));
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Query.TopCount = 1;
                enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    retValue = enCol.Entities[0].ToEntityReference();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            return retValue;
        }
        private Entity ExecuteScalar(IOrganizationService service, string entityName, string primaryField, string primaryFieldValue, string[] columns, ref Entity entity)
        {
            Entity retEntity = null;
            try
            {
                QueryExpression Query = new QueryExpression(entityName);
                Query.ColumnSet = new ColumnSet(columns);
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression(primaryField, ConditionOperator.Equal, primaryFieldValue));
                Query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                EntityCollection enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count >= 1)
                {
                    retEntity = enCol.Entities[0];
                }
            }
            catch (Exception ex)
            {
                entity["hil_description"] = ex.Message;
                entity["hil_pmsjobstatus"] = false;
            }
            return retEntity;
        }
        private bool GetCallSubTypeByNOC(IOrganizationService service, Guid callSubTypeGuid, Guid nocGuid)
        {
            bool retValue = false;
            QueryExpression Query;
            EntityCollection enCol;
            try
            {
                Query = new QueryExpression("hil_natureofcomplaint")
                {
                    ColumnSet = new ColumnSet("hil_natureofcomplaintid"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                Query.Criteria.AddCondition(new ConditionExpression("hil_natureofcomplaintid", ConditionOperator.Equal, nocGuid));
                Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, callSubTypeGuid));
                Query.TopCount = 1;
                enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    retValue = true;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            return retValue;
        }
    }
}
