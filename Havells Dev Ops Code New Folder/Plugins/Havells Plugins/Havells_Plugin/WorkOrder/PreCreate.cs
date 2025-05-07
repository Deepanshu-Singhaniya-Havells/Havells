using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Text.RegularExpressions;

namespace Havells_Plugin.WorkOrder
{
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    SetAutoNumber(entity, service);
                    msdyn_workorder Job = entity.ToEntity<msdyn_workorder>();
                    JobDefaultValues(service, Job);
                    RestrictDuplicateJobs(service, entity);
                    ChecksAndValidations(service, entity);
                    SetJobPriority(service, entity);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.WorkOrder_Pre_Create.SetAutoNumber: " + ex.Message);
            }
            #endregion
        }
        public static void SetAutoNumber(Entity _wOd, IOrganizationService service)
        {
            try
            {
                msdyn_workorder iJob = _wOd.ToEntity<msdyn_workorder>();
                string sSystemAutoNumber = string.Empty;
                if (_wOd.Attributes.Contains("msdyn_name"))
                {
                    sSystemAutoNumber = (string)_wOd["msdyn_name"];
                    // sSystemAutoNumber = GetNextAutoNumber(service);
                }
                DateTime iSysTime = DateTime.Now.AddMinutes(330);
                string Date = string.Empty;
                string Month = string.Empty;
                string Year = string.Empty;
                if (iSysTime.Day < 10)
                    Date = "0" + iSysTime.Day.ToString();
                else
                    Date = iSysTime.Day.ToString();
                if (iSysTime.Month < 10)
                    Month = "0" + iSysTime.Month.ToString();
                else
                    Month = iSysTime.Month.ToString();
                int iYear = iSysTime.Year % 100;
                String sPreFix = Date.ToString() + Month.ToString() + iYear.ToString();//getAutoNumberPreFix();
                _wOd["msdyn_name"] = sPreFix + sSystemAutoNumber;
                Guid ServiceAccount = Helper.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
                Guid PriceList = Helper.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
                _wOd["msdyn_serviceaccount"] = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                _wOd["msdyn_billingaccount"] = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                if (PriceList != Guid.Empty)
                {
                    _wOd["msdyn_pricelist"] = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
                }
                if (_wOd.Contains("hil_callsubtype") && iJob.hil_CallSubType != null)
                {
                    EntityReference CallSubType = (EntityReference)_wOd["hil_callsubtype"];
                    EntityReference CallType = new EntityReference(hil_calltype.EntityLogicalName);
                    hil_callsubtype ClSb = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, CallSubType.Id, new ColumnSet("hil_calltype"));
                    if (ClSb.hil_CallType != null)
                    {
                        CallType = ClSb.hil_CallType;
                        _wOd["hil_calltype"] = CallType;
                    }
                }
                //if (_wOd.Attributes.Contains("hil_email") && _wOd.Attributes.Contains("hil_callsubtype") &&
                //    _wOd.Attributes.Contains("hil_productcategory"))
                //{
                //    CreateWOForPortal(service, _wOd);
                //}
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.WorkOrder_Pre_Create.SetAutoNumber: " + ex.Message);
            }
        }
        public static String getAutoNumberPreFix()
        {
            String PreFix = String.Empty;
            try
            {

                DateTime dt = DateTime.Now.ToLocalTime();
                Int32 tempyear = dt.Year - 2010;

                // String temp = dt.Year.ToString();
                String temp = tempyear.ToString();
                //Year
                //  temp = temp.Substring(temp.Length - 1, 1);
                int unicode = Convert.ToInt16(temp) + 64;
                char character = (char)unicode;
                PreFix = character.ToString();
                //Month
                temp = dt.Month.ToString();
                unicode = Convert.ToInt16(temp) + 64;
                character = (char)unicode;
                String sDayofMonth = dt.Day.ToString();
                if (sDayofMonth.Length == 1) sDayofMonth = "0" + sDayofMonth;
                PreFix = PreFix + character.ToString() + sDayofMonth;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.WorkOrder_Pre_Create.getAutoNumberPreFix: " + ex.Message);
            }
            return PreFix;
        }
        public static void CreateWOForPortal(IOrganizationService service, Entity _WrkOd)
        {
            Guid ContId = Helper.GetGuidbyName(Contact.EntityLogicalName, "emailaddress1", Convert.ToString(_WrkOd["hil_email"]), service);
            Guid ServiceAccount = Helper.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
            Guid PriceList = Helper.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
            Guid Nature = Helper.GetGuidbyName(hil_natureofcomplaint.EntityLogicalName, "hil_name", "Service", service);
            Guid IncidentType = Helper.GetGuidbyName(msdyn_incidenttype.EntityLogicalName, "msdyn_name", "Service", service);
            EntityReference CallSubType = (EntityReference)_WrkOd["hil_callsubtype"];
            EntityReference CallType = new EntityReference(hil_calltype.EntityLogicalName);
            hil_callsubtype ClSb = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, CallSubType.Id, new ColumnSet(true));
            if (ClSb.hil_CallType != null)
            {
                CallType = ClSb.hil_CallType;
                _WrkOd["hil_calltype"] = CallType;
            }
            _WrkOd["hil_customerref"] = new EntityReference(Contact.EntityLogicalName, ContId);
            _WrkOd["msdyn_serviceaccount"] = new EntityReference(Account.EntityLogicalName, ServiceAccount);
            _WrkOd["msdyn_billingaccount"] = new EntityReference(Account.EntityLogicalName, ServiceAccount);
            if (PriceList != Guid.Empty)
            {
                _WrkOd["msdyn_pricelist"] = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
            }
            if (Nature != Guid.Empty)
            {
                _WrkOd["hil_natureofcomplaint"] = new EntityReference(hil_natureofcomplaint.EntityLogicalName, Nature);
            }
            if (IncidentType != Guid.Empty)
            {
                _WrkOd["msdyn_primaryincidenttype"] = new EntityReference(msdyn_incidenttype.EntityLogicalName, IncidentType);
            }
        }

        public static String GetNextAutoNumber(IOrganizationService service)
        {
            try
            {
                String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <attribute name='createdon' />
                    <attribute name='msdyn_workorderid' />
                    <order attribute='msdyn_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='createdon' operator='today' />
                    </filter>
                  </entity>
                </fetch>";
                EntityCollection enColl = service.RetrieveMultiple(new FetchExpression(fetch));
                Int32 nextNumber = enColl.Entities.Count + 1;
                String snextNumber = nextNumber.ToString();
                while (snextNumber.Length != 4)
                {
                    snextNumber = "0" + snextNumber;
                }

                return snextNumber;

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.WorkOrderPreCreate.SetAutoNumber" + ex.Message);
            }

        }
        #region Set Default Value Job
        public static void JobDefaultValues(IOrganizationService service, msdyn_workorder Job)
        {
            if (Job.hil_quantity == null || Job.hil_quantity == 0)
            {
                throw new InvalidPluginExecutionException("Job Quantity is Mandatory and should be greater than 0");
            }

            if (Job.hil_mobilenumber == null || Job.hil_mobilenumber.Length != 10 || !Regex.IsMatch(Job.hil_mobilenumber, "[6-9]([0-9]){9}"))
            {
                throw new InvalidPluginExecutionException("Mobile Number is Mandatory and should have 10 digits starts with 6,7,8 or 9");
            }

            //----------------------- Kuldeep Khare 06/Sep/2019--------------------------------
            //Validation Removed to faciliate Job Creation by SMS
            //---------------------------------------------------------------------------------
            if (Job.hil_Productcategory != null)
            {
                if (Job.hil_Productcategory == null || Job.hil_ProductCatSubCatMapping == null)
                {
                    throw new InvalidPluginExecutionException("Product Category and Product Sub Category must be selected");
                }

                hil_stagingdivisonmaterialgroupmapping sdmMapping = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(
                hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, Job.GetAttributeValue<EntityReference>("hil_productcatsubcatmapping").Id, new ColumnSet(new string[] { "hil_productsubcategorymg", "hil_productcategorydivision" }));

                if (sdmMapping.hil_ProductSubCategoryMG == null)
                {
                    throw new InvalidPluginExecutionException("Product Sub Category cannot be Empty. Please contact to Administrator");
                }

                //if (Job.hil_SourceofJob.Value != 12)
                //{
                //    if (sdmMapping.hil_ProductCategoryDivision == null || Job.hil_Productcategory.Id != sdmMapping.hil_ProductCategoryDivision.Id)
                //    {
                //        throw new InvalidPluginExecutionException("Product Sub Category does not belong to respective Product Category");
                //    }
                //}
                Product prod_Category = (Product)service.Retrieve(Product.EntityLogicalName, sdmMapping.hil_ProductSubCategoryMG.Id, new ColumnSet(new string[] { "hil_isserialized", "hil_minimumthreshold", "hil_maximumthreshold" }));

                Job.hil_MinQuantity = prod_Category.hil_MinimumThreshold != null ? prod_Category.hil_MinimumThreshold : 1;
                Job.hil_MaxQuantity = prod_Category.hil_MaximumThreshold != null ? prod_Category.hil_MaximumThreshold : 1;
            }
            else
            {
                Job.hil_MinQuantity = 1;
                Job.hil_MaxQuantity = 1;
            }

            if (Job.hil_Address == null || Job.hil_pincode == null || Job.hil_Branch == null || !Job.Attributes.Contains("hil_salesoffice"))
            {

                if (Job.hil_CustomerRef != null)
                {
                    QueryExpression Query = new QueryExpression(hil_address.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet(new string[] { "hil_street1", "hil_street2", "hil_state", "hil_city", "hil_district", "hil_area", "hil_pincode" });
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    if (Job.hil_Address != null)
                    {
                        Query.Criteria.AddCondition(new ConditionExpression("hil_addressid", ConditionOperator.Equal, Job.hil_Address.Id));
                    }
                    else
                    {
                        Query.Criteria.AddCondition(new ConditionExpression("hil_customer", ConditionOperator.Equal, Job.hil_CustomerRef.Id));
                        Query.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.NotNull));
                    }
                    Query.AddOrder("createdon", OrderType.Descending);
                    EntityCollection Found = service.RetrieveMultiple(Query);
                    if (Found.Entities.Count > 0)
                    {
                        hil_address Add = Found.Entities[0].ToEntity<hil_address>();
                        Job.hil_Address = new EntityReference(hil_address.EntityLogicalName, Add.Id);
                        Job.hil_pincode = Add.hil_PinCode;
                        Job.hil_PinCodeText = Add.hil_PinCode.Name;
                        QueryByAttribute Qry = new QueryByAttribute();
                        Qry.EntityName = hil_businessmapping.EntityLogicalName;
                        ColumnSet Col = new ColumnSet("hil_district", "hil_area", "hil_region", "hil_branch", "hil_salesoffice", "hil_city", "hil_state");
                        Qry.ColumnSet = Col;
                        Qry.AddAttributeValue("hil_pincode", Add.hil_PinCode.Id);
                        if (Add.hil_Area != null)
                        {
                            Qry.AddAttributeValue("hil_area", Add.hil_Area.Id);
                        }
                        Qry.AddAttributeValue("statuscode", 1); //Active Record

                        EntityCollection Found1 = service.RetrieveMultiple(Qry);
                        if (Found1.Entities.Count > 0)
                        {
                            foreach (hil_businessmapping Bus in Found1.Entities)
                            {
                                if (Bus.hil_branch != null && Job.hil_Branch == null)
                                {
                                    Job.hil_Branch = Bus.hil_branch;
                                }
                                if (Bus.hil_region != null)
                                {
                                    Job.hil_Region = Bus.hil_region;
                                }
                                if (Bus.hil_area != null)
                                {
                                    Job["hil_area"] = Bus.hil_area;
                                }
                                if (Bus.hil_district != null)
                                {
                                    Job["hil_district"] = Bus.hil_district;
                                }
                                if (Bus.hil_salesoffice != null && !Job.Attributes.Contains("hil_salesoffice"))
                                {
                                    Job["hil_salesoffice"] = Bus.hil_salesoffice;
                                }
                                if (Bus.hil_state != null)
                                {
                                    Job.hil_state = Bus.hil_state;
                                    Job.hil_StateText = Bus.hil_state.Name;
                                }
                                if (Bus.hil_city != null)
                                {
                                    Job.hil_City = Bus.hil_city;
                                    Job.hil_CityText = Bus.hil_city.Name;
                                }
                                //if (Add.hil_FullAddress != null)
                                //{
                                //    Job.hil_FullAddress = Add.hil_FullAddress;
                                //}
                                //else
                                {
                                    string Consumer_address = Add.hil_Street1 != null ? Add.hil_Street1 + ", " : "";

                                    if (Add.hil_Street2 != null)
                                        Consumer_address = Consumer_address + Add.hil_Street2 + ", ";
                                    if (Add.hil_Area != null)
                                        Consumer_address = Consumer_address + Add.hil_Area.Name + ", ";
                                    if (Add.hil_District != null)
                                        Consumer_address = Consumer_address + Add.hil_District.Name + ", ";
                                    if (Add.hil_CIty != null)
                                        Consumer_address = Consumer_address + Add.hil_CIty.Name + ", ";
                                    if (Add.hil_State != null)
                                        Consumer_address = Consumer_address + Add.hil_State.Name + ", ";
                                    if (Add.hil_PinCode != null)
                                        Consumer_address = Consumer_address + Add.hil_PinCode.Name + " ";

                                    if (Consumer_address.Length > 0)
                                    {

                                    }
                                    Job.hil_FullAddress = Consumer_address;
                                }
                            }
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - NO ADDRESS FOUND FOR THIS CUSTOMER - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                }
                else
                {
                    throw new InvalidPluginExecutionException("Customer can't be blank");
                }
            }
            else if (Job.hil_Address != null)
            {
                hil_address address = (hil_address)service.Retrieve(hil_address.EntityLogicalName, Job.hil_Address.Id, new ColumnSet(
                    new string[] { "hil_street1", "hil_street2", "hil_state", "hil_city", "hil_district", "hil_area", "hil_pincode", "hil_businessgeo" }));
                if (address.hil_BusinessGeo != null)
                {
                    QueryByAttribute Qry = new QueryByAttribute();
                    Qry.EntityName = hil_businessmapping.EntityLogicalName;
                    ColumnSet Col = new ColumnSet("hil_district", "hil_area", "hil_region", "hil_branch", "hil_salesoffice");
                    Qry.ColumnSet = Col;
                    //Qry.AddAttributeValue("hil_businessmappingid", address.hil_BusinessGeo.Id);
                    Qry.AddAttributeValue("hil_pincode", address.hil_PinCode.Id);
                    if (address.hil_Area != null)
                    {
                        Qry.AddAttributeValue("hil_area", address.hil_Area.Id);
                    }
                    Qry.AddAttributeValue("statuscode", 1); //Active Record

                    EntityCollection Found = service.RetrieveMultiple(Qry);
                    if (Found.Entities.Count > 0)
                    {
                        foreach (hil_businessmapping Bus in Found.Entities)
                        {
                            if (Bus.hil_branch != null)
                            {
                                Job["hil_branch"] = Bus.hil_branch;
                            }
                            if (Bus.hil_area != null)
                            {
                                Job["hil_area"] = Bus.hil_area;
                            }
                            if (Bus.hil_district != null)
                            {
                                Job["hil_district"] = Bus.hil_district;
                            }
                            if (Bus.hil_salesoffice != null)
                            {
                                Job["hil_salesoffice"] = Bus.hil_salesoffice;
                            }
                        }
                    }
                }
                else
                {
                    //if (Job.hil_FullAddress == null)
                    {
                        string add = address.hil_Street1 != null ? address.hil_Street1 + ", " : "";

                        if (address.hil_Street2 != null)
                            add = add + address.hil_Street2 + ", ";
                        if (address.hil_Area != null)
                            add = add + address.hil_Area.Name + ", ";
                        if (address.hil_District != null)
                            add = add + address.hil_District.Name + ", ";
                        if (address.hil_CIty != null)
                            add = add + address.hil_CIty.Name + ", ";
                        if (address.hil_State != null)
                            add = add + address.hil_State.Name + ", ";
                        if (address.hil_PinCode != null)
                            add = add + address.hil_PinCode.Name + " ";

                        Job.hil_FullAddress = add;
                    }
                    if (address.hil_SalesOffice != null)
                    {
                        Job["hil_salesoffice"] = address.hil_SalesOffice;
                    }
                }
            }
            //if (Job.hil_SourceofJob != null)
            //{
            //    if (Job.hil_SourceofJob.Value == 3 && Job.hil_Address != null)
            //    {
            //        hil_address iAddress = (hil_address)service.Retrieve(hil_address.EntityLogicalName, Job.hil_Address.Id, new ColumnSet("hil_salesoffice", "hil_branch"));
            //        if (iAddress.hil_SalesOffice != null)
            //        {
            //            Job["hil_salesoffice"] = iAddress.hil_SalesOffice;
            //        }
            //    }
            //    //if (Job.hil_SourceofJob.Value == 4 && Job.hil_Address != null)
            //    //{
            //    //    hil_address iAddress = (hil_address)service.Retrieve(hil_address.EntityLogicalName, Job.hil_Address.Id, new ColumnSet("hil_salesoffice", "hil_branch", "hil_fulladdress"));
            //    //
            //    //    if (iAddress.hil_FullAddress != null)
            //    //    {
            //    //        Job.hil_FullAddress = iAddress.hil_FullAddress;
            //    //    }
            //    //}
            //}
            if (Job.hil_ProductCatSubCatMapping != null)
            {
                if (Job.hil_Productcategory == null || Job.hil_ProductSubcategory == null)
                {
                    hil_stagingdivisonmaterialgroupmapping Map = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, Job.hil_ProductCatSubCatMapping.Id, new ColumnSet(true));
                    if (Map.hil_ProductCategoryDivision != null)
                    {
                        Job.hil_Productcategory = Map.hil_ProductCategoryDivision;
                        if (Map.hil_ProductSubCategory != null)
                            Job.hil_ProductSubcategory = Map.hil_ProductSubCategoryMG;
                        //service.Update(Job);
                    }
                }
            }
            //if(Job.hil_SourceofJob != null)
            //{
            //    if(Job.hil_SourceofJob.Value == 3 && Job.hil_CustomerRef != null)
            //    {
            //        QueryExpression Query = new QueryExpression(hil_address.EntityLogicalName);
            //    }
            //}
        }
        #endregion
        static void RestrictDuplicateJobs(IOrganizationService _service, Entity _jobEntity)
        {
            try
            {
                if (_jobEntity.Contains("hil_productcategory") && _jobEntity.Contains("hil_mobilenumber"))
                {
                    Guid _productDivisionId = _jobEntity.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                    string _mobileNumber = _jobEntity.GetAttributeValue<string>("hil_mobilenumber");

                    QueryExpression qryExp = new QueryExpression("hil_jobsquantitymatrix");
                    qryExp.ColumnSet = new ColumnSet("hil_quantity", "hil_frequency");
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_divisionname", ConditionOperator.Equal, _productDivisionId));
                    qryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                    if (entCol.Entities.Count > 0)
                    {
                        Entity _entObj = entCol.Entities[0];
                        if (_entObj != null)
                        {
                            int _frequency = _entObj.GetAttributeValue<int>("hil_frequency");
                            int _quantity = _entObj.GetAttributeValue<int>("hil_quantity");
                            qryExp = new QueryExpression("msdyn_workorder");
                            qryExp.ColumnSet = new ColumnSet(false);
                            qryExp.Criteria.AddCondition(new ConditionExpression("hil_productcategory", ConditionOperator.Equal, _productDivisionId));
                            qryExp.Criteria.AddCondition(new ConditionExpression("hil_mobilenumber", ConditionOperator.Equal, _mobileNumber));
                            qryExp.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.LastXHours, _frequency));
                            EntityCollection _entColObj = _service.RetrieveMultiple(qryExp);
                            if (_entColObj.Entities.Count >= _quantity)
                            {
                                throw new InvalidPluginExecutionException("Duplicacy Rule !!! Concurrent Jobs for Mobile Number# " + _mobileNumber);
                            }
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("Data Duplicacy Rule does not exist for Division " + _jobEntity.GetAttributeValue<EntityReference>("hil_productcategory").Name + " has been traced.");
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Data Duplicacy Rule does not exist for Division " + _jobEntity.GetAttributeValue<EntityReference>("hil_productcategory").Name + " has been traced.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.WorkOrder_Pre_Create.RestrictDuplicateJobs: " + ex.Message);
            }
        }
        static void ChecksAndValidations(IOrganizationService _service, Entity _jobEntity)
        {
            try
            {
                if (_jobEntity.Contains("hil_natureofcomplaint"))
                {
                    Guid _natureofcomplaintId = _jobEntity.GetAttributeValue<EntityReference>("hil_natureofcomplaint").Id;

                    QueryExpression qryExp = new QueryExpression("hil_natureofcomplaint");
                    qryExp.ColumnSet = new ColumnSet(false);
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_natureofcomplaintid", ConditionOperator.Equal, _natureofcomplaintId));
                    qryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                    if (entCol.Entities.Count == 0)
                    {
                        throw new InvalidPluginExecutionException("Nature of Complaint is Inactive.");
                    }
                }
                if (_jobEntity.Contains("hil_productsubcategory") && _jobEntity.Contains("hil_natureofcomplaint"))
                {
                    Guid _productsubCategoryId = _jobEntity.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
                    Guid _natureofcomplaintId = _jobEntity.GetAttributeValue<EntityReference>("hil_natureofcomplaint").Id;

                    QueryExpression qryExp = new QueryExpression("hil_natureofcomplaint");
                    qryExp.ColumnSet = new ColumnSet(false);
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_relatedproduct", ConditionOperator.Equal, _productsubCategoryId));
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_natureofcomplaintid", ConditionOperator.Equal, _natureofcomplaintId));
                    qryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                    if (entCol.Entities.Count == 0)
                    {
                        throw new InvalidPluginExecutionException("Nature of Complaint does not belong to Selected Product Subcategory.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.WorkOrder_Pre_Create.ChecksAndValidations: " + ex.Message);
            }
        }
        static void SetJobPriority(IOrganizationService _service, Entity _jobEntity)
        {
            try
            {
                if (_jobEntity.Contains("hil_customerref"))
                {
                    bool IsLoyaltyProgramEnabled = false;
                    Guid _customerId = ((EntityReference)_jobEntity.Attributes["hil_customerref"]).Id;
                    Entity entContact = _service.Retrieve("contact", _customerId, new ColumnSet("hil_loyaltyprogramenabled", "hil_loyaltyprogramtier"));
                    IsLoyaltyProgramEnabled = entContact.Contains("hil_loyaltyprogramenabled") ? entContact.GetAttributeValue<bool>("hil_loyaltyprogramenabled") : false;
                    int loyaltyTierValue = entContact.Contains("hil_loyaltyprogramtier") ? entContact.GetAttributeValue<OptionSetValue>("hil_loyaltyprogramtier").Value : 0;
                    //Check Loyalty Is enable or not if enable and Tire is 3 or Previlaged then set High prority jobs 
                    if (IsLoyaltyProgramEnabled && loyaltyTierValue == 3)
                    {
                        _jobEntity["hil_jobpriority"] = 0;
                        _jobEntity["hil_priorityindicator"] = DateTime.Now.AddHours(24).AddMinutes(330).ToString("yyyy-MM-dd HH:mm");
                        _jobEntity["hil_loyaltyprogramtier"] = new OptionSetValue(loyaltyTierValue);
                    }
                    else
                        _jobEntity["hil_jobpriority"] = 1;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.WorkOrder_Pre_Create.SetJobPriority: " + ex.Message);
            }
        }
    }
}