using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ServiceCallRequest
{
    public class BotServiceRequestJobCreation
    {
        public void CreateServiceBotJob(IOrganizationService service)
        {
            String ErrorDetails = string.Empty;
            bool _dataValidated = true;
            Guid consumerTypeGuid = new Guid("484897de-2abd-e911-a957-000d3af0677f");
            Guid consumerCategoryGuid = new Guid("baa52f5e-2bbd-e911-a957-000d3af0677f");

            string fetchServiceBot = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_servicerequestvoicebot'>
                    <attribute name='hil_servicerequestvoicebotid' />
                    <attribute name='hil_name' />
                    <attribute name='hil_syncresponse' />
                    <attribute name='statuscode' />
                    <attribute name='statecode' />
                    <attribute name='hil_sourceofrequest' />
                    <attribute name='hil_productcategory' />
                    <attribute name='hil_preferredlanguage' />
                    <attribute name='hil_pincode' />
                    <attribute name='hil_noc' />
                    <attribute name='hil_customername' />
                    <attribute name='hil_addressline1' />
                    <attribute name='hil_customeremail' />
                    <attribute name='hil_chiefcomplaint' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_syncresponse' operator='null' />
                    </filter>
                </entity>
                </fetch>";
            EntityCollection botCol = service.RetrieveMultiple(new FetchExpression(fetchServiceBot));
            if(botCol.Entities.Count > 0)
            {
                int i = 1;
                foreach (Entity EntBot in botCol.Entities)
                {
                    ErrorDetails = string.Empty;_dataValidated = true;
                    EntityReference _entPreferredLang = null;
                    Console.WriteLine("Processing Service Requests ..." + i++.ToString() + "/" + botCol.Entities.Count.ToString());
                    try
                    {
                        EntityReference CustId = null;
                        Guid Addressguid = new Guid();
                        EntityReference BusinesMapId = null;
                        EntityReference PinCodeId = null;
                        EntityReference DistrictId = null;
                        EntityReference ProductCat = null;
                        EntityReference ProductMat = null;
                        EntityReference StagingDivisonMaterialGroupMapping = null;
                        EntityReference CallSubType = null;
                        EntityReference NatureOfComplaint = null;
                        EntityReference AreaId = null;

                        string MobileNo = EntBot.GetAttributeValue<string>("hil_name");
                        string Pincode = EntBot.GetAttributeValue<string>("hil_pincode");
                        string CustomerName = EntBot.GetAttributeValue<string>("hil_customername");
                        string Noc = EntBot.GetAttributeValue<string>("hil_noc");
                        string customeremail = EntBot.GetAttributeValue<string>("hil_customeremail");
                        string chiefcomplaint = EntBot.GetAttributeValue<string>("hil_chiefcomplaint");
                        string preferredLanguage = EntBot.GetAttributeValue<string>("PreferredLanguage");
                        if (preferredLanguage == "1")
                        {
                            preferredLanguage = "en";
                        }
                        else if (preferredLanguage == "2")
                        {
                            preferredLanguage = "hi";
                        }
                        else {
                            preferredLanguage = "en";
                        }
                        QueryExpression qspreferLang = new QueryExpression("hil_preferredlanguageforcommunication");
                        qspreferLang.ColumnSet = new ColumnSet(false);
                        ConditionExpression condExp = new ConditionExpression("hil_code", ConditionOperator.Equal, preferredLanguage.ToString());
                        qspreferLang.Criteria.AddCondition(condExp);
                        EntityCollection entColPreferLang = service.RetrieveMultiple(qspreferLang);
                        if (entColPreferLang.Entities.Count > 0)
                        {
                            _entPreferredLang = entColPreferLang.Entities[0].ToEntityReference();
                        }

                        //Step 1
                        QueryExpression queryMobile = new QueryExpression("contact");
                        queryMobile.ColumnSet = new ColumnSet(false);
                        queryMobile.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, MobileNo));
                        EntityCollection entColCust = service.RetrieveMultiple(queryMobile);
                        if (entColCust.Entities.Count > 0)
                        {
                            CustId = entColCust.Entities[0].ToEntityReference();
                        }
                        else
                        {
                            Entity entConsumer = new Entity("contact");
                            entConsumer["firstname"] = CustomerName;
                            entConsumer["mobilephone"] = MobileNo;
                            entConsumer["emailaddress1"] = customeremail;
                            entConsumer["hil_consumersource"] = new OptionSetValue(11); // Voice Bot
                            if (_entPreferredLang != null)
                                entConsumer["hil_preferredlanguageforcommunication"] = _entPreferredLang;
                            Guid ConsumerId = service.Create(entConsumer);
                            CustId = new EntityReference("contact", ConsumerId);
                        }
                        //Step 2

                        QueryExpression queryPhone = new QueryExpression("hil_businessmapping");
                        queryPhone.ColumnSet = new ColumnSet("hil_pincode", "hil_district", "hil_area");
                        queryPhone.Criteria.AddCondition(new ConditionExpression("hil_pincodename", ConditionOperator.Like, Pincode));
                        queryPhone.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        queryPhone.AddOrder("modifiedon", OrderType.Descending);
                        queryPhone.TopCount = 1;
                        EntityCollection EntColBusMap = service.RetrieveMultiple(queryPhone);
                        if (EntColBusMap.Entities.Count > 0)
                        {
                            DistrictId = EntColBusMap.Entities[0].GetAttributeValue<EntityReference>("hil_district");
                            PinCodeId = EntColBusMap.Entities[0].GetAttributeValue<EntityReference>("hil_pincode");
                            AreaId = EntColBusMap.Entities[0].GetAttributeValue<EntityReference>("hil_area");
                            BusinesMapId = EntColBusMap.Entities[0].ToEntityReference();

                            Entity addres = new Entity("hil_address");
                            addres["hil_customer"] = CustId;
                            addres["hil_district"] = DistrictId;
                            addres["hil_businessgeo"] = BusinesMapId;
                            addres["hil_pincode"] = PinCodeId;
                            addres["hil_area"] = AreaId;
                            addres["hil_street1"] = EntBot.GetAttributeValue<string>("hil_addressline1");
                            addres["hil_addresstype"] = new OptionSetValue(1);
                            Addressguid = service.Create(addres);
                        }
                        else
                        {
                            ErrorDetails = ErrorDetails + "Pincode: Invalid PinCode ";
                            _dataValidated = false;
                        }

                        //Step 3
                        string ProductCateGuid = EntBot.GetAttributeValue<string>("hil_productcategory");
                        QueryExpression WhatsappProductConfig = new QueryExpression("hil_whatsappproductdivisionconfig");
                        WhatsappProductConfig.ColumnSet = new ColumnSet("hil_productmaterialgroup", "hil_productdivision");
                        WhatsappProductConfig.Criteria.AddCondition(new ConditionExpression("hil_productdivision", ConditionOperator.Equal, new Guid(ProductCateGuid)));
                        WhatsappProductConfig.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection entWhatsConfig = service.RetrieveMultiple(WhatsappProductConfig);

                        if (entWhatsConfig.Entities.Count > 0)
                        {
                            ProductCat = entWhatsConfig.Entities[0].GetAttributeValue<EntityReference>("hil_productdivision");
                            ProductMat = entWhatsConfig.Entities[0].GetAttributeValue<EntityReference>("hil_productmaterialgroup");

                            QueryExpression QueryProdCatSubCat = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                            QueryProdCatSubCat.ColumnSet = new ColumnSet("hil_productcategorydivision", "hil_productsubcategorymg");
                            QueryProdCatSubCat.Criteria = new FilterExpression(LogicalOperator.And);
                            QueryProdCatSubCat.Criteria.AddCondition(new ConditionExpression("hil_productcategorydivision", ConditionOperator.Equal, ProductCat.Id));
                            QueryProdCatSubCat.Criteria.AddCondition(new ConditionExpression("hil_productsubcategorymg", ConditionOperator.Equal, ProductMat.Id));
                            EntityCollection entProdCatSubCatg = service.RetrieveMultiple(QueryProdCatSubCat);

                            if (entProdCatSubCatg.Entities.Count > 0)
                            {
                                StagingDivisonMaterialGroupMapping = entProdCatSubCatg.Entities[0].ToEntityReference();
                            }
                            else
                            {
                                ErrorDetails = ErrorDetails + "Product Cat and Material Group : Product Cat or Material Group not found in staging division material group mapping ";
                                _dataValidated = false;
                            }
                        }
                        else
                        {
                            ErrorDetails = ErrorDetails + "ProductCat : Product Cat not found in  Whatsapp Product Division Mapping Entity ";
                            _dataValidated = false;
                        }

                        QueryExpression qruNatureOfComplaint = new QueryExpression("hil_natureofcomplaint");
                        qruNatureOfComplaint.ColumnSet = new ColumnSet("hil_callsubtype");
                        qruNatureOfComplaint.Criteria = new FilterExpression(LogicalOperator.And);
                        qruNatureOfComplaint.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, Noc));
                        EntityCollection entColNComplaint = service.RetrieveMultiple(qruNatureOfComplaint);
                        //CallSubType

                        if(entColNComplaint.Entities.Count > 0)
                        {
                            NatureOfComplaint = entColNComplaint.Entities[0].ToEntityReference();
                            CallSubType = entColNComplaint.Entities[0].GetAttributeValue<EntityReference>("hil_callsubtype");
                        }
                        else
                        {
                            ErrorDetails = ErrorDetails + "Noc: Noc not found ";
                            _dataValidated = false;
                        }
                        //Now Create Job
                        if (_dataValidated)
                        {
                            Entity enWorkorder = new Entity("msdyn_workorder");
                            enWorkorder["hil_customerref"] = new EntityReference("contact", CustId.Id);
                            enWorkorder["hil_customername"] = CustomerName;
                            enWorkorder["hil_mobilenumber"] = MobileNo;
                            enWorkorder["hil_email"] = customeremail;
                            enWorkorder["hil_address"] = new EntityReference("hil_address", Addressguid);
                            enWorkorder["hil_productcategory"] = new EntityReference("product", ProductCat.Id);
                            enWorkorder["hil_productsubcategory"] = new EntityReference("product", ProductMat.Id);
                            enWorkorder["hil_productcatsubcatmapping"] = new EntityReference("hil_stagingdivisonmaterialgroupmapping", StagingDivisonMaterialGroupMapping.Id);
                            enWorkorder["hil_natureofcomplaint"] = new EntityReference("hil_natureofcomplaint", NatureOfComplaint.Id);
                            enWorkorder["hil_consumertype"] = new EntityReference("hil_consumertype", consumerTypeGuid);
                            enWorkorder["hil_consumercategory"] = new EntityReference("hil_consumercategory", consumerCategoryGuid);

                            enWorkorder["hil_callsubtype"] = new EntityReference("hil_callsubtype", CallSubType.Id);

                            enWorkorder["hil_quantity"] = 1;
                            enWorkorder["hil_customercomplaintdescription"] = chiefcomplaint;
                            enWorkorder["hil_callertype"] = new OptionSetValue(910590001);// {SourceofJob:"Customer"}

                            enWorkorder["hil_sourceofjob"] = new OptionSetValue(18); // SourceofJob:[{"18": "Voice Bot"}]

                            enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("b8168e04-7d0a-e911-a94f-000d3af00f43")); // {ServiceAccount:"Dummy Account"}
                            enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("b8168e04-7d0a-e911-a94f-000d3af00f43")); // {BillingAccount:"Dummy Account"}
                            try
                            {
                                Guid jobcreated = service.Create(enWorkorder);
                                Entity servicerequestvoicebot = new Entity("hil_servicerequestvoicebot", EntBot.Id);
                                servicerequestvoicebot["statecode"] = new OptionSetValue(1);
                                servicerequestvoicebot["statuscode"] = new OptionSetValue(2);
                                service.Update(servicerequestvoicebot);
                            }
                            catch (Exception ex)
                            {
                                ErrorDetails = ErrorDetails + ex.Message;
                                Entity servicerequestvoicebot = new Entity("hil_servicerequestvoicebot", EntBot.Id);
                                servicerequestvoicebot["hil_syncresponse"] = ErrorDetails;
                                service.Update(servicerequestvoicebot);
                            }
                        }
                        else {
                            Entity servicerequestvoicebot = new Entity("hil_servicerequestvoicebot", EntBot.Id);
                            servicerequestvoicebot["hil_syncresponse"] = ErrorDetails;
                            service.Update(servicerequestvoicebot);
                        }
                    }
                    catch (Exception ex)
                    {
                        Entity servicerequestvoicebot = new Entity("hil_servicerequestvoicebot", EntBot.Id);
                        servicerequestvoicebot["hil_syncresponse"] = ex.Message + "  " + ErrorDetails;
                        service.Update(servicerequestvoicebot);
                    }
                }
            }
        }
    }
}
