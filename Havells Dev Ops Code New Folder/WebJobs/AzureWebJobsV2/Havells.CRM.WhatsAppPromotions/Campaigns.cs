using Havells.Crm.CommonLibrary;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace Havells.CRM.WhatsAppPromotions
{
    public class Campaigns : Program//, AzureWebJobsLogs
    {
        public Campaigns() { }
        public static void Campaign_NewBuyer(string _campaignName, ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string _prodCategory)
        {
            try
            {
                int totalCount = 0;
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                logs _logs = new logs();
                D365Campaign d365 = new D365Campaign();
                EntityCollection cuatomerAssetColl = new EntityCollection();
                bool moreRecord = true;
                Console.WriteLine("Query execution starts for Campaign Name: " + _campaignName);

                
                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"<fetch version=""1.0"" page=""{pageNumber}"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                        <entity name=""msdyn_customerasset"">
                        <attribute name=""msdyn_customerassetid"" />
                        <attribute name=""msdyn_name"" />
                        <attribute name=""hil_productcategory"" />
                        <attribute name=""hil_productsubcategory"" />
                        <attribute name=""msdyn_product"" />
                        <attribute name=""hil_customer"" />
                        <attribute name='createdon' />
                        <order attribute=""msdyn_name"" descending=""false"" />
                        <filter type=""and"">
                        <condition attribute=""createdon"" operator=""on"" value=""{dateString}"" />
                        <condition attribute=""hil_productcategory"" operator=""eq"" value=""{_prodCategory}"" />
                        </filter>
                        <link-entity name=""contact"" from=""contactid"" to=""hil_customer"" visible=""false"" link-type=""outer"" alias=""consumer"">
                        <attribute name=""mobilephone"" />
                        </link-entity>
                        </entity>
                        </fetch>";
                    EntityCollection jobsColl = service.RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
                    if (jobsColl.Entities.Count > 0)
                    {
                        cuatomerAssetColl.Entities.AddRange(jobsColl.Entities);
                        pageNumber++;
                    }
                    else
                    {
                        moreRecord = false;
                    }
                }
                Console.WriteLine("Query execution ends for Campaign Name: " + _campaignName + " Total records found: " + cuatomerAssetColl.Entities.Count);
                int i = 1;
                string customerName = "Customer";
                string customerMobileNumber = string.Empty;
                string productsubcategory = string.Empty;
                string productmodel = string.Empty;
                string serialNumber = string.Empty;
                string installation_date = string.Empty;
                string _message = string.Empty;
                totalCount = cuatomerAssetColl.Entities.Count;
                foreach (Entity asset in cuatomerAssetColl.Entities)
                {
                    Console.WriteLine("Creating Campaign Data record for: " + _campaignName + ". Count: " + cuatomerAssetColl.Entities.Count + "/" + i++);
                    _message = null;
                    d365 = new D365Campaign();
                    d365.CampaignRunOn = dateString;
                    d365.CampaignName = _campaignName;
                    d365.Customer_Name = asset.Contains("hil_customer") ? asset.GetAttributeValue<EntityReference>("hil_customer").Name : "Customer";
                    if (asset.Contains("hil_customer"))
                        d365.Consumer = asset.GetAttributeValue<EntityReference>("hil_customer");
                    if (asset.Contains("consumer.mobilephone"))
                        d365.Mobile_Number = asset.GetAttributeValue<AliasedValue>("consumer.mobilephone").Value.ToString();
                    if (asset.Contains("hil_productsubcategory"))
                        d365.Model = asset.GetAttributeValue<EntityReference>("hil_productsubcategory");
                    d365.Product_Cat =new EntityReference("product", new Guid(_prodCategory));
                    d365.Serial_Number = asset.ToEntityReference();
                    d365.Registration_Date = asset.GetAttributeValue<DateTime>("createdon");
                   
                    if (_ModeOfComm == ModeOfCommunication.Whatsapp)
                    {
                        d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.Whatsapp);
                    }
                    else
                    {
                        d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.SMS);
                        if (templateName.Trim() == "1107168654225056811")
                        {
                            _message = string.Format("Hi {0},Protect your new Lloyd AC with Havells assured AMC plan. Special price starts at Rs.2499. TnC.  Buy Now https://bit.ly/3py9qIY -Havells", customerName);
                        }
                        else if (templateName.Trim() == "1107169702429550488")
                        {
                            _message = string.Format("Hi {0},Buy Havells AMC plan for your new Havells Water Purifier at flat 20%25 off. TnC. Visit https://bit.ly/3XzawAJ - Havells", customerName);
                        }
                    }
                    d365.Template_ID = templateName;
                    d365.Message = _message;
                    try
                    {
                        createCampaignInD365(d365);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Creating Campaign Data record for: " + _campaignName + ". ERROR: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void Campaign_OutWarrantyBreakdownJobClosure(string _campaignName, ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string _prodCategory)
        {
            try
            {
                int totalCount = 0;
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                logs _logs = new logs();
                D365Campaign d365 = new D365Campaign();
                EntityCollection cuatomerAssetColl = new EntityCollection();
                bool moreRecord = true;
                Console.WriteLine("Query execution starts for Campaign Name: " + _campaignName);
                string _jobStatus = "1727FA6C-FA0F-E911-A94E-000D3AF060A1";
                string _callSubtype = "6560565A-3C0B-E911-A94E-000D3AF06CD4";
                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"<fetch version='1.0' page='{pageNumber}' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_customerref' />
                        <attribute name='msdyn_workorderid' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_mobilenumber' />
                        <attribute name='msdyn_customerasset' />
                        <attribute name='msdyn_timeclosed' />
                        <order attribute='msdyn_timeclosed' descending='false' />
                        <filter type='and'>
                          <condition attribute='hil_isocr' operator='ne' value='1' />
                          <condition attribute='msdyn_timeclosed' operator='on' value='{dateString}' />
                          <condition attribute='msdyn_substatus' operator='eq' value='{_jobStatus}' />
                          <condition attribute='hil_productcategory' operator='eq' value='{_prodCategory}' />
                          <condition attribute='msdyn_customerasset' operator='not-null' />
                          <condition attribute='hil_callsubtype' operator='eq' value='{_callSubtype}' />
                          <condition attribute='hil_laborinwarranty' operator='ne' value='1' />
                          <condition attribute='hil_warrantystatus' operator='ne' value='1' />
                        </filter>
                      </entity>
                     </fetch>";
                    EntityCollection jobsColl = service.RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
                    if (jobsColl.Entities.Count > 0)
                    {
                        cuatomerAssetColl.Entities.AddRange(jobsColl.Entities);
                        pageNumber++;
                    }
                    else
                    {
                        moreRecord = false;
                    }
                }
                Console.WriteLine("Query execution ends for Campaign Name: " + _campaignName + " Total records found: " + cuatomerAssetColl.Entities.Count);
                int i = 1;
                totalCount = cuatomerAssetColl.Entities.Count;
                foreach (Entity asset in cuatomerAssetColl.Entities)
                {
                    Console.WriteLine("Creating Campaign Data record for: " + _campaignName + ". Count: " + cuatomerAssetColl.Entities.Count + "/" + i++);
                    d365 = new D365Campaign();
                    d365.CampaignRunOn = dateString;
                    d365.CampaignName = _campaignName;

                    if (asset.Contains("hil_customerref"))
                    {
                        d365.Customer_Name = asset.Contains("hil_customerref") ? asset.GetAttributeValue<EntityReference>("hil_customerref").Name : "Customer";
                        d365.Consumer = asset.GetAttributeValue<EntityReference>("hil_customerref");
                    }
                    if (asset.Contains("hil_mobilenumber"))
                        d365.Mobile_Number = asset.GetAttributeValue<string>("hil_mobilenumber");
                    if (asset.Contains("hil_productsubcategory"))
                        d365.Model = asset.GetAttributeValue<EntityReference>("hil_productsubcategory");

                    d365.Product_Cat = new EntityReference("product", new Guid(_prodCategory));
                    d365.Serial_Number = asset.GetAttributeValue<EntityReference>("msdyn_customerasset");
                    d365.Registration_Date = asset.GetAttributeValue<DateTime>("createdon");
                    d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.Whatsapp);
                    d365.Job_ID = asset.ToEntityReference();
                    d365.Template_ID = templateName;
                    d365.Message = null;
                    try
                    {
                        createCampaignInD365(d365);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Creating Campaign Data record for: " + _campaignName + ". ERROR: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void Campaign_WarrantyExpireNear(string _campaignName, ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string _prodCategory)
        {
            try
            {
                int totalCount = 0;
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                logs _logs = new logs();
                D365Campaign d365 = new D365Campaign();
                EntityCollection cuatomerAssetColl = new EntityCollection();
                EntityCollection uniqueCustomerAssetColl = new EntityCollection();

                bool moreRecord = true;
                Console.WriteLine("Query execution starts for Campaign Name: " + _campaignName);

                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"<fetch version='1.0' page='{pageNumber}' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='hil_unitwarranty'>
                    <attribute name='hil_customerasset' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_warrantyenddate' operator='on' value='{dateString}' />
                    </filter>
                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='av'>
                        <filter type='and'>
                        <condition attribute='hil_type' operator='in'>
                            <value>3</value>
                            <value>1</value>
                        </condition>
                        </filter>
                    </link-entity>
                    <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='hil_customerasset' link-type='inner' alias='as'>
                        <attribute name='msdyn_product' />
                        <attribute name='hil_customer' />
                        <filter type='and'>
                            <condition attribute='hil_productcategory' operator='eq' value='{_prodCategory}' />
                        </filter>
                        <link-entity name='contact' from='contactid' to='hil_customer' visible='false' link-type='inner' alias='cnt'>
                            <attribute name='mobilephone' />
                        </link-entity>
                    </link-entity>
                    </entity>
                    </fetch>";
                    EntityCollection jobsColl = service.RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
                    if (jobsColl.Entities.Count > 0)
                    {
                        uniqueCustomerAssetColl.Entities.AddRange(jobsColl.Entities);
                        pageNumber++;
                    }
                    else
                    {
                        moreRecord = false;
                    }
                }
                Console.WriteLine("Checking for Duplicate Unit Warranty Lines...");
                int j = 1;
                foreach (Entity entCA in uniqueCustomerAssetColl.Entities) {
                    Console.WriteLine("Checking for Asset : " + j++.ToString() + " : " + entCA.GetAttributeValue<EntityReference>("hil_customerasset").Name);
                    string _fetchXMLDuplicateUWL = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' aggregate='true'>
                      <entity name='hil_unitwarranty'>
                        <attribute name='hil_producttype' groupby='true' alias='ptype'/>
                        <attribute name='hil_customerasset' groupby='true' alias='ca'/>
                        <attribute name='hil_unitwarrantyid' aggregate='count' alias='ca_count'/>
                        <filter type='and'>
                          <condition attribute='hil_customerasset' operator='eq' value='{entCA.GetAttributeValue<EntityReference>("hil_customerasset").Id}' />
                        </filter>
                        <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
                          <filter type='and'>
                            <condition attribute='hil_type' operator='eq' value='1' />
                          </filter>
                        </link-entity>
                      </entity>
                    </fetch>";
                    EntityCollection entColUWL = service.RetrieveMultiple(new FetchExpression(_fetchXMLDuplicateUWL));
                    if (entColUWL.Entities.Count > 1)
                    {
                        EntityReference _entCA = (EntityReference)entColUWL.Entities[0].GetAttributeValue<AliasedValue>("ca").Value;

                        string _fetchDeleteUWL = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_unitwarranty'>
                                <attribute name='hil_unitwarrantyid' />
                                <filter type='and'>
                                  <condition attribute='hil_customerasset' operator='eq' value='{_entCA.Id}' />
                                  <condition attribute='hil_producttype' operator='eq' value='2' />
                                </filter>
                                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='a_c475818aea04e911a94d000d3af06c56'>
                                  <attribute name='hil_warrantyperiod' />
                                  <attribute name='hil_type' />
                                  <filter type='and'>
                                    <condition attribute='hil_type' operator='eq' value='1' />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";
                        EntityCollection entColUWL1 = service.RetrieveMultiple(new FetchExpression(_fetchDeleteUWL));
                        int _rowCount = 1;
                        foreach (Entity ent in entColUWL1.Entities)
                        {
                            service.Delete(ent.LogicalName, ent.Id);
                            Console.WriteLine("Deleting Unit Warranty Lines: " + _rowCount++.ToString());
                        }
                    }
                }
                moreRecord = true;
                pageNumber = 1;
                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"<fetch version='1.0' page='{pageNumber}' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_unitwarranty'>
                    <attribute name='createdon' />
                    <attribute name='hil_customerasset' />
                    <attribute name='hil_unitwarrantyid' />
                    <attribute name='hil_customer' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_warrantyenddate' operator='on' value='{dateString}' />
                    </filter>
                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='av'>
                        <filter type='and'>
                        <condition attribute='hil_type' operator='in'>
                            <value>3</value>
                            <value>1</value>
                        </condition>
                        </filter>
                    </link-entity>
                    <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='hil_customerasset' link-type='inner' alias='as'>
                        <attribute name='msdyn_product' />
                        <attribute name='hil_customer' />
                        <filter type='and'>
                            <condition attribute='hil_productcategory' operator='eq' value='{_prodCategory}' />
                        </filter>
                        <link-entity name='contact' from='contactid' to='hil_customer' visible='false' link-type='inner' alias='cnt'>
                            <attribute name='mobilephone' />
                        </link-entity>
                    </link-entity>
                    </entity>
                    </fetch>";
                    EntityCollection jobsColl = service.RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
                    if (jobsColl.Entities.Count > 0)
                    {
                        cuatomerAssetColl.Entities.AddRange(jobsColl.Entities);
                        pageNumber++;
                    }
                    else
                    {
                        moreRecord = false;
                    }
                }
                Console.WriteLine("Query execution ends for Campaign Name: " + _campaignName + " Total records found: " + cuatomerAssetColl.Entities.Count);
                int i = 1;
                totalCount = cuatomerAssetColl.Entities.Count;
                foreach (Entity asset in cuatomerAssetColl.Entities)
                {
                    Console.WriteLine("Creating Campaign Data record for: " + _campaignName + ". Count: " + cuatomerAssetColl.Entities.Count + "/" + i++);
                    d365 = new D365Campaign();
                    d365.CampaignRunOn = dateString;
                    d365.CampaignName = _campaignName;

                    if (asset.Contains("as.hil_customer"))
                    {
                        d365.Consumer = (EntityReference)asset.GetAttributeValue<AliasedValue>("as.hil_customer").Value;
                        d365.Customer_Name = d365.Consumer.Name;
                    }
                    if (asset.Contains("cnt.mobilephone"))
                        d365.Mobile_Number = asset.GetAttributeValue<AliasedValue>("cnt.mobilephone").Value.ToString();
                    if (asset.Contains("as.msdyn_product"))
                    {
                        d365.Model =  (EntityReference)asset.GetAttributeValue<AliasedValue>("as.msdyn_product").Value;
                        d365.Product_Cat = new EntityReference("product", new Guid(_prodCategory));
                        d365.Serial_Number = asset.GetAttributeValue<EntityReference>("hil_customerasset");
                    }
                    
                    d365.Registration_Date = asset.GetAttributeValue<DateTime>("createdon");
                    d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.Whatsapp);
                    d365.Job_ID = null;
                    d365.Template_ID = templateName;
                    d365.Message = null;
                    d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.Whatsapp);
                    try
                    {
                        createCampaignInD365(d365);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Creating Campaign Data record for: " + _campaignName + ". ERROR: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
      
        //public static void CampaignOnUnitWarranty(string _campaignRunOn, string _campaignName, ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string ProductCategoryGUID)
        //{
        //    int done = 0;
        //    int error = 0;
        //    int skip = 0;
        //    int totalCount = 0;
        //    logs _logs = new logs();
        //    try
        //    {
        //        string dateString = DateTime.Today.AddDays(dayDif).Year.ToString().PadLeft(4, '0') + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
        //        int pageNumber = 1;
        //        EntityCollection cuatomerAssetColl = new EntityCollection();
        //        bool moreRecord = true;
        //        while (moreRecord)
        //        {
        //            string fetchUnitwarranty = $@"<fetch version=""1.0"" page=""{pageNumber}"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
        //                                         <entity name='hil_unitwarranty'>
        //                                            <attribute name='hil_name'/>
        //                                            <attribute name='createdon'/>
        //                                            <attribute name='hil_warrantytemplate'/>
        //                                            <attribute name='hil_warrantystartdate'/>
        //                                            <attribute name='hil_warrantyenddate'/>
        //                                            <attribute name='hil_producttype'/>
        //                                            <attribute name='hil_customerasset'/>
        //                                            <attribute name='hil_unitwarrantyid'/>
        //                                            <order attribute='hil_name' descending='false'/>
        //                                            <filter type='and'>
        //                                                <condition attribute='hil_warrantyenddate' operator='on' value='{dateString}'/>
        //                                            </filter>
        //                                            <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ac'>
        //                                                <filter type='and'>
        //                                                    <condition attribute='hil_type' operator='in'>
        //                                                        <value>3</value>
        //                                                        <value>1</value>
        //                                                    </condition>
        //                                                </filter>
        //                                            </link-entity>
        //                                            <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='hil_customerasset' link-type='inner' alias='ad'>
        //                                                <attribute name='hil_customer'/>
        //                                                <attribute name='msdyn_customerassetid'/>
        //                                                <attribute name='msdyn_name'/>
        //                                                <attribute name='hil_productsubcategorymapping'/>
        //                                                <attribute name='msdyn_product'/>
        //                                                <filter type='and'>
        //                                                    <condition attribute='hil_productcategory' operator='eq' uiname='HAVELLS AQUA' uitype='product' value='{ProductCategoryGUID}'/>
        //                                                </filter>
        //                                            </link-entity>
        //                                          </entity>
        //                                        </fetch>";

        //            EntityCollection warrantyColl = service.RetrieveMultiple(new FetchExpression(fetchUnitwarranty));
        //            if (warrantyColl.Entities.Count > 0)

        //            {
        //                cuatomerAssetColl.Entities.AddRange(warrantyColl.Entities);
        //                pageNumber++;
        //            }
        //            else
        //            {
        //                moreRecord = false;
        //            }
        //        }

        //        D365Campaign d365 = new D365Campaign();
        //        string customerName = "Customer";
        //        string customerMobileNumber = string.Empty;
        //        string productsubcategory = string.Empty;
        //        string productmodel = string.Empty;
        //        string serialNumber = string.Empty;
        //        string installation_date = string.Empty;

        //        Console.WriteLine("Total Jobs Found " + cuatomerAssetColl.Entities.Count);
        //        totalCount = cuatomerAssetColl.Entities.Count;
        //        foreach (Entity job in cuatomerAssetColl.Entities)
        //        {
        //            d365 = new D365Campaign();
        //            customerName = "Customer";
        //            customerMobileNumber = string.Empty;
        //            productmodel = string.Empty;
        //            productsubcategory = string.Empty;
        //            serialNumber = string.Empty;
        //            installation_date = string.Empty;
        //            _logs = new logs();
        //            try
        //            {
        //                string dateStringNew = DateTime.Today.AddDays(dayDif + 2).Year.ToString().PadLeft(4, '0') + "-" + DateTime.Today.AddDays(dayDif + 2).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif + 2).Day.ToString().PadLeft(2, '0');

        //                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                                        <entity name='hil_unitwarranty'>
        //                                            <attribute name='hil_name'/>
        //                                            <attribute name='createdon'/>
        //                                            <attribute name='hil_warrantytemplate'/>
        //                                            <attribute name='hil_warrantystartdate'/>
        //                                            <attribute name='hil_warrantyenddate'/>
        //                                            <attribute name='hil_producttype'/>
        //                                            <attribute name='hil_customerasset'/>
        //                                            <attribute name='hil_unitwarrantyid'/>
        //                                            <order attribute='hil_name' descending='false'/>
        //                                            <filter type='and'>
        //                                                <condition attribute='hil_customerasset' operator='eq' value='{job.GetAttributeValue<EntityReference>("hil_customerasset").Id}'/>
        //                                                <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{dateStringNew}'/>
        //                                            </filter>
        //                                        </entity>
        //                                    </fetch>";

        //                EntityCollection unitwarrColl = service.RetrieveMultiple(new FetchExpression(fetchXml));
        //                if (unitwarrColl.Entities.Count > 0)
        //                {
        //                    skip++;
        //                }
        //                else
        //                {
        //                    customerName = job.Contains("ad.hil_customer") ? ((EntityReference)job.GetAttributeValue<AliasedValue>("ad.hil_customer").Value).Name.ToString() : "";
        //                    d365.CampaignName = _campaignName;
        //                    d365.CampaignRunOn = _campaignRunOn;
        //                    d365.Customer_Name = customerName;
        //                    d365.Consumer = (EntityReference)job.GetAttributeValue<AliasedValue>("ad.hil_customer").Value;
        //                    customerMobileNumber = service.Retrieve("contact", ((EntityReference)job.GetAttributeValue<AliasedValue>("ad.hil_customer").Value).Id,
        //                        new ColumnSet("mobilephone")).GetAttributeValue<string>("mobilephone");
        //                    d365.Mobile_Number = customerMobileNumber;

        //                    productsubcategory = job.Contains("ad.hil_productsubcategorymapping") ? ((EntityReference)job.GetAttributeValue<AliasedValue>("ad.hil_productsubcategorymapping").Value).Name.ToString() : "";
        //                    serialNumber = job.Contains("ad.msdyn_name") ? job.GetAttributeValue<AliasedValue>("ad.msdyn_name").Value.ToString() : "";
        //                    installation_date = job.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");

        //                    d365.Model = (EntityReference)job.GetAttributeValue<AliasedValue>("ad.msdyn_product").Value;
        //                    d365.Product_Cat = new EntityReference("product", new Guid(ProductCategoryGUID));// job.Contains("ad.hil_productsubcategorymapping") ? ((EntityReference)job.GetAttributeValue<AliasedValue>("ad.hil_productsubcategorymapping").Value) : null;
        //                    d365.Serial_Number = job.GetAttributeValue<EntityReference>("hil_customerasset");
                           
        //                    if (_ModeOfComm == ModeOfCommunication.Whatsapp)
        //                    {
        //                        d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.Whatsapp);
        //                        d365.Template_ID = templateName;
        //                        d365.Registration_Date = job.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
        //                    }
        //                    else
        //                    {
        //                        d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.SMS);
        //                    }
        //                    createCampaignInD365(d365);
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                error++;
        //                Console.WriteLine("***** " + error + "/" + cuatomerAssetColl.Entities.Count + "Error Occured " + ex.Message);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("!!!!!!!!!!!!!!! Error " + ex.Message);
        //    }
        //    Console.WriteLine("===================================================================");
        //    Console.WriteLine("Summary");
        //    Console.WriteLine("Template : " + templateName);
        //    Console.WriteLine("Total Recourd : " + totalCount);
        //    Console.WriteLine("Skiped : " + skip);
        //    Console.WriteLine("Send : " + done);
        //    Console.WriteLine("===================================================================");
        //    //SendEmailOnCampaignCompleation(service, templateName, done, _ModeOfComm == ModeOfCommunication.SMS ? "SMS" : "What's App");
        //}
        //public static void CampaignOnWorkOrder(string _campaignRunOn, string _campaignName, ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string ProductCategoryGUID, string Callsubtype, string filter)
        //{
        //    int done = 0;
        //    int error = 0;
        //    int skip = 0;
        //    int totalCount = 0;
        //    logs _logs = new logs();
        //    try
        //    {
        //        D365Campaign d365 = new D365Campaign();
        //        string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
        //        int pageNumber = 1;
        //        EntityCollection cuatomerAssetColl = new EntityCollection();
        //        bool moreRecord = true;
        //        while (moreRecord)
        //        {
        //            string fetchoutwarranty = $@"<fetch version=""1.0"" page=""{pageNumber}"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
        //                <entity name='msdyn_workorder'>
        //                    <attribute name='msdyn_name'/>
        //                    <attribute name='createdon'/>
        //                    <attribute name='hil_productsubcategory'/>
        //                    <attribute name='hil_productcategory'/>
        //                    <attribute name='hil_customerref'/>
        //                    <attribute name='hil_callsubtype'/>
        //                    <attribute name='msdyn_workorderid'/>
        //                    <attribute name='msdyn_timeclosed'/>
        //                    <attribute name='msdyn_customerasset'/>
        //                    <attribute name='hil_laborinwarranty'/>
        //                    <order attribute='msdyn_name' descending='false'/>
        //                    {filter}
        //                    <link-entity name=""msdyn_customerasset"" from=""msdyn_customerassetid"" to=""msdyn_customerasset"" visible=""false"" link-type=""outer"" alias=""a"">
        //                        <attribute name=""msdyn_product""/>
        //                    </link-entity>
        //                </entity>
        //            </fetch>";

        //            EntityCollection outWarrantyColl = service.RetrieveMultiple(new FetchExpression(fetchoutwarranty));
        //            if (outWarrantyColl.Entities.Count > 0)
        //            {
        //                cuatomerAssetColl.Entities.AddRange(outWarrantyColl.Entities);
        //                pageNumber++;
        //            }
        //            else
        //            {
        //                moreRecord = false;
        //            }
        //        }
        //        string customerName = "Customer";
        //        string customerMobileNumber = string.Empty;
        //        string serialNumber = string.Empty;
        //        string installation_date = string.Empty;
        //        Console.WriteLine("Total Jobs Found " + cuatomerAssetColl.Entities.Count);
        //        totalCount = cuatomerAssetColl.Entities.Count;
        //        foreach (Entity job in cuatomerAssetColl.Entities)
        //        {
        //            d365 = new D365Campaign();
        //            customerName = "Customer";
        //            customerMobileNumber = string.Empty;
        //            serialNumber = string.Empty;
        //            installation_date = string.Empty;
        //            try
        //            {
        //                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                    <entity name='hil_unitwarranty'>
        //                        <attribute name='hil_name'/>
        //                        <attribute name='createdon'/>
        //                        <attribute name='hil_warrantytemplate'/>
        //                        <attribute name='hil_warrantystartdate'/>
        //                        <attribute name='hil_warrantyenddate'/>
        //                        <attribute name='hil_producttype'/>
        //                        <attribute name='hil_customerasset'/>
        //                        <attribute name='hil_unitwarrantyid'/>
        //                        <order attribute='hil_name' descending='false'/>
        //                        <filter type='and'>
        //                            <condition attribute='hil_customerasset' operator='eq' value='{job.GetAttributeValue<EntityReference>("msdyn_customerasset").Id}'/>
        //                            <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{dateString}'/>
        //                            <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{dateString}'/>
        //                        </filter>
        //                    </entity>
        //                </fetch>";

        //                EntityCollection unitwarrColl = service.RetrieveMultiple(new FetchExpression(fetchXml));
        //                if (unitwarrColl.Entities.Count > 0)
        //                {
        //                    skip++;
        //                    //Console.WriteLine("Skiped " + skip + "/" + cuatomerAssetColl.Entities.Count + " of searial Number " + job.GetAttributeValue<EntityReference>("msdyn_customerasset").Name);

        //                }
        //                else
        //                {
        //                    _logs = new logs();
        //                    customerName = job.Contains("hil_customerref") ? job.GetAttributeValue<EntityReference>("hil_customerref").Name : "";
        //                    customerMobileNumber = service.Retrieve("contact", job.GetAttributeValue<EntityReference>("hil_customerref").Id,
        //                        new ColumnSet("mobilephone")).GetAttributeValue<string>("mobilephone");
        //                    d365.CampaignRunOn = _campaignRunOn;
        //                    d365.CampaignName = _campaignName;
        //                    d365.Customer_Name = customerName;
        //                    d365.Consumer = job.GetAttributeValue<EntityReference>("hil_customerref");
        //                    ///////Customer Mobile Number HardCode
        //                    // customerMobileNumber = "8005084995";

        //                    serialNumber = job.GetAttributeValue<String>("msdyn_name").ToString();
        //                    installation_date = job.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");
        //                    d365.Registration_Date = job.GetAttributeValue<DateTime>("createdon");
        //                    d365.Job_ID = job.ToEntityReference();
        //                    d365.Mobile_Number = customerMobileNumber;
        //                    d365.Model = (EntityReference)job.GetAttributeValue<AliasedValue>("a.msdyn_product").Value;
        //                    d365.Product_Cat = job.GetAttributeValue<EntityReference>("hil_productcategory");

        //                    //_logs.SerialNumber = serialNumber;
        //                    //_logs.MobileNumber = customerMobileNumber;
        //                    //_logs.JobId = "-";
        //                    //_logs.Template = templateName;
        //                    //_logs.ModeOfComunication = _ModeOfComm == ModeOfCommunication.Whatsapp ? "Whats App" : "SMS";

        //                    if (_ModeOfComm == ModeOfCommunication.Whatsapp)
        //                    {
        //                        d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.Whatsapp);
        //                    }
        //                    else
        //                    {
        //                        d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.SMS);
        //                    }
        //                    d365.Template_ID = templateName;

        //                    createCampaignInD365(d365);
        //                    done++;
        //                    Console.WriteLine("Executed Whatsapp Campaign# " + templateName + " -> " + done.ToString() + "/" + cuatomerAssetColl.Entities.Count + " to " + customerMobileNumber);

        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                error++;
        //                Console.WriteLine("***** " + error + "/" + cuatomerAssetColl.Entities.Count + "Error Occured " + ex.Message);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("!!!!!!!!!!!!!!! Error " + ex.Message);
        //    }
        //    Console.WriteLine("===================================================================");
        //    Console.WriteLine("Summary");
        //    Console.WriteLine("Template : " + templateName);
        //    Console.WriteLine("Total Recourd : " + totalCount);
        //    Console.WriteLine("Skiped : " + skip);
        //    Console.WriteLine("Send : " + done);
        //    Console.WriteLine("===================================================================");
        //   // SendEmailOnCampaignCompleation(service, templateName, done, _ModeOfComm == ModeOfCommunication.SMS ? "SMS" : "What's App");
        //}

        public static void createCampaignInD365(D365Campaign d365)
        {
            string _fetchXMLCondStr = string.Empty;
            if (d365.Job_ID != null)
            {
                _fetchXMLCondStr = $@"<condition attribute='hil_jobid' operator='eq' value='{d365.Job_ID.Id}' />";
            }
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='hil_campaigndata'>
            <attribute name='hil_campaigndataid' />
            <attribute name='hil_name' />
            <attribute name='createdon' />
            <order attribute='hil_name' descending='false' />
            <filter type='and'>
                <condition attribute='hil_name' operator='eq' value='{d365.CampaignName}' />
                <condition attribute='hil_mobilenumber' operator='eq' value='{d365.Mobile_Number}' />
                <condition attribute='hil_campaigndatafordate' operator='on' value='{d365.CampaignRunOn}' />
                <condition attribute='hil_serialnumber' operator='eq' value='{d365.Serial_Number.Id}' />{_fetchXMLCondStr}
            </filter>
            </entity>
            </fetch>";
            
            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count == 0)
            {
                Entity entity = new Entity("hil_campaigndata");
                entity["hil_campaigndatafordate"] = Convert.ToDateTime(d365.CampaignRunOn);
                entity["hil_name"] = d365.CampaignName;
                entity["hil_consumer"] = d365.Consumer;
                entity["hil_mobilenumber"] = d365.Mobile_Number;
                entity["hil_communicationmode"] = d365.Communication_Mode;
                entity["hil_jobid"] = d365.Job_ID;
                entity["hil_customername"] = d365.Customer_Name;
                entity["hil_productcat"] = d365.Product_Cat;
                entity["hil_model"] = d365.Model;
                entity["hil_message"] = d365.Message;
                entity["hil_templateid"] = d365.Template_ID;
                entity["hil_serialnumber"] = d365.Serial_Number;
                entity["hil_registrationdate"] = d365.Registration_Date;
                service.Create(entity);
            }
            else {
                Console.WriteLine("Duplicate Record Found:: Campaign: " + d365.CampaignName + " Mobile Number: " + d365.Mobile_Number + " Campaign Runon: " + d365.CampaignRunOn);
            }
        }
    }
}
