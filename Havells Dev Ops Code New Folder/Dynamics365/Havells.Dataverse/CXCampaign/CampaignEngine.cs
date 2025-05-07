using System;
using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using static CXCampaign.CXCampaignModels;

namespace CXCampaign
{
    internal class CampaignEngine : DataverseServiceFactory
    {
        internal void Campaign_MoengageKKGClosure(string _campaignName, ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string _prodCategory)
        {
            try
            {
                int totalCount = 0;
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                D365Campaign d365 = new D365Campaign();
                EntityCollection cuatomerAssetColl = new EntityCollection();
                bool moreRecord = true;
                Console.WriteLine("Query execution starts for Campaign Name: " + _campaignName);

                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"";
                    EntityCollection jobsColl = RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
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
                    d365.Product_Cat = new EntityReference("product", new Guid(_prodCategory));
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
                        CreateCampaignInD365(d365);
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
        internal void Campaign_NewBuyer(string _campaignName, ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string _prodCategory)
        {
            try
            {
                int totalCount = 0;
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                D365Campaign d365 = new D365Campaign();
                EntityCollection cuatomerAssetColl = new EntityCollection();
                bool moreRecord = true;
                Console.WriteLine("Query execution starts for Campaign Name: " + _campaignName);

                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true' page='{pageNumber}'>
                        <entity name='contact'>
                        <attribute name='fullname' />
                        <attribute name='contactid' />
                        <attribute name='mobilephone' />
                        <order attribute='fullname' descending='false' />
                        <link-entity name='msdyn_customerasset' from='hil_customer' to='contactid' link-type='inner' alias='ad'>
                            <attribute name='hil_productcategory' />
                            <filter type='and'>
                            <condition attribute='createdon' operator='on' value='{dateString}' />
                            <condition attribute='hil_productcategory' operator='eq' value='{_prodCategory}' />
                            </filter>
                        </link-entity>
                        </entity>
                        </fetch>";
                    EntityCollection jobsColl = RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
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
                    d365.Customer_Name = asset.Contains("fullname") ? asset.GetAttributeValue<string>("fullname") : "Customer";
                    d365.Consumer = asset.ToEntityReference();
                    if (asset.Contains("mobilephone"))
                        d365.Mobile_Number = asset.GetAttributeValue<string>("mobilephone");
                    d365.Product_Cat = (EntityReference)asset.GetAttributeValue<AliasedValue>("ad.hil_productcategory").Value;
                    d365.Registration_Date = Convert.ToDateTime(dateString).Date;

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
                        CreateCampaignInD365(d365);
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
        internal void Campaign_OutWarrantyBreakdownJobClosure(string _campaignName, ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string _prodCategory)
        {
            try
            {
                int totalCount = 0;
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                D365Campaign d365 = new D365Campaign();
                EntityCollection cuatomerAssetColl = new EntityCollection();
                bool moreRecord = true;
                Console.WriteLine("Query execution starts for Campaign Name: " + _campaignName);
                string _jobStatus = "1727FA6C-FA0F-E911-A94E-000D3AF060A1";//Closed
                string _callSubtype = "6560565A-3C0B-E911-A94E-000D3AF06CD4"; //Breakdown
                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true' page='{pageNumber}'>
                        <entity name='contact'>
                        <attribute name='fullname' />
                        <attribute name='contactid' />
                        <attribute name='mobilephone' />
                        <order attribute='fullname' descending='false' />
                        <link-entity name='msdyn_workorder' from='hil_customerref' to='contactid' link-type='inner' alias='ad'>
                            <attribute name='hil_productcategory' />
                            <filter type='and'>
                            <condition attribute='hil_isocr' operator='ne' value='1' />
                            <condition attribute='msdyn_timeclosed' operator='on' value='{dateString}' />
                            <condition attribute='msdyn_substatus' operator='eq' value='{_jobStatus}' />
                            <condition attribute='hil_productcategory' operator='eq' value='{_prodCategory}' />
                            <condition attribute='msdyn_customerasset' operator='not-null' />
                            <condition attribute='hil_callsubtype' operator='eq' value='{_callSubtype}' />
                            <condition attribute='hil_laborinwarranty' operator='ne' value='1' />
                            </filter>
                        </link-entity>
                        </entity>
                        </fetch>";
                    EntityCollection jobsColl = RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
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

                    if (asset.Contains("fullname"))
                    {
                        d365.Customer_Name = asset.Contains("fullname") ? asset.GetAttributeValue<string>("fullname") : "Customer";
                        d365.Consumer = asset.ToEntityReference();
                    }
                    if (asset.Contains("mobilephone"))
                        d365.Mobile_Number = asset.GetAttributeValue<string>("mobilephone");

                    d365.Model = null;
                    d365.Product_Cat = (EntityReference)asset.GetAttributeValue<AliasedValue>("ad.hil_productcategory").Value;
                    d365.Serial_Number = null;
                    d365.Registration_Date = Convert.ToDateTime(dateString).Date;
                    d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.Whatsapp);
                    d365.Job_ID = null;
                    d365.Template_ID = templateName;
                    d365.Message = null;
                    try
                    {
                        CreateCampaignInD365(d365);
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
        internal void Campaign_WarrantyExpireNear(string _campaignName, ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string _prodCategory)
        {
            try
            {
                int totalCount = 0;
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                D365Campaign d365 = new D365Campaign();
                EntityCollection cuatomerAssetColl = new EntityCollection();
                EntityCollection uniqueCustomerAssetColl = new EntityCollection();

                bool moreRecord = true;
                Console.WriteLine("Query execution starts for Campaign Name: " + _campaignName);

                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true' page='{pageNumber}'>
                      <entity name='contact'>
                        <attribute name='fullname' />
                        <attribute name='contactid' />
                        <attribute name='mobilephone' />
                        <order attribute='fullname' descending='false' />
                        <link-entity name='msdyn_customerasset' from='hil_customer' to='contactid' link-type='inner' alias='aj'>
                          <attribute name='hil_productcategory' />
                          <filter type='and'>
                            <condition attribute='hil_productcategory' operator='eq' value='{_prodCategory}' />
                          </filter>
                          <link-entity name='hil_unitwarranty' from='hil_customerasset' to='msdyn_customerassetid' link-type='inner' alias='ak'>
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                              <condition attribute='hil_warrantyenddate' operator='on' value='{dateString}' />
                            </filter>
                            <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='al'>
                              <filter type='and'>
                                <condition attribute='hil_type' operator='in'>
                                  <value>3</value>
                                  <value>1</value>
                                </condition>
                              </filter>
                            </link-entity>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
                    EntityCollection jobsColl = RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
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
                Console.WriteLine("Query execution ends for Campaign Name: " + _campaignName + " Total records found: " + uniqueCustomerAssetColl.Entities.Count);
                int i = 1;
                totalCount = uniqueCustomerAssetColl.Entities.Count;
                foreach (Entity asset in uniqueCustomerAssetColl.Entities)
                {
                    Console.WriteLine("Creating Campaign Data record for: " + _campaignName + ". Count: " + uniqueCustomerAssetColl.Entities.Count + "/" + i++);
                    d365 = new D365Campaign();
                    d365.CampaignRunOn = dateString;
                    d365.CampaignName = _campaignName;

                    if (asset.Contains("fullname"))
                    {
                        d365.Consumer = asset.ToEntityReference();
                        d365.Customer_Name = asset.GetAttributeValue<string>("fullname");
                    }
                    if (asset.Contains("mobilephone"))
                        d365.Mobile_Number = asset.GetAttributeValue<string>("mobilephone");

                    d365.Model = null;
                    d365.Serial_Number = null;
                    if (asset.Contains("aj.hil_productcategory"))
                    {
                        d365.Product_Cat = (EntityReference)asset.GetAttributeValue<AliasedValue>("aj.hil_productcategory").Value;
                    }

                    d365.Registration_Date = Convert.ToDateTime(dateString).Date;
                    d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.Whatsapp);
                    d365.Job_ID = null;
                    d365.Template_ID = templateName;
                    d365.Message = null;
                    d365.Communication_Mode = new OptionSetValue((int)ModeOfCommunication.Whatsapp);
                    try
                    {
                        CreateCampaignInD365(d365);
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
        internal void CreateCampaignInD365(D365Campaign d365)
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
                    {_fetchXMLCondStr}
                </filter>
                </entity>
                </fetch>";

            EntityCollection entCol = RetrieveMultiple(new FetchExpression(_fetchXML));
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
                Create(entity);
            }
            else
            {
                Console.WriteLine("Duplicate Record Found:: Campaign: " + d365.CampaignName + " Mobile Number: " + d365.Mobile_Number + " Campaign Runon: " + d365.CampaignRunOn);
            }
        }
    }
}
