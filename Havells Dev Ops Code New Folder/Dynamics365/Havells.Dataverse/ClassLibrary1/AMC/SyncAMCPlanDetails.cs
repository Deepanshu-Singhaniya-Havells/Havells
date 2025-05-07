using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.AMC
{
    public class SyncAMCPlanDetails : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            string syncDateTime = Convert.ToString(context.InputParameters["syncDateTime"]);
            context.OutputParameters["data"] = JsonSerializer.Serialize(SynAMCPlanDetails(syncDateTime, service));
        }
        public AMCPlanDetailsList SynAMCPlanDetails(string syncDateTime, IOrganizationService service)
        {
            AMCPlanDetailsList objAMCPlanDetailsList = new AMCPlanDetailsList();
            objAMCPlanDetailsList.AMCPlanDetails = new List<AMCPlanDetails>();
            Guid integrationSource = new Guid("608e899b-a8a3-ed11-aad1-6045bdad27a7");
            Guid _amcPriceList = new Guid("c05e7837-8fc4-ed11-b597-6045bdac5e78");
            int _Pageno = 1;
            List<AMCPlanDetails> lstAMCPlanDetails = new List<AMCPlanDetails>();
            try
            {
                if (string.IsNullOrWhiteSpace(syncDateTime))
                {
                    objAMCPlanDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = "No Content : Sync Date time is required." };
                    return objAMCPlanDetailsList;
                }
                DateTime _syncDateTime;
                if (!DateTime.TryParseExact(syncDateTime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out _syncDateTime))
                {
                    objAMCPlanDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = "No Content : Invalid Sync Datetime format. Required format : yyyy-MM-dd HH:mm" };
                    return objAMCPlanDetailsList;
                }
                while (true)
                {
                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' page='{_Pageno}'>
                    <entity name='hil_amcplansetup'>
                    <attribute name='hil_amcplansetupid' />
                    <attribute name='hil_amcplan' />
                    <attribute name='hil_model' />
                    <attribute name='modifiedon' />
                    <order attribute='hil_amcplan' descending='false' />
                    <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />                              
                        <filter type='or'>
                            <condition attribute='hil_applicablesource' operator='eq' value='{integrationSource}' />
                            <condition attribute='hil_applicablesource' operator='null' />
                        </filter>
                    </filter>
                    <link-entity name='product' from='productid' to='hil_model' link-type='inner' alias='aq'>
                    <attribute name='description' />
                    <attribute name='name' />
                    <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='bf'>
                        <attribute name='modifiedon' />
                        <filter type='and'>
                            <condition attribute='hil_eligibleforamc' operator='eq' value='1' />
                        </filter>
                    </link-entity>
                    </link-entity>
                    <link-entity name='product' from='productid' to='hil_amcplan' link-type='inner' alias='bg'>
                        <attribute name='modifiedon' />
                        <filter type='and'>
                            <condition attribute='hil_hierarchylevel' operator='eq' value='910590001' />
                        </filter>
                    <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='pc'>
                        <attribute name='hil_name' />   
                        <attribute name='createdon' />   
                        <attribute name='hil_plantclink' />   
                        <attribute name='hil_planperiod' />   
                        <attribute name='hil_notcovered' />
                        <attribute name='hil_coverage' />
                        <attribute name='hil_amctandc' />
                        <attribute name='modifiedon' />
                        <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />                              
                        </filter>
                    </link-entity>
                    <link-entity name='productpricelevel' from='productid' to='productid' link-type='inner' alias='pricelist'>
                    <attribute name='amount' />
                    <attribute name='modifiedon' />
                    <filter type='and'>
                        <condition attribute='pricelevelid' operator='eq' value='{_amcPriceList}' />
                    </filter>
                    </link-entity>
                    </link-entity>
                    </entity>
                    </fetch>";

                    EntityCollection entCollProduct = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entCollProduct.Entities.Count > 0)
                    {
                        var PlanList = entCollProduct.Entities.Where(m => m.GetAttributeValue<DateTime>("modifiedon").AddMinutes(330) >= _syncDateTime
                                   || ((DateTime)m.GetAttributeValue<AliasedValue>("bf.modifiedon").Value).AddMinutes(330) >= _syncDateTime
                                   || ((DateTime)m.GetAttributeValue<AliasedValue>("bg.modifiedon").Value).AddMinutes(330) >= _syncDateTime
                                   || ((DateTime)m.GetAttributeValue<AliasedValue>("pc.modifiedon").Value).AddMinutes(330) >= _syncDateTime
                                   || ((DateTime)m.GetAttributeValue<AliasedValue>("pricelist.modifiedon").Value).AddMinutes(330) >= _syncDateTime).Select(m => m).ToList();

                        decimal DiscPer = GetDiscountValueBySource(integrationSource.ToString(), service);
                        foreach (Entity entProduct in PlanList)
                        {
                            AMCPlanDetails objAMCPlanInfo = new AMCPlanDetails();
                            objAMCPlanInfo.ModelId = entProduct.GetAttributeValue<EntityReference>("hil_model").Id;
                            objAMCPlanInfo.ModelName = entProduct.Contains("aq.description") ? entProduct.GetAttributeValue<AliasedValue>("aq.description").Value.ToString() : "";
                            objAMCPlanInfo.ModelNumber = entProduct.Contains("aq.name") ? entProduct.GetAttributeValue<AliasedValue>("aq.name").Value.ToString() : "";
                            if (entProduct.Contains("hil_amcplan"))
                            {
                                Guid sPart = entProduct.GetAttributeValue<EntityReference>("hil_amcplan").Id;
                                objAMCPlanInfo.PlanId = sPart;
                                objAMCPlanInfo.DiscountPercent = DiscPer;
                                objAMCPlanInfo.MRP = decimal.Round((entProduct.Contains("pricelist.amount") ? ((Money)(entProduct.GetAttributeValue<AliasedValue>("pricelist.amount").Value)).Value : 0), 2);
                                objAMCPlanInfo.Coverage = entProduct.Contains("pc.hil_coverage") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_coverage").Value.ToString() : null;
                                objAMCPlanInfo.NonCoverage = entProduct.Contains("pc.hil_notcovered") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_notcovered").Value.ToString() : null;
                                objAMCPlanInfo.PlanName = entProduct.Contains("pc.hil_name") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_name").Value.ToString() : null;
                                objAMCPlanInfo.PlanPeriod = entProduct.Contains("pc.hil_planperiod") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_planperiod").Value.ToString() : null;
                                objAMCPlanInfo.PlanTCLink = entProduct.Contains("pc.hil_plantclink") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_plantclink").Value.ToString() : null;
                            }
                            lstAMCPlanDetails.Add(objAMCPlanInfo);
                        }
                        objAMCPlanDetailsList.AMCPlanDetails = lstAMCPlanDetails;
                        objAMCPlanDetailsList.Result = new ResResult { ResultStatus = true, ResultMessage = "Success" };
                        _Pageno++;
                    }
                    else
                    {
                        break;
                    }
                }
                return objAMCPlanDetailsList;
            }
            catch (Exception ex)
            {
                objAMCPlanDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = ex.Message };
                return objAMCPlanDetailsList;
            }
        }
        public decimal GetDiscountValueBySource(string source, IOrganizationService service)
        {
            decimal DiscPer = 0;
            string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                              <entity name='hil_amcdiscountmatrix'>
                                 <attribute name='hil_discper' />|
                                  <order attribute='modifiedon' descending='true' />
                                  <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0'/>
                                    <condition attribute='hil_appliedto' operator='eq' value='{source}'/>
                                  </filter>
                              </entity>
                            </fetch>";
            EntityCollection entamcdiscountmatrix = service.RetrieveMultiple(new FetchExpression(fetchQuery));
            if (entamcdiscountmatrix.Entities.Count > 0)
            {
                if (entamcdiscountmatrix.Entities[0].Contains("hil_discper"))
                {
                    DiscPer = entamcdiscountmatrix.Entities[0].GetAttributeValue<Decimal>("hil_discper");
                }
            }
            return DiscPer;
        }
    }
    public class AMCPlanDetailsList
    {
        public List<AMCPlanDetails> AMCPlanDetails { get; set; }
        public ResResult Result { get; set; }
    }
    public class AMCPlanDetails
    {
        public Guid ModelId { get; set; }
        public string ModelNumber { get; set; }
        public string ModelName { get; set; }
        public Guid PlanId { get; set; }
        public string PlanName { get; set; }
        public string PlanPeriod { get; set; }
        public decimal MRP { get; set; }
        public decimal DiscountPercent { get; set; }
        public string Coverage { get; set; }
        public string NonCoverage { get; set; }
        public string PlanTCLink { get; set; }
    }
    public class ResResult
    {
        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }
    }
    public class AMCPlanInfo
    {
        public Guid PlanId { get; set; }
        public string PlanName { get; set; }
        public string PlanPeriod { get; set; }
        public string MRP { get; set; }
        public string DiscountPercent { get; set; }
        public string EffectivePrice
        {
            get
            {
                if (DiscountPercent != null)
                {
                    return decimal.Round(Convert.ToDecimal(MRP) - ((Convert.ToDecimal(MRP) * Convert.ToDecimal(DiscountPercent)) / 100), 2).ToString();
                }
                else
                {
                    return decimal.Round(Convert.ToDecimal(MRP), 2).ToString();
                }
            }
        }
        public string Coverage { get; set; }
        public string NonCoverage { get; set; }
        public string PlanTCLink { get; set; }
        public string PurchaseDate { get; set; }
    }
}

