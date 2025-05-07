using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace D365WebJobs
{
    public class TechnicianMobileExt
    {
        public List<Visit> vis = new List<Visit>();
        public List<Consumable> con = new List<Consumable>();
        public ConsumablesData RequestData(Request req, IOrganizationService service)
        {
            ConsumablesData response = new ConsumablesData() { ResultStatus = true, ResultMessage = "Success" };
            try
            {
                if (service != null)
                {
                    if (!string.IsNullOrEmpty(req.AssetID) && !string.IsNullOrWhiteSpace(req.AssetID))
                    {
                        string FetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='msdyn_customerasset'>
                            <attribute name='msdyn_customerassetid' />
                            <attribute name='hil_invoicedate' />
                            <attribute name='hil_warrantytilldate' />
                            <attribute name='hil_warrantystatus' />
                            <attribute name='hil_warrantysubstatus' />
                            <filter type='and'>
                              <condition attribute='msdyn_name' operator='eq' value='{req.AssetID}' />
                            </filter>
                          </entity>
                        </fetch>";
                        EntityCollection assetdet = service.RetrieveMultiple(new FetchExpression(FetchQuery));
                        if (assetdet.Entities.Count > 0)
                        {
                            if (assetdet.Entities[0].Attributes.Contains("hil_invoicedate"))
                            {
                                DateTime _dop = assetdet.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330);
                                int age = Convert.ToInt32(((_dop - DateTime.Now).TotalDays / 365));
                                response.AssetAge = age < 1 ? "<1 yr" : age + " yr";
                                response.DOP = assetdet.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330).ToString("yyyy-MMMM-dd");
                            }
                            if (assetdet.Entities[0].Attributes.Contains("hil_warrantystatus")) {
                                response.WarrantyStatus = assetdet.Entities[0].FormattedValues["hil_warrantystatus"].ToString();
                            }
                            if (assetdet.Entities[0].Attributes.Contains("hil_warrantysubstatus"))
                            {
                                response.WarrantyStatus = "(" + assetdet.Entities[0].FormattedValues["hil_warrantysubstatus"].ToString() + ")";
                            }
                            if (assetdet.Entities[0].Attributes.Contains("hil_warrantytilldate"))
                            {
                                response.WarrantyEndDate = assetdet.Entities[0].GetAttributeValue<DateTime>("hil_warrantytilldate").AddMinutes(330).ToString("yyyy-MMMM-dd");
                            }

                            string _currentDate = DateTime.Now.ToString("yyyy-MM-dd");
                            FetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_unitwarranty'>
                                <attribute name='hil_warrantyenddate' />
                                <order attribute='hil_warrantyenddate' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_customerasset' operator='eq' value='{assetdet.Entities[0].Id}' />
                                  <condition attribute='hil_warrantystartdate' operator='on-or-before' value='{_currentDate}' />
                                  <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{_currentDate}' />
                                  <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' visible='false' link-type='outer' alias='wt'>
                                  <attribute name='hil_description' />
                                  <attribute name='hil_amcplan' />
                                </link-entity>
                              </entity>
                            </fetch>";
                            EntityCollection assetWtydet = service.RetrieveMultiple(new FetchExpression(FetchQuery));
                            if (assetWtydet.Entities.Count > 0)
                            {
                                if (assetWtydet.Entities[0].Attributes.Contains("wt.hil_amcplan"))
                                {
                                    if (assetWtydet.Entities[0].Contains("wt.hil_amcplan"))
                                        response.AMCPlan = ((EntityReference)assetWtydet.Entities[0].GetAttributeValue<AliasedValue>("wt.hil_amcplan").Value).Name;
                                    if (assetWtydet.Entities[0].Contains("wt.hil_description"))
                                        response.AMCPlanCoverage = assetWtydet.Entities[0].GetAttributeValue<AliasedValue>("wt.hil_description").Value.ToString();
                                }
                            }
                            else {
                                response.WarrantyStatus = "Out Waranty";
                            }
                        }

                        FetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorderservice'>
                        <attribute name='createdon' />
                        <attribute name='msdyn_workorder' />
                        <attribute name='msdyn_name' />
                        <attribute name='msdyn_workorderincident' />
                        <attribute name='msdyn_service' />
                        <order attribute='createdon' descending='true' />
                        <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' link-type='inner' alias='ac'>
                        <filter type='and'>
                            <condition attribute='msdyn_name' operator='eq' value='{req.AssetID}' />
                        </filter>
                        </link-entity>
                        <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' visible='false' link-type='outer' alias='wo'>
                            <attribute name='ownerid' />
                            <filter type='and'>
                                <condition attribute='msdyn_name' operator='ne' value='{req.JobID}' />
                            </filter>
                        </link-entity>
                        </entity>
                        </fetch>";

                        EntityCollection visitCollection = service.RetrieveMultiple(new FetchExpression(FetchQuery));
                        int cnt = 1;
                        string _lastVisitSummary = string.Empty, _lastVisitBy = string.Empty;
                        Guid _lastIncidentId = Guid.Empty;

                        foreach (Entity entity in visitCollection.Entities)
                        {
                            Visit temp = new Visit();
                            temp.JobID = entity.GetAttributeValue<EntityReference>("msdyn_workorder").Name.ToString();
                            if (entity.Contains("createdon"))
                            {
                                temp.VisitDate = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString("yyyy-MMMM-dd");
                            }
                            
                            temp.Service = entity.GetAttributeValue<EntityReference>("msdyn_service").Name;
                            vis.Add(temp);
                            if (cnt == 1)
                            {
                                if (entity.Contains("wo.ownerid"))
                                {
                                    response.LastVisitBy = ((EntityReference)entity.GetAttributeValue<AliasedValue>("wo.ownerid").Value).Name.ToString();
                                }
                                response.LastVisitDate = temp.VisitDate.ToString();
                                _lastVisitSummary = temp.Service;
                                _lastIncidentId = entity.GetAttributeValue<EntityReference>("msdyn_workorderincident").Id;
                            }
                            cnt++;
                        }
                        response.Visits = vis;
                        response.NumOfVisits = vis.Count;

                        FetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorderproduct'>
                        <attribute name='createdon' />
                        <attribute name='hil_replacedpart' />
                        <attribute name='msdyn_quantity' />
                        <attribute name='hil_partamount' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='msdyn_workorder' />
                        <attribute name='msdyn_workorderincident' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                            <condition attribute='hil_markused' operator='eq' value='1' />
                        </filter>
                        <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' link-type='inner' alias='ar'>
                        <filter type='and'>
                            <condition attribute='msdyn_name' operator='eq' value='{req.AssetID}' />
                        </filter>
                        </link-entity>
                        <link-entity name='product' from='productid' to='hil_replacedpart' visible='false' link-type='outer' alias='prd'>
                            <attribute name='description' />
                            <attribute name='hil_sparepartfamily' />
                        </link-entity>
                        </entity>
                        </fetch>";

                        EntityCollection consumablesCollection = service.RetrieveMultiple(new FetchExpression(FetchQuery));
                        OptionSetValue _wrt;
                        int count = consumablesCollection.Entities.Count;
                        foreach (Entity entity in consumablesCollection.Entities)
                        {
                            Consumable temp = new Consumable();
                            if (entity.Contains("msdyn_workorder"))
                                temp.JobID = entity.GetAttributeValue<EntityReference>("msdyn_workorder").Name.ToString();
                            if (entity.Contains("hil_partamount"))
                                temp.Amount = entity.GetAttributeValue<Decimal>("hil_partamount").ToString("0:00");
                            if (entity.Contains("msdyn_quantity"))
                                temp.Quantity = entity.GetAttributeValue<Double>("msdyn_quantity").ToString();
                            if (entity.Contains("hil_replacedpart"))
                                temp.SparePartCode = entity.GetAttributeValue<EntityReference>("hil_replacedpart").Name.ToString();
                            if (entity.Contains("prd.description"))
                                temp.SparePartName = entity.GetAttributeValue<AliasedValue>("prd.description").Value.ToString();
                            if (entity.Contains("hil_warrantystatus"))
                            {
                                _wrt = entity.GetAttributeValue<OptionSetValue>("hil_warrantystatus");
                                temp.WarrantyStatus = _wrt.Value == 1 ? "IN" : "OUT";
                            }

                            con.Add(temp);
                            if (_lastIncidentId == entity.GetAttributeValue<EntityReference>("msdyn_workorderincident").Id)
                            {
                                _lastVisitSummary += ", " + temp.SparePartName;
                            }
                        }
                        response.LastVisitSummary = _lastVisitSummary;
                        response.Consumables = con;
                        response.SparePartFamily = getSparePartFamiy(req.AssetID, service);
                    }
                    else if (!string.IsNullOrEmpty(req.MobileNumber) && !string.IsNullOrEmpty(req.AssetSubcategoryID))
                    {

                        List<CustomerAsset> lstCustomerAsset = new List<CustomerAsset>();
                        string fetchXML = string.Empty;
                        EntityCollection entColl = null;
                        if (!string.IsNullOrEmpty(req.JobID))
                        {
                            fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='msdyn_workorder'>
                                <attribute name='msdyn_customerasset' />
                                <attribute name='hil_productsubcategory' />
                                <filter type='and'>
                                  <condition attribute='msdyn_name' operator='eq' value='{req.JobID}' />
                                  <condition attribute='msdyn_customerasset' operator='not-null' />
                                </filter>
                              </entity>
                            </fetch>";
                            entColl = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entColl.Entities.Count > 0)
                            {
                                foreach (Entity ent in entColl.Entities)
                                {
                                    lstCustomerAsset.Add(new CustomerAsset()
                                    {
                                        SerialNumber = ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Name,
                                        SubCategory = ent.Contains("hil_productsubcategory") ? ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : null
                                    });
                                }
                                return new ConsumablesData { ResultStatus = true, ResultMessage = "Success", CustomerAssets = lstCustomerAsset };
                            }
                        }

                        fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='msdyn_customerasset'>
                                <attribute name='msdyn_name' />
                                <attribute name='hil_productsubcategory' />
                            <filter type='and'>
                                <condition attribute='hil_productsubcategory' operator='eq' value='{new Guid(req.AssetSubcategoryID)}' />
                            </filter>
                            <link-entity name='contact' from='contactid' to='hil_customer' link-type='inner' alias='ad'>
                                <filter type='and'>
                                <condition attribute='mobilephone' operator='eq' value='{req.MobileNumber}' />
                                </filter>
                            </link-entity>
                            </entity>
                            </fetch>";
                        EntityCollection entCollProduct = service.RetrieveMultiple(new FetchExpression(fetchXML));

                        if (entCollProduct.Entities.Count > 0)
                        {
                            foreach (Entity ent in entCollProduct.Entities)
                            {
                                lstCustomerAsset.Add(new CustomerAsset()
                                {
                                    SerialNumber = ent.Contains("msdyn_name") ? ent.GetAttributeValue<string>("msdyn_name") : null,
                                    SubCategory = ent.Contains("hil_productsubcategory") ? ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : null
                                });
                            }
                        }
                        response = new ConsumablesData { ResultStatus = true, ResultMessage = "Success", CustomerAssets = lstCustomerAsset };

                    }
                    else
                    {
                        response = new ConsumablesData { ResultStatus = false, ResultMessage = "Asset Id/Mobile Number with SubCategory Id is required." };
                    }
                }
                else
                {
                    response = new ConsumablesData { ResultStatus = false, ResultMessage = "D365 Service is not available. : " };
                }
            }
            catch (Exception ex)
            {
                response = new ConsumablesData { ResultStatus = false, ResultMessage = "D365 Internal Server Error : " + ex.Message };
            }
            return response;
        }
        public List<SparePartFamily> getSparePartFamiy(string AssetID, IOrganizationService _service)
        {
            string FetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorderproduct'>
                        <attribute name='createdon' />
                        <attribute name='hil_replacedpart' />
                        <attribute name='msdyn_quantity' />
                        <attribute name='hil_partamount' />
                        <attribute name='hil_warrantystatus' />
                        <attribute name='msdyn_workorder' />
                        <attribute name='msdyn_workorderincident' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                            <condition attribute='hil_markused' operator='eq' value='1' />
                        </filter>
                        <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' link-type='inner' alias='ar'>
                            <filter type='and'>
                                <condition attribute='msdyn_name' operator='eq' value='{AssetID}' />
                            </filter>
                        </link-entity>
                        <link-entity name='product' from='productid' to='hil_replacedpart' visible='false' link-type='outer' alias='prd'>
                            <attribute name='description' />
                            <attribute name='hil_sparepartfamily' />
                        </link-entity>
                        </entity>
                        </fetch>
            ";

            EntityCollection consumablesCollection = _service.RetrieveMultiple(new FetchExpression(FetchQuery));
            List<SparePartFamily> list = new List<SparePartFamily>();
            foreach (Entity entity in consumablesCollection.Entities)
            {
                SparePartFamily sparePart = new SparePartFamily();
                sparePart.SpareFamilyName = entity.Contains("prd.hil_sparepartfamily") ? ((EntityReference)entity.GetAttributeValue<AliasedValue>("prd.hil_sparepartfamily").Value).Name : "NA";
                if (sparePart.SpareFamilyName.Contains("-"))
                {
                    sparePart.SpareFamilyName = sparePart.SpareFamilyName.Substring(0, sparePart.SpareFamilyName.IndexOf('-'));
                }
                sparePart.Quantity = (int)entity.GetAttributeValue<Double>("msdyn_quantity");
                list.Add(sparePart);
            }
            List<string> ABC = list.Select(s => s.SpareFamilyName).Distinct().ToList();
            List<SparePartFamily> SparePartFamilyCount = new List<SparePartFamily>();
            foreach (string spare in ABC)
            {
                int s = list.Where(sa => sa.SpareFamilyName.Equals(spare)).Sum(sa => sa.Quantity);
                SparePartFamilyCount.Add(new SparePartFamily { Quantity = s, SpareFamilyName = spare });
            }
            return SparePartFamilyCount;
        }
    }
    
    public class Request
    {
        
        public string AssetID { get; set; }
        
        public string JobID { get; set; }
        
        public string MobileNumber { get; set; }
        
        public string AssetSubcategoryID { get; set; }
    }
    
    public class ConsumablesData
    {
        
        public bool ResultStatus { get; set; }
        
        public string ResultMessage { get; set; }
        
        public int NumOfVisits { get; set; }
        
        public string WarrantyStatus { get; set; }
        
        public string AMCPlan { get; set; }
        
        public string AMCPlanCoverage { get; set; }
        
        public string WarrantyEndDate { get; set; }
        
        public string DOP { get; set; }
        
        public string AssetAge { get; set; }
        
        public string LastVisitDate { get; set; }
        
        public string LastVisitBy { get; set; }
        
        public string LastVisitSummary { get; set; }
        
        public List<Visit> Visits { get; set; }
        
        public List<Consumable> Consumables { get; set; }
        
        public List<SparePartFamily> SparePartFamily { get; set; }
        
        public List<CustomerAsset> CustomerAssets { get; set; }
    }
    
    public class Visit
    {
        
        public string JobID { get; set; }
        
        public string VisitDate { get; set; }
        
        public string Service { get; set; }
    }
    
    public class Consumable
    {
        
        public string JobID { get; set; }
        
        public string SparePartCode { get; set; }
        
        public string SparePartName { get; set; }
        
        public string SparePartFamilyName { get; set; }
        
        public string Quantity { get; set; }
        
        public string Amount { get; set; }
        
        public string WarrantyStatus { get; set; }
    }

    
    public class SparePartFamily
    {
        
        public int Quantity { get; set; }
        
        public string SpareFamilyName { get; set; }
    }
    
    public class CustomerAsset
    {
        
        public string SerialNumber { get; set; }
        
        public string SubCategory { get; set; }
    }
}
