using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class NatureOfComplaint
    {
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public string ProductCategoryName { get; set; }
        [DataMember]
        public Guid Guid { get; set; }
        [DataMember]
        public Guid ProductCategoryGuid { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }

        public List<NatureOfComplaint> GetNatureOfComplaints(NatureOfComplaint natureOfComplaint)
        {
            NatureOfComplaint objNatureOfComplaint;
            List<NatureOfComplaint> lstNatureOfComplaint = new List<NatureOfComplaint>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (natureOfComplaint.SerialNumber.Trim().Length == 0)
                {
                    objNatureOfComplaint = new NatureOfComplaint { ResultStatus = false, ResultMessage = "Product Serial Number is required." };
                    lstNatureOfComplaint.Add(objNatureOfComplaint);
                    return lstNatureOfComplaint;
                }

                Query = new QueryExpression("msdyn_customerasset");
                Query.ColumnSet = new ColumnSet("msdyn_name");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, natureOfComplaint.SerialNumber);
                Query.TopCount = 1;
                entcoll = service.RetrieveMultiple(Query);

                if (entcoll.Entities.Count == 0)
                {
                    objNatureOfComplaint = new NatureOfComplaint { ResultStatus = false, ResultMessage = "Product Serial Number does not exist." };
                    lstNatureOfComplaint.Add(objNatureOfComplaint);
                    return lstNatureOfComplaint;
                }
                else
                {
                    string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='hil_natureofcomplaint'>
                        <attribute name='hil_name' />
                        <attribute name='hil_natureofcomplaintid' />
                        <order attribute='hil_name' descending='false' />
                        <link-entity name='product' from='productid' to='hil_relatedproduct' link-type='inner' alias='ae'>
                            <link-entity name='msdyn_customerasset' from='hil_productsubcategory' to='productid' link-type='inner' alias='af'>
                                <filter type='and'>
                                    <condition attribute='msdyn_name' operator='eq' value='" + natureOfComplaint.SerialNumber + @"' />
                                </filter>
                            </link-entity>
                        </link-entity>
                    </entity>
                    </fetch>";
                    entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entcoll.Entities.Count == 0)
                    {
                        objNatureOfComplaint = new NatureOfComplaint { ResultStatus = false, ResultMessage = "No Nature of Complaint is mapped with Serial Number."};
                        lstNatureOfComplaint.Add(objNatureOfComplaint);
                    }
                    else {
                        foreach (Entity ent in entcoll.Entities) {
                            objNatureOfComplaint = new NatureOfComplaint();
                            objNatureOfComplaint.Guid = ent.GetAttributeValue<Guid>("hil_natureofcomplaintid");
                            objNatureOfComplaint.Name = ent.GetAttributeValue<string>("hil_name");
                            objNatureOfComplaint.SerialNumber = natureOfComplaint.SerialNumber;
                            objNatureOfComplaint.ResultStatus = true;
                            objNatureOfComplaint.ResultMessage = "Success";
                            lstNatureOfComplaint.Add(objNatureOfComplaint);
                        }
                    }
                    return lstNatureOfComplaint;
                }
            }
            catch (Exception ex)
            {
                objNatureOfComplaint = new NatureOfComplaint { ResultStatus = false, ResultMessage = ex.Message };
                lstNatureOfComplaint.Add(objNatureOfComplaint);
                return lstNatureOfComplaint;
            }
        }

        public List<NatureOfComplaint> GetAllNOCs()
        {
            NatureOfComplaint objNatureOfComplaint;
            List<NatureOfComplaint> lstNatureOfComplaint = new List<NatureOfComplaint>();
            EntityCollection entcoll;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='hil_natureofcomplaint'>
                        <attribute name='hil_name' />
                        <attribute name='hil_natureofcomplaintid' />
                        <attribute name='hil_relatedproduct' />
                        <order attribute='hil_name' descending='false' />
                    </entity>
                    </fetch>";
                entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (entcoll.Entities.Count == 0)
                {
                    objNatureOfComplaint = new NatureOfComplaint { ResultStatus = false, ResultMessage = "No Nature of Complaint is mapped with Serial Number." };
                    lstNatureOfComplaint.Add(objNatureOfComplaint);
                }
                else
                {
                    foreach (Entity ent in entcoll.Entities)
                    {
                        objNatureOfComplaint = new NatureOfComplaint();
                        objNatureOfComplaint.Guid = ent.GetAttributeValue<Guid>("hil_natureofcomplaintid");
                        objNatureOfComplaint.Name = ent.GetAttributeValue<string>("hil_name");
                        if (ent.Contains("hil_relatedproduct") || ent.Attributes.Contains("hil_relatedproduct"))
                        {
                            objNatureOfComplaint.ProductCategoryGuid = ent.GetAttributeValue<EntityReference>("hil_relatedproduct").Id;
                            objNatureOfComplaint.ProductCategoryName = ent.GetAttributeValue<EntityReference>("hil_relatedproduct").Name;
                        }
                        objNatureOfComplaint.ResultStatus = true;
                        objNatureOfComplaint.ResultMessage = "Success";
                        lstNatureOfComplaint.Add(objNatureOfComplaint);
                    }
                }
                return lstNatureOfComplaint;
            }
            catch (Exception ex)
            {
                objNatureOfComplaint = new NatureOfComplaint { ResultStatus = false, ResultMessage = ex.Message };
                lstNatureOfComplaint.Add(objNatureOfComplaint);
                return lstNatureOfComplaint;
            }
        }
    }
}
