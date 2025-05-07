using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class TechnicianInfo
    {
        [DataMember]
        public string ImageBase64String { get; set; }
        [DataMember]
        public string Designtion { get; set; }
        [DataMember]
        public string CareID { get; set; }
        [DataMember]
        public string MobileNo { get; set; }
        [DataMember]
        public string BloodGroup { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public string Status { get; set; }

        public TechnicianInfo GetTechnicianInfo(TechnicianInfo care360ID)
        {
            TechnicianInfo technician = new TechnicianInfo();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                QueryExpression queryExpression = new QueryExpression();
                queryExpression.EntityName = "hil_technician";
                ColumnSet columnSet = new ColumnSet(
                    "entityimage",
                    "hil_care360id",
                    "hil_firstname",
                    "hil_lastname",
                    "hil_mobilenum",
                    "hil_bloodgroup",
                    "statecode",
                    "createdby"
                    );
                queryExpression.ColumnSet = columnSet;
                queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
                queryExpression.Criteria.AddCondition(new ConditionExpression("hil_care360id", ConditionOperator.Equal, care360ID.CareID));
                EntityCollection collection = service.RetrieveMultiple(queryExpression);

                if (collection.Entities.Count == 1)
                {

                    foreach (Entity entity in collection.Entities)
                    {
                        if (entity.Attributes.Contains("entityimage"))
                        {
                            technician.ImageBase64String = Convert.ToBase64String(entity.GetAttributeValue<byte[]>("entityimage"));
                        }
                        if (entity.Attributes.Contains("hil_care360id"))
                        {
                            technician.CareID = Convert.ToString(entity["hil_care360id"]);
                        }
                        if (entity.Attributes.Contains("hil_firstname"))
                        {
                            technician.Name = Convert.ToString(entity["hil_firstname"]);
                        }
                        if (entity.Attributes.Contains("hil_lastname"))
                        {
                            technician.Name = technician.Name + " " + Convert.ToString(entity["hil_lastname"]);
                        }
                        if (entity.Attributes.Contains("createdby"))
                        {
                            QueryExpression queryExpression1 = new QueryExpression();
                            queryExpression1.EntityName = "systemuser";
                            queryExpression1.ColumnSet = new ColumnSet("mobilephone");
                            queryExpression1.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExpression1.Criteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("createdby").Id));
                            EntityCollection collection1 = service.RetrieveMultiple(queryExpression1);
                            if (collection1.Entities.Count == 1)
                            {
                                technician.MobileNo = collection1.Entities[0].GetAttributeValue<string>("mobilephone");
                            }
                        }
                        if (entity.Attributes.Contains("hil_bloodgroup"))
                        {
                            OptionSetValue bloodGroup = entity.GetAttributeValue<OptionSetValue>("hil_bloodgroup");
                            if (bloodGroup.Value == 1) {
                                technician.BloodGroup = "A+";
                            }
                            else if(bloodGroup.Value == 2) {
                                technician.BloodGroup = "B+";
                            }
                            else if (bloodGroup.Value == 3)
                            {
                                technician.BloodGroup = "AB+";
                            }
                            else if (bloodGroup.Value == 4)
                            {
                                technician.BloodGroup = "O+";
                            }
                            else if (bloodGroup.Value == 5)
                            {
                                technician.BloodGroup = "A-";
                            }
                            else if (bloodGroup.Value == 6)
                            {
                                technician.BloodGroup = "B-";
                            }
                            else if (bloodGroup.Value == 7)
                            {
                                technician.BloodGroup = "AB-";
                            }
                            else if (bloodGroup.Value == 8)
                            {
                                technician.BloodGroup = "O-";
                            }

                            //RetrieveAttributeRequest Req = new RetrieveAttributeRequest
                            //{
                            //    EntityLogicalName = entity.LogicalName,
                            //    LogicalName = "hil_bloodgroup",
                            //    RetrieveAsIfPublished = true
                            //};
                            //RetrieveAttributeResponse Resp = (RetrieveAttributeResponse)service.Execute(Req);
                            //EnumAttributeMetadata data = (EnumAttributeMetadata)Resp.AttributeMetadata;
                            //foreach (OptionMetadata picklist in data.OptionSet.Options)
                            //{
                            //    technician.BloodGroup = picklist.Label.UserLocalizedLabel.Label;
                            //}

                        }
                        if (entity.Attributes.Contains("statecode"))
                        {
                            int status = entity.GetAttributeValue<OptionSetValue>("statecode").Value;
                            technician.Status = status == 0 ? "ACTIVE" : "INACTIVE";
                        }
                        technician.Designtion = "Technician";
                    }
                }
                else
                {
                    technician.CareID = "INVALID";
                    technician.Status = "INVALID";
                }
            }
            catch (Exception ex)
            {
                technician.CareID = "ERROR";
                technician.Status = "ERROR";
            }
            return (technician);
        }
    }
}
