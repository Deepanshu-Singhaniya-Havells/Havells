using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class DistrictMaster
    {
        [DataMember]
        public string STATE_ID { get; set; }
        public List<OutputDistrictMaster> GetListofDistricts(DistrictMaster iMaster)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            List<OutputDistrictMaster> obj = new List<OutputDistrictMaster>();
            try
            {
                QueryExpression Qry = new QueryExpression();
                Qry.EntityName = hil_businessmapping.EntityLogicalName;
                ColumnSet Col = new ColumnSet("hil_district");
                Qry.ColumnSet = Col;
                Qry.Criteria = new FilterExpression(LogicalOperator.And);
                Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_state", ConditionOperator.Equal, new Guid(iMaster.STATE_ID)));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_district", ConditionOperator.NotNull));
                Qry.Distinct = true;
                EntityCollection Found = service.RetrieveMultiple(Qry);
                foreach (hil_businessmapping et in Found.Entities)
                {
                    obj.Add(
                        new OutputDistrictMaster
                        {
                            DISTRICT_ID = et.hil_district.Id.ToString(),
                            DISTRICT_NAME = et.hil_district.Name
                        });
                }
            }
            catch (Exception ex)
            {
                obj.Add
                    (
                        new OutputDistrictMaster
                        {
                            DISTRICT_ID = ex.Message.ToUpper(),
                            DISTRICT_NAME = ex.Message.ToUpper()
                        });
            }
            return (obj);
        }
    }
    public class OutputDistrictMaster
    {
        [DataMember(IsRequired = false)]
        public string DISTRICT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_NAME { get; set; }
    }
}