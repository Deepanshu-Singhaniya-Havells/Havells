using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class UnFiltered_District
    {
        [DataMember(IsRequired = false)]
        public string DISTRICT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string STATE_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string STATE_ID { get; set; }
        public List<UnFiltered_District> GetListofDistricts()
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            List<UnFiltered_District> obj = new List<UnFiltered_District>();
            try
            {
                QueryExpression Qry = new QueryExpression();
                Qry.EntityName = hil_businessmapping.EntityLogicalName;
                ColumnSet Col = new ColumnSet("hil_district", "hil_state");
                Qry.ColumnSet = Col;
                Qry.Criteria = new FilterExpression(LogicalOperator.And);
                Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Qry.Distinct = true;
                EntityCollection Found = service.RetrieveMultiple(Qry);
                foreach (hil_businessmapping et in Found.Entities)
                {
                    obj.Add(
                        new UnFiltered_District
                        {
                            DISTRICT_ID = et.hil_district.Id.ToString(),
                            DISTRICT_NAME = et.hil_district.Name,
                            STATE_ID = et.hil_state.Id.ToString(),
                            STATE_NAME  =et.hil_state.Name
                        });
                }
            }
            catch (Exception ex)
            {
                obj.Add
                    (
                        new UnFiltered_District
                        {
                            DISTRICT_ID = ex.Message.ToUpper(),
                            DISTRICT_NAME = ex.Message.ToUpper()
                        });
            }
            return (obj);
        }
    }
}
