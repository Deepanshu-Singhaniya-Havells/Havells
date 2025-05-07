using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class CityMaster
    {
        [DataMember]
        public string StateGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string CityNames { get; set; }
        [DataMember(IsRequired = false)]
        public string CityGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string StateNames { get; set; }
        public List<CityMaster> GetAllActiveCities(CityMaster City)
        {
            List<CityMaster> obj = new List<CityMaster>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = hil_businessmapping.EntityLogicalName;
            ColumnSet Col = new ColumnSet("hil_city", "hil_district", "hil_state");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            if(City.DISTRICT_ID != null)
            {
                Qry.Criteria.AddCondition(new ConditionExpression("hil_district", ConditionOperator.Equal, new Guid(City.DISTRICT_ID)));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_state", ConditionOperator.Equal, new Guid(City.StateGuId)));
            }
            else
            {
                Qry.Criteria.AddCondition(new ConditionExpression("hil_state", ConditionOperator.Equal, new Guid(City.StateGuId)));
            }
            
            Qry.Distinct = true;
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            foreach (hil_businessmapping et in Colec.Entities)
            {
                try
                {
                    EntityReference State = (EntityReference)et.hil_state;
                    obj.Add(
                    new CityMaster
                    {
                        StateGuId = Convert.ToString(State.Id),
                        StateNames = State.Name,
                        CityNames = Convert.ToString(et.hil_city.Name),
                        CityGuId = Convert.ToString(et.hil_city.Id),
                        DISTRICT_ID = Convert.ToString(et.hil_district.Id),
                        DISTRICT_NAME = et.hil_district.Name
                    });
                }
                catch
                {
                    continue;
                }
            }
            return (obj);
        }
    }
}