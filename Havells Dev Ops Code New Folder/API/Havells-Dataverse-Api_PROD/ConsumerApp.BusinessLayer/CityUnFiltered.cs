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
using Microsoft.Xrm.Sdk.Deployment;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class CityUnFiltered
    {
        [DataMember(IsRequired = false)]
        public string CityNames { get; set; }
        [DataMember(IsRequired = false)]
        public string CityGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string StateGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string StateName { get; set; }
        public List<CityUnFiltered> GetAllActiveCitiesUnFiltered()
        {
            List<CityUnFiltered> obj = new List<CityUnFiltered>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            string State_Name = string.Empty;
            string State_Guid = string.Empty;
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = hil_city.EntityLogicalName;
            ColumnSet Col = new ColumnSet("hil_name", "hil_cityid");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            foreach (hil_city et in Colec.Entities)
            {
                if(et.hil_State != null)
                {
                    State_Name = et.hil_State.Name;
                    State_Guid = et.hil_State.Id.ToString();

                }
                obj.Add(
                new CityUnFiltered
                {
                    CityNames = et.hil_name,
                    CityGuId = et.hil_cityId.ToString(),
                    StateGuId = State_Guid,
                    StateName = State_Name
                });
            }
            return (obj);
        }
    }
}
