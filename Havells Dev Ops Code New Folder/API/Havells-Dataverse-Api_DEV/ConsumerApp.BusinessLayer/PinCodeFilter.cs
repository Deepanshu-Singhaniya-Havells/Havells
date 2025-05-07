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
    public class PinCodeFilter
    {
        [DataMember(IsRequired = false)]
        public string PinCode { get; set; }
        [DataMember(IsRequired = false)]
        public string PinCodesName { get; set; }
        [DataMember(IsRequired = false)]
        public string PinCodeGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string StateName { get; set; }
        [DataMember(IsRequired = false)]
        public string StateGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string CityName { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_ID { get; set; }
        [DataMember]
        public string CityGuId { get; set; }
        public List<PinCodeFilter> GetAllActivePinCodes(PinCodeFilter Pin)
        {
            List<PinCodeFilter> obj = new List<PinCodeFilter>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                QueryExpression Qry = new QueryExpression();
                Qry.EntityName = hil_businessmapping.EntityLogicalName;
                ColumnSet Col = new ColumnSet("hil_city", "hil_state", "hil_pincode", "hil_district");
                Qry.ColumnSet = Col;
                Qry.Criteria = new FilterExpression(LogicalOperator.And);
                Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_city", ConditionOperator.Equal, new Guid(Pin.CityGuId)));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_state", ConditionOperator.NotNull));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.NotNull));
                Qry.Distinct = true;
                EntityCollection Colec = service.RetrieveMultiple(Qry);
                if (Colec.Entities.Count >= 1)
                {
                    foreach (hil_businessmapping et in Colec.Entities)
                    {
                        EntityReference State = new EntityReference();
                        EntityReference City = new EntityReference();
                        EntityReference District = new EntityReference();
                        if (et.Attributes.Contains("hil_state"))
                        {
                            State = et.hil_state;
                        }
                        if(et.Attributes.Contains("hil_district"))
                        {
                            District = et.hil_district;
                        }
                        City = et.hil_city;
                        obj.Add(
                        new PinCodeFilter
                        {
                            StateGuId = Convert.ToString(State.Id),
                            StateName = State.Name,
                            CityGuId = Convert.ToString(City.Id),
                            CityName = City.Name,
                            PinCodeGuId = Convert.ToString(et.hil_pincode.Id),
                            PinCodesName = Convert.ToString(et.hil_pincode.Name),
                            DISTRICT_ID = District.Id.ToString(),
                            DISTRICT_NAME = District.Name
                        });
                    }
                }
                else
                {
                    obj.Add(
                        new PinCodeFilter
                        {
                            StateGuId = "",
                            StateName = "",
                            CityGuId = "",
                            CityName = "",
                            PinCodeGuId = "",
                            PinCodesName = ""
                        });
                }
            }
            catch(Exception ex)
            {
                obj.Add(
                        new PinCodeFilter
                        {
                            StateGuId = ex.Message.ToUpper(),
                            StateName = "",
                            CityGuId = "",
                            CityName = "",
                            PinCodeGuId = "",
                            PinCodesName = ""
                        });
            }
            return (obj);
        }
    }
}