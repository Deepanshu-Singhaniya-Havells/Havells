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
    public class PinCodeMasterData
    {
        [DataMember(IsRequired = false)]
        public string PinCodesName { get; set; }
        [DataMember(IsRequired = false)]
        public string PinCodeGuId { get; set; }
        public List<PinCodeMasterData> GetAllActivePinCodesMaster()
        {
            List<PinCodeMasterData> obj = new List<PinCodeMasterData>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "hil_pincode";
            ColumnSet Col = new ColumnSet("hil_name", "hil_pincodeid");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            if (Colec.Entities.Count > 0)
            {
                foreach (Entity et in Colec.Entities)
                {
                    obj.Add(
                    new PinCodeMasterData
                    {
                        PinCodeGuId = Convert.ToString(et["hil_pincodeid"]),
                        PinCodesName = Convert.ToString(et["hil_name"])
                    });
                }
            }
            else
            {
                obj.Add(
                    new PinCodeMasterData
                    {
                        PinCodeGuId = "",
                        PinCodesName = ""
                    });
            }
            return (obj);
        }
    }
}
