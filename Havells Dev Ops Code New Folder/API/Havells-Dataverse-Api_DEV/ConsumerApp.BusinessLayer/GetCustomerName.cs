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
    public class GetCustomerName
    {
        [DataMember]
        public string MOBILE { get; set; }
        [DataMember(IsRequired = false)]
        public string FULLNAME { get; set; }
        [DataMember(IsRequired = false)]
        public string LANGUAGE { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT_DESC { get; set; }
        public GetCustomerName GetFullNameBasisMobNo(GetCustomerName Cust)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "contact";
            ColumnSet Col = new ColumnSet("fullname", "hil_language");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, Cust.MOBILE));
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            if (Colec.Entities.Count == 1)
            {
                foreach (Entity et in Colec.Entities)
                {
                    if(et.Attributes.Contains("fullname"))
                    {
                        Cust.FULLNAME = Convert.ToString(et["fullname"]);
                    }
                    if(et.Attributes.Contains("hil_language"))
                    {
                        OptionSetValue LangCode = (OptionSetValue)et["hil_language"];
                        if(LangCode.Value == 1)
                        {
                            Cust.LANGUAGE = "ENGLISH";
                        }
                        else if(LangCode.Value == 2)
                        {
                            Cust.LANGUAGE = "HINDI";
                        }
                    }
                    Cust.RESULT = "SUCCESS";
                    Cust.RESULT_DESC = "";
                }
            }
            else
            {
                Cust.FULLNAME = "";
                Cust.LANGUAGE = "";
                Cust.RESULT = "FAILURE";
                Cust.RESULT_DESC = "CUSTOMER NOT FOUND";
            }
            return (Cust);
        }
    }
}