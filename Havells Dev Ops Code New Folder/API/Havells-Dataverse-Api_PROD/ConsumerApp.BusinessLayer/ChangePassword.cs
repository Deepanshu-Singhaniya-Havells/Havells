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
    public class ChangePassword
    {
        [DataMember]
        public string UName { get; set; }
        [DataMember]
        public string Pwd { get; set; }
        [DataMember]
        public string Method { get; set; }
        [DataMember(IsRequired = false)]
        public string ContGuid { get; set; }
        [DataMember(IsRequired = false)]
        public bool status { get; set; }
        [DataMember(IsRequired = false)]
        public bool IfValidated { get; set; }
        [DataMember(IsRequired = false)]
        public string ErrCode { get; set; }
        [DataMember(IsRequired = false)]
        public string ErrDesc { get; set; }

        public ChangePassword SetNewPassword(ChangePassword bridge)
        {
            //IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            //OrganizationRequest req = new OrganizationRequest("hil_ConsumerApp_ChangePassword2a02978f9973e811a95a000d3af068d4");
            //req["Method"] = bridge.Method;
            //req["UserName"] = bridge.UName;
            //req["Password"] = bridge.Pwd;
            //OrganizationResponse response = service.Execute(req);
            //bridge.ContGuid = (string)response["ContGuid"];
            //bridge.status = (bool)response["Status"];
            //bridge.ErrCode = (string)response["ErrorCode"];
            //bridge.ErrDesc = (string)response["ErrorDescription"];

            bridge.status = false;
            bridge.ErrCode = "404";
            bridge.ErrDesc = "Access Denied!!!";
            return (bridge);
        }
    }
}
