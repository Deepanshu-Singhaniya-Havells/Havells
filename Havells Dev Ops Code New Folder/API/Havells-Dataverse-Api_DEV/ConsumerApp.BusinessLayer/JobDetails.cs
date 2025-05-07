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
    public class JobDetails
    {
        [DataMember]
        public string JOB_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string JOB_STATUS { get; set; }
        [DataMember(IsRequired = false)]
        public string ASSIGNED_TO { get; set; }
        [DataMember(IsRequired = false)]
        public string TYPE_OF_OWNER { get; set; }
        [DataMember(IsRequired = false)]
        public DateTime CREATED_ON { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_SUB_CAT { get; set; }
        [DataMember(IsRequired = false)]
        public string CALL_STYPE { get; set; }
        [DataMember(IsRequired = false)]
        public string CONSUMER_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string MOBILE_NO { get; set; }
        [DataMember(IsRequired = false)]
        public string CONSUMER_ADDRESS { get; set; }
        public JobDetails iGetJobDetails(JobDetails iJobDetails)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                msdyn_workorder iJob = new msdyn_workorder();
                QueryExpression Qry = new QueryExpression();
                Qry.EntityName = msdyn_workorder.EntityLogicalName;
                ColumnSet Col = new ColumnSet(true);
                Qry.ColumnSet = Col;
                Qry.Criteria = new FilterExpression(LogicalOperator.And);
                Qry.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, iJobDetails.JOB_ID));
                EntityCollection Colec = service.RetrieveMultiple(Qry);
                {
                    if (Colec.Entities.Count > 0)
                    {
                        iJob = Colec[0].ToEntity<msdyn_workorder>();
                        iJobDetails.JOB_STATUS = iJob.msdyn_SubStatus.Name;
                        iJobDetails.ASSIGNED_TO = iJob.OwnerId.Name;
                        iJobDetails.TYPE_OF_OWNER = iJob.hil_typeofassignee.Name;
                        iJobDetails.CREATED_ON = Convert.ToDateTime(iJob.CreatedOn);
                        iJobDetails.PROD_SUB_CAT = iJob.hil_ProductSubcategory.Name;
                        iJobDetails.CALL_STYPE = iJob.hil_CallSubType.Name;
                        iJobDetails.MOBILE_NO = (!String.IsNullOrEmpty(iJob.hil_mobilenumber)) ? iJob.hil_mobilenumber : string.Empty;
                        iJobDetails.CONSUMER_NAME = (iJob.hil_CustomerRef != null) ? iJob.hil_CustomerRef.Name : string.Empty;
                        iJobDetails.CONSUMER_ADDRESS = (!String.IsNullOrEmpty(iJob.hil_FullAddress)) ? iJob.hil_FullAddress : string.Empty;
                    }
                    else
                    {
                        iJobDetails.JOB_ID = "JOB NOT FOUND";
                        iJobDetails.JOB_STATUS = "JOB NOT FOUND";
                        iJobDetails.ASSIGNED_TO = "JOB NOT FOUND";
                        iJobDetails.TYPE_OF_OWNER = "JOB NOT FOUND";
                        iJobDetails.CREATED_ON = new DateTime();
                        iJobDetails.PROD_SUB_CAT = "JOB NOT FOUND";
                        iJobDetails.CALL_STYPE = "JOB NOT FOUND";
                    }
                }
            }
            catch(Exception ex)
            {
                iJobDetails.JOB_ID = ex.Message.ToUpper();
                iJobDetails.JOB_STATUS = ex.Message.ToUpper();
                iJobDetails.ASSIGNED_TO = ex.Message.ToUpper();
                iJobDetails.TYPE_OF_OWNER = ex.Message.ToUpper();
                iJobDetails.CREATED_ON = new DateTime();
                iJobDetails.PROD_SUB_CAT = ex.Message.ToUpper();
                iJobDetails.CALL_STYPE = ex.Message.ToUpper();
            }
            return iJobDetails;
        }
    }
}
