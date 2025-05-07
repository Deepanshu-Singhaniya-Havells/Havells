using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Deployment;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class Job
    {
        [DataMember]
        public string Job_ID { get; set; }

        [DataMember]
        public string MobileNumber { get; set; }


        public List<JobOutput> getJobs(Job job)
        {
            List<JobOutput> jobList = new List<JobOutput>();

            if (String.IsNullOrWhiteSpace(job.MobileNumber))
            {
                return jobList;
            }
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression query = new QueryExpression()
            {
                EntityName = msdyn_workorder.EntityLogicalName,
                ColumnSet = new ColumnSet(true)
            };
            FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
            if (!String.IsNullOrWhiteSpace(job.Job_ID))
            {
                filterExpression.Conditions.Add(new ConditionExpression("msdyn_name", ConditionOperator.Equal, job.Job_ID));
            }
            filterExpression.Conditions.Add(new ConditionExpression("hil_mobilenumber", ConditionOperator.Equal, job.MobileNumber));

            query.Criteria.AddFilter(filterExpression);

            EntityCollection collection = service.RetrieveMultiple(query);

            if (collection.Entities != null && collection.Entities.Count > 0)
            {
                foreach (Entity item in collection.Entities)
                {
                    JobOutput jobObj = new JobOutput();
                    if (item.Attributes.Contains("msdyn_customerasset"))
                    {
                        jobObj.Job_Asset = item.GetAttributeValue<EntityReference>("msdyn_customerasset").Name;
                    }
                    if (item.Attributes.Contains("msdyn_name"))
                    {
                        jobObj.Job_ID = item.GetAttributeValue<string>("msdyn_name");
                    }
                    if (item.Attributes.Contains("msdyn_substatus"))
                    {
                        jobObj.Job_Status = item.GetAttributeValue<EntityReference>("msdyn_substatus").Name;
                    }
                    if (item.Attributes.Contains("hil_productcategory"))
                    {
                        jobObj.Job_Category = item.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                    }
                    if (item.Attributes.Contains("createdon"))
                    {
                        jobObj.Job_Loggedon = item.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                    }
                    if (item.Attributes.Contains("hil_jobclosuredon"))
                    {
                        jobObj.Job_ClosedOn = item.GetAttributeValue<DateTime>("hil_jobclosuredon").AddMinutes(330).ToString();
                    }
                    if (item.Attributes.Contains("hil_mobilenumber"))
                    {
                        jobObj.MobileNumber = item.GetAttributeValue<string>("hil_mobilenumber");
                    }
                    if (item.Attributes.Contains("hil_fulladdress"))
                    {
                        jobObj.Customer_Address = item.GetAttributeValue<string>("hil_fulladdress");
                    }
                    if (item.Attributes.Contains("hil_customerref"))
                    {
                        jobObj.Customer_name = item.GetAttributeValue<EntityReference>("hil_customerref").Name;
                    }
                    if (item.Attributes.Contains("hil_owneraccount"))
                    {
                        jobObj.Job_AssignedTo = item.GetAttributeValue<EntityReference>("hil_owneraccount").Name;
                    }
                    if (item.Attributes.Contains("hil_callsubtype"))
                    {
                        jobObj.CallSubType = item.GetAttributeValue<EntityReference>("hil_callsubtype").Name;
                    }
                    if (item.Attributes.Contains("hil_customercomplaintdescription"))
                    {
                        jobObj.ChiefComplaint = item.GetAttributeValue<string>("hil_customercomplaintdescription");
                    }
                    if (item.Attributes.Contains("msdyn_customerasset"))
                    {
                        Entity ec = service.Retrieve("msdyn_customerasset", item.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_modelname"));
                        if (ec != null)
                        {
                            jobObj.Product = ec.GetAttributeValue<string>("hil_modelname");
                        }
                    }
                    if (item.Attributes.Contains("hil_productcategory"))
                    {
                        jobObj.ProductCategoryGuid = item.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                        jobObj.ProductCategoryName = item.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                    }
                    jobList.Add(jobObj);
                }
            }
            return jobList;
        }
    }

    [DataContract]
    public class JobList
    {
        public List<Job> Jobs { get; set; }
        public string message { get; set; }
    }

    [DataContract]
    public class JobOutput
    {
        [DataMember]
        public string Job_ID { get; set; }

        [DataMember]
        public string MobileNumber { get; set; }

        [DataMember]
        public string CallSubType { get; set; }

        [DataMember]
        public string Job_Loggedon { get; set; }

        [DataMember]
        public string Job_Status { get; set; }

        [DataMember]
        public string Job_AssignedTo { get; set; }

        [DataMember]
        public string Job_Asset { get; set; }

        [DataMember]
        public string Job_Category { get; set; }

        [DataMember]
        public string Job_NOC { get; set; }

        [DataMember]
        public string Job_ClosedOn { get; set; }

        [DataMember]
        public string Customer_name { get; set; }

        [DataMember]
        public string Customer_Address { get; set; }

        [DataMember]
        public string Product { get; set; }

        [DataMember]
        public Guid ProductCategoryGuid { get; set; }

        [DataMember]
        public string ProductCategoryName { get; set; }

        [DataMember]
        public string ChiefComplaint { get; set; }
    }
}
