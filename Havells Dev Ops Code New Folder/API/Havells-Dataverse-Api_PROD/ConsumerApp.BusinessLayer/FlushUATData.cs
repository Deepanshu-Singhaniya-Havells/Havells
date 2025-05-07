using System.Collections.Generic;
using System;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Runtime.Serialization.Json;
using System.IO;
using Havells_Plugin.HelperIntegration;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class FlushUATData
    {
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string EntryType { get; set; }

        public FlushUATDataResult FlushData(FlushUATData flushUATData)
        {
            FlushUATDataResult objFlushUATDataResult = null;
            EntityCollection entcoll;
            EntityCollection entcoll1;
            QueryExpression Query = null;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (DateTime.Now.Year != 2021)
                    {
                        objFlushUATDataResult = new FlushUATDataResult { StatusCode = "204", StatusDescription = "Service is not available." };
                        return objFlushUATDataResult;
                    }
                    if (flushUATData.MobileNumber == null || flushUATData.MobileNumber.Trim().Length == 0)
                    {
                        objFlushUATDataResult = new FlushUATDataResult { StatusCode = "204", StatusDescription = "Customer Mobile Number is required." };
                        return objFlushUATDataResult;
                    }
                    if (flushUATData.EntryType == null || flushUATData.EntryType.ToString().Trim().Length == 0)
                    {
                        objFlushUATDataResult = new FlushUATDataResult { StatusCode = "204", StatusDescription = "Entry Type is required. !!!\n Please select Entry type : <J> Jobs <P> Product <A> Address <C> Customer" };
                        return objFlushUATDataResult;
                    }

                    Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("contactid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, flushUATData.MobileNumber);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count >0)
                    {
                        if (flushUATData.EntryType == "P")//Registered Product
                        {
                            Query = new QueryExpression(msdyn_customerasset.EntityLogicalName);
                            Query.ColumnSet = new ColumnSet("msdyn_customerassetid");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, entcoll.Entities[0].Id);
                            entcoll1 = service.RetrieveMultiple(Query);
                            if (entcoll1.Entities.Count > 0)
                            {
                                foreach (msdyn_customerasset ent in entcoll1.Entities)
                                {
                                    service.Delete(msdyn_customerasset.EntityLogicalName, ent.Id);
                                }
                                objFlushUATDataResult = new FlushUATDataResult { StatusCode = "200", StatusDescription = "OK" };
                                return objFlushUATDataResult;
                            }
                        }
                        else if (flushUATData.EntryType == "J") // Registered Jobs
                        {
                            Query = new QueryExpression(msdyn_workorder.EntityLogicalName);
                            Query.ColumnSet = new ColumnSet("msdyn_name");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_customerref", ConditionOperator.Equal, entcoll.Entities[0].Id);
                            entcoll1 = service.RetrieveMultiple(Query);
                            if (entcoll1.Entities.Count > 0)
                            {
                                foreach (msdyn_workorder ent in entcoll1.Entities)
                                {
                                    service.Delete(msdyn_workorder.EntityLogicalName, ent.Id);
                                }
                                objFlushUATDataResult = new FlushUATDataResult { StatusCode = "200", StatusDescription = "OK" };
                                return objFlushUATDataResult;
                            }
                        }
                        else if (flushUATData.EntryType == "C") // Customer
                        {
                            service.Delete("contact", entcoll.Entities[0].Id);
                            objFlushUATDataResult = new FlushUATDataResult { StatusCode = "200", StatusDescription = "OK" };
                            return objFlushUATDataResult;
                        }
                        else if (flushUATData.EntryType == "A") //Address
                        {
                            Query = new QueryExpression("hil_address");
                            Query.ColumnSet = new ColumnSet("hil_customer");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, entcoll.Entities[0].Id);
                            entcoll1 = service.RetrieveMultiple(Query);
                            if (entcoll1.Entities.Count > 0)
                            {
                                foreach (hil_address ent in entcoll1.Entities)
                                {
                                    service.Delete("hil_address", ent.Id);
                                }
                                objFlushUATDataResult = new FlushUATDataResult { StatusCode = "200", StatusDescription = "OK" };
                                return objFlushUATDataResult;
                            }
                        }
                        else {
                            objFlushUATDataResult = new FlushUATDataResult { StatusCode = "204", StatusDescription = "Invalid Entry Type." };
                            return objFlushUATDataResult;
                        }
                    }
                    else {
                        objFlushUATDataResult = new FlushUATDataResult { StatusCode = "204", StatusDescription = "Customer does not exist." };
                        return objFlushUATDataResult;
                    }
                }
            }
            catch (Exception ex)
            {
                objFlushUATDataResult = new FlushUATDataResult {StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objFlushUATDataResult;
        }
    }

    [DataContract]
    public class FlushUATDataResult
    {
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
}
