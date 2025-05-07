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

namespace TestApp
{
    public class FlushUATData
    {
        public string MobileNumber { get; set; }
        public string EntryType { get; set; }

        public FlushUATDataResult FlushData(FlushUATData flushUATData, IOrganizationService service)
        {
            FlushUATDataResult objFlushUATDataResult = null;
            EntityCollection entcoll;
            EntityCollection entcoll1;
            QueryExpression Query = null;
            try
            {
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (flushUATData.MobileNumber == null || flushUATData.MobileNumber.Trim().Length == 0)
                    {
                        objFlushUATDataResult = new FlushUATDataResult { StatusCode = "204", StatusDescription = "Customer Mobile Number is required." };
                        return objFlushUATDataResult;
                    }
                    //if (flushUATData.EntryType == null || flushUATData.EntryType.ToString().Trim().Length == 0)
                    //{
                    //    objFlushUATDataResult = new FlushUATDataResult { StatusCode = "204", StatusDescription = "Entry Type is required. !!!\n Please select Entry type : <J> Jobs <P> Product <A> Address <C> Customer" };
                    //    return objFlushUATDataResult;
                    //}
                    int i = 0;
                    Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("contactid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, flushUATData.MobileNumber);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count > 0)
                    {
                        if (flushUATData.EntryType == "J" || flushUATData.EntryType == "") // Registered Jobs
                        {
                            Query = new QueryExpression(msdyn_workorder.EntityLogicalName);
                            Query.ColumnSet = new ColumnSet("msdyn_name");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_customerref", ConditionOperator.Equal, entcoll.Entities[0].Id);
                            entcoll1 = service.RetrieveMultiple(Query);
                            if (entcoll1.Entities.Count > 0)
                            {
                                i = 1;
                                foreach (msdyn_workorder ent in entcoll1.Entities)
                                {
                                    service.Delete(msdyn_workorder.EntityLogicalName, ent.Id);
                                    Console.WriteLine("Job# " + i++.ToString() + "/" + entcoll1.Entities.Count.ToString());
                                }
                            }
                        }
                        if (flushUATData.EntryType == "P" || flushUATData.EntryType == "")//Registered Product
                        {
                            Query = new QueryExpression(msdyn_customerasset.EntityLogicalName);
                            Query.ColumnSet = new ColumnSet("msdyn_customerassetid");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, entcoll.Entities[0].Id);
                            entcoll1 = service.RetrieveMultiple(Query);
                            if (entcoll1.Entities.Count > 0)
                            {
                                i = 1;
                                foreach (msdyn_customerasset ent in entcoll1.Entities)
                                {
                                    service.Delete(msdyn_customerasset.EntityLogicalName, ent.Id);
                                    Console.WriteLine("Product# " + i++.ToString() + "/" + entcoll1.Entities.Count.ToString());
                                }
                            }
                        }
                        if (flushUATData.EntryType == "A" || flushUATData.EntryType == "") //Address
                        {
                            Query = new QueryExpression("hil_address");
                            Query.ColumnSet = new ColumnSet("hil_customer");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, entcoll.Entities[0].Id);
                            entcoll1 = service.RetrieveMultiple(Query);
                            if (entcoll1.Entities.Count > 0)
                            {
                                i = 1;
                                foreach (hil_address ent in entcoll1.Entities)
                                {
                                    service.Delete("hil_address", ent.Id);
                                    Console.WriteLine("Address# " + i++.ToString() + "/" + entcoll1.Entities.Count.ToString());
                                }
                            }
                        }
                        if (flushUATData.EntryType == "C" || flushUATData.EntryType == "") // Customer
                        {
                            service.Delete("contact", entcoll.Entities[0].Id);
                        }
                        else
                        {
                            objFlushUATDataResult = new FlushUATDataResult { StatusCode = "204", StatusDescription = "Invalid Entry Type." };
                            return objFlushUATDataResult;
                        }
                        Console.WriteLine("DONE!!!");
                    }
                    else
                    {
                        objFlushUATDataResult = new FlushUATDataResult { StatusCode = "204", StatusDescription = "Customer does not exist." };
                        return objFlushUATDataResult;
                    }
                }
            }
            catch (Exception ex)
            {
                objFlushUATDataResult = new FlushUATDataResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objFlushUATDataResult;
        }
    }

    public class FlushUATDataResult
    {
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
}
