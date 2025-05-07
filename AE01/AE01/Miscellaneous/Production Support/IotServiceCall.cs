using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Reflection;

namespace AE01.Miscellaneous.Production_Support
{
    internal class IotServiceCall
    {
        private IOrganizationService service;
        public IotServiceCall(IOrganizationService _service)
        {
            this.service = _service;
        }

        public List<IoTServiceCallResult> GetIoTServiceCalls(Guid jobId)
        {
            var watch = new System.Diagnostics.Stopwatch();
            watch.Start();
            List<IoTServiceCallResult> list = new List<IoTServiceCallResult>();
            try
            {
                IoTServiceCallResult item;
                if (jobId.ToString().Trim().Length == 0)
                {
                    item = new IoTServiceCallResult
                    {
                        StatusCode = "204",
                        StatusDescription = "Customer GUID is required."
                    };
                    list.Add(item);
                    return list;
                }
                if (service != null)
                {
                    QueryExpression queryExpression = new QueryExpression
                    {
                        EntityName = "msdyn_workorder",
                        ColumnSet = new ColumnSet("msdyn_workorderid", "msdyn_name", "hil_jobclosuredon", "msdyn_substatus", "hil_owneraccount", "hil_productcategory", "hil_natureofcomplaint", "hil_fulladdress", "hil_customerref", "createdon", "hil_callsubtype", "msdyn_customerasset", "hil_preferredtime", "hil_preferreddate", "hil_customercomplaintdescription")
                    };
                    FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
                    filterExpression.Conditions.Add(new ConditionExpression("hil_customerref", ConditionOperator.Equal, jobId));
                    queryExpression.Criteria.AddFilter(filterExpression);
                    queryExpression.AddOrder("createdon", OrderType.Descending);
                    LinkEntity customerAssetLink = queryExpression.AddLink(
                        linkToEntityName: "msdyn_customerasset",
                        linkFromAttributeName: "msdyn_customerasset",
                        linkToAttributeName: "msdyn_customerassetid", 
                        joinOperator: JoinOperator.LeftOuter);

                    customerAssetLink.Columns.AddColumn("hil_modelname");
                    customerAssetLink.EntityAlias = "asset"; 
                    Console.WriteLine("Before Retriving: " + watch.ElapsedMilliseconds);
                    EntityCollection entityCollection = service.RetrieveMultiple(queryExpression);
                    Console.WriteLine("After Retriving: " + watch.ElapsedMilliseconds);


                    int i = 1;
                    if (entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                    {
                        foreach (Entity entity2 in entityCollection.Entities)
                        {
                            Console.WriteLine("For record " + i + " " + watch.ElapsedMilliseconds);

                            IoTServiceCallResult ioTServiceCallResult = new IoTServiceCallResult();
                            if (entity2.Attributes.Contains("msdyn_name"))
                            {
                                ioTServiceCallResult.JobId = entity2.GetAttributeValue<string>("msdyn_name");
                            }

                            if (entity2.Attributes.Contains("msdyn_name"))
                            {
                                ioTServiceCallResult.JobGuid = entity2.Id;
                            }

                            if (entity2.Attributes.Contains("hil_callsubtype"))
                            {
                                ioTServiceCallResult.CallSubType = entity2.GetAttributeValue<EntityReference>("hil_callsubtype").Name;
                            }

                            if (entity2.Attributes.Contains("createdon"))
                            {
                                ioTServiceCallResult.JobLoggedon = entity2.GetAttributeValue<DateTime>("createdon").AddMinutes(330.0).ToString();
                            }

                            if (entity2.Attributes.Contains("msdyn_substatus"))
                            {
                                ioTServiceCallResult.JobStatus = entity2.GetAttributeValue<EntityReference>("msdyn_substatus").Name;
                            }

                            if (entity2.Attributes.Contains("hil_owneraccount"))
                            {
                                ioTServiceCallResult.JobAssignedTo = entity2.GetAttributeValue<EntityReference>("hil_owneraccount").Name;
                            }

                            if (entity2.Attributes.Contains("msdyn_customerasset"))
                            {
                                ioTServiceCallResult.CustomerAsset = entity2.GetAttributeValue<EntityReference>("msdyn_customerasset").Name;
                            }

                            if (entity2.Attributes.Contains("hil_productcategory"))
                            {
                                ioTServiceCallResult.ProductCategory = entity2.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                            }

                            if (entity2.Attributes.Contains("hil_natureofcomplaint"))
                            {
                                ioTServiceCallResult.NatureOfComplaint = entity2.GetAttributeValue<EntityReference>("hil_natureofcomplaint").Name;
                            }

                            if (entity2.Attributes.Contains("hil_jobclosuredon"))
                            {
                                ioTServiceCallResult.JobClosedOn = entity2.GetAttributeValue<DateTime>("hil_jobclosuredon").AddMinutes(330.0).ToString();
                            }

                            if (entity2.Attributes.Contains("hil_customerref"))
                            {
                                ioTServiceCallResult.CustomerName = entity2.GetAttributeValue<EntityReference>("hil_customerref").Name;
                            }

                            if (entity2.Attributes.Contains("hil_fulladdress"))
                            {
                                ioTServiceCallResult.ServiceAddress = entity2.GetAttributeValue<string>("hil_fulladdress");
                            }

                            if (entity2.Attributes.Contains("asset.hil_modelname"))
                            {
                                string modelName = (string)((AliasedValue)entity2["asset.hil_modelname"]).Value; 

                                ioTServiceCallResult.Product = modelName;
                            }

                            if (entity2.Attributes.Contains("hil_customercomplaintdescription"))
                            {
                                ioTServiceCallResult.ChiefComplaint = entity2.GetAttributeValue<string>("hil_customercomplaintdescription");
                            }

                            if (entity2.Attributes.Contains("hil_preferredtime"))
                            {
                                ioTServiceCallResult.PreferredPartOfDay = entity2.GetAttributeValue<OptionSetValue>("hil_preferredtime").Value;
                                ioTServiceCallResult.PreferredPartOfDayName = entity2.FormattedValues["hil_preferredtime"].ToString();
                            }

                            if (entity2.Attributes.Contains("hil_preferreddate"))
                            {
                                ioTServiceCallResult.PreferredDate = entity2.GetAttributeValue<DateTime>("hil_preferreddate").AddMinutes(330.0).ToShortDateString();
                            }

                            ioTServiceCallResult.StatusCode = "200";
                            ioTServiceCallResult.StatusDescription = "OK";
                            list.Add(ioTServiceCallResult);

                            i++;
                        }
                    }
                    watch.Stop();
                    Console.WriteLine("Before Returning: " + watch.ElapsedMilliseconds);

                    i = 0;
                    foreach (IoTServiceCallResult tempitem in list)
                    {
                        PropertyInfo[] proInfo = typeof(IoTServiceCallResult).GetProperties();
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine();
                        Console.WriteLine("Printing Record for " + i);
                        Console.WriteLine("==============================================================");
                        Console.WriteLine();
                        Console.ForegroundColor = ConsoleColor.White;
                        foreach (PropertyInfo pi in proInfo)
                        {
                            Console.WriteLine(pi.Name + "\t \t \t: " + pi.GetValue(tempitem, null));
                        }
                        i++;
                    }
                    return list;
                }

                item = new IoTServiceCallResult
                {
                    StatusCode = "503",
                    StatusDescription = "D365 Service Unavailable"
                };
                list.Add(item);
            }
            catch (Exception ex)
            {
                IoTServiceCallResult item = new IoTServiceCallResult
                {
                    StatusCode = "500",
                    StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper()
                };
                list.Add(item);
            }

            return list;
        }

    }

    public class IoTServiceCallResult
    {
        [DataMember]
        public string JobId { get; set; }

        [DataMember]
        public Guid JobGuid { get; set; }

        [DataMember]
        public string CallSubType { get; set; }

        [DataMember]
        public string JobLoggedon { get; set; }

        [DataMember]
        public string JobStatus { get; set; }

        [DataMember]
        public string JobAssignedTo { get; set; }

        [DataMember]
        public string CustomerAsset { get; set; }

        [DataMember]
        public string ProductCategory { get; set; }

        [DataMember]
        public string NatureOfComplaint { get; set; }

        [DataMember]
        public string JobClosedOn { get; set; }

        [DataMember]
        public string CustomerName { get; set; }

        [DataMember]
        public string ServiceAddress { get; set; }

        [DataMember]
        public string Product { get; set; }

        [DataMember]
        public string ChiefComplaint { get; set; }

        [DataMember]
        public string PreferredDate { get; set; }

        [DataMember]
        public int PreferredPartOfDay { get; set; }

        [DataMember]
        public string PreferredPartOfDayName { get; set; }

        [DataMember]
        public string StatusCode { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
    }
}
