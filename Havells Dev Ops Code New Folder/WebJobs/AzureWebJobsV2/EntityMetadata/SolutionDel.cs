using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EntityMetadata
{
    public class SolutionDel
    {
        public static Dictionary<Guid, String> contactList = new Dictionary<Guid, string>();
        public static void _main(IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("solution");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition("publisherid", ConditionOperator.Equal, new Guid("D21AAB71-79E7-11DD-8874-00188B01E34F"));
            EntityCollection entityColl = service.RetrieveMultiple(query);

            //service.Delete(entityColl[0].LogicalName, entityColl[0].Id);

            foreach (Entity entity in entityColl.Entities)
            {
                Console.WriteLine(entity.GetAttributeValue<string>("friendlyname"));
                service.Delete(entity.LogicalName, entity.Id);
            }
            Console.WriteLine();
        }

        public static void GetCustomerDate(IOrganizationService service)
        {
            //string cityID = "69BC99AD-C9F7-E811-A94C-000D3AF0694E";//KOLKATA-WB

            //string cityID = "a1f3e606-c2f7-e811-a94c-000d3af0677f"; //DELHI CENTRAL-DL
            //string cityID = "d1c7b367-cdf7-e811-a94c-000d3af0694e"; // Mumbai
            string cityID = "89470a3b-bdf7-e811-a94c-000d3af06a98"; //Bangalore
            cityID = "8cdb4689-d6f7-e811-a94c-000d3af06c56"; //Vizag
            cityID = "0f87c3a3-d4f7-e811-a94c-000d3af06a98"; //Surat
            cityID = "6f864819-cbf7-e811-a94c-000d3af0694e"; //Lucknow
           cityID = "8a2f3ee7-cff7-e811-a94c-000d3af0694e"; //Patna
            Console.WriteLine("Customer" + "|" + "City" + "|" + "State" + "|" + "Full Name" + "|" + "Registered On " + "|" + "Mobile Number");
            int monthCount = 0;
            DateTime date = new DateTime(2022, 01, 01);
            while (contactList.Count < 100)
            {
                if (contactList.Count < 100)
                {
                    
                    if (monthCount > 10)
                    {
                        int day = date.Day - 1;
                        monthCount = 0;
                        date = date.AddMonths(1).AddDays(-day);
                    }
                    string _onOrAfterDate = date.Date.ToString("yyyy-MM-dd");//YYYY-MM-DD

                    string fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true' top='1'  >
                              <entity name='hil_address'>
                                <attribute name='hil_addressid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <attribute name='hil_state' />
                                <attribute name='hil_city' />
                                <attribute name='hil_customer' />
                                <order attribute='createdon' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_state' operator='eq' uiname='DELHI' uitype='hil_state' value='{{27FCB4DF-BBF7-E811-A94C-000D3AF06091}}' />
                                </filter>
                                <link-entity name='contact' from='contactid' to='hil_customer' link-type='inner' alias='ac'>
                                  <attribute name='createdon' />
                                  <attribute name='fullname' />
                                  <attribute name='mobilephone' />
                                  <filter type='and'>
                                    <condition attribute='hil_consumersource' operator='eq' value='5' />
                                    <condition attribute='createdon' operator='on' value='{_onOrAfterDate}' />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";
                    EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(fetch));


                    foreach (Entity entity in entColl.Entities)
                    {
                        //if (contactList.Count < 10)
                        {
                            var dictval = from x in contactList
                                          where x.Key.Equals(entity.GetAttributeValue<EntityReference>("hil_customer").Id)
                                          select x;
                            if (dictval.ToList().Count == 0)
                            {
                                contactList.Add(entity.GetAttributeValue<EntityReference>("hil_customer").Id, (string)entity.GetAttributeValue<AliasedValue>("ac.fullname").Value);

                                Console.WriteLine(entity.GetAttributeValue<EntityReference>("hil_customer").Id + "|" +
                                    entity.GetAttributeValue<EntityReference>("hil_city").Name + "|" +
                                    entity.GetAttributeValue<EntityReference>("hil_state").Name + "|" +
                                    entity.GetAttributeValue<AliasedValue>("ac.fullname").Value + "|" +
                                    entity.GetAttributeValue<AliasedValue>("ac.createdon").Value + "|" +
                                    entity.GetAttributeValue<AliasedValue>("ac.mobilephone").Value);
                                monthCount++;
                            }
                        }
                    }
                    //Console.WriteLine(entColl.Entities.Count);

                    date = date.AddDays(1);
                }

            }
            Console.WriteLine("");
        }
    }
}
