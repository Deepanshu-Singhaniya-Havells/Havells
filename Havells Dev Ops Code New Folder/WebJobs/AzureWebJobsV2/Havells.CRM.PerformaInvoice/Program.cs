using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Havells.CRM.PerformaInvoice
{
    class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        static void Main(string[] args)
        {
            try
            {


                ClaimPerformaNew obj = new ClaimPerformaNew();
                IOrganizationService _service = obj.ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));

                //string fet = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //                  <entity name='hil_claimheader'>
                //                    <attribute name='hil_claimheaderid' />
                //                    <attribute name='hil_name' />
                //                    <attribute name='createdon' />
                //                    <order attribute='hil_name' descending='false' />
                //                    <filter type='and'>
                //                      <condition attribute='hil_isperformainvoicecreated' operator='eq' value='1' />
                //                    </filter>
                //                  </entity>
                //                </fetch>";

                //EntityCollection entityCollection1 = _service.RetrieveMultiple(new FetchExpression(fet));
                //int i = 0;
                //foreach(Entity entity in entityCollection1.Entities)
                //{
                //    i++;
                //    Console.WriteLine(i);
                //    Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                //    entity1["hil_isperformainvoicecreated"] = false;
                //    _service.Update(entity1);
                    
                //}
                 
                EntityCollection ClaimInvoiceCollection = new EntityCollection();
                EntityCollection entityCollection = new EntityCollection();
                int pageNo = 1;

            hil_claimheader:


                String fetchXML = @"<fetch version='1.0' output-format='xml-platform'  page='" + pageNo + @"' distinct='true' mapping='logical'>
                          <entity name='hil_claimheader'>
                            <attribute name='hil_claimheaderid' />
                            <attribute name='hil_name' />
                            <attribute name='hil_franchiseeinvoiceno' />
                            <attribute name='hil_franchisee' />
                            <attribute name='hil_fiscalmonth' />
                            <attribute name='hil_sapponumber'/>
                            <attribute name='hil_syncdoneon'/>
                            <attribute name='modifiedon'/>
                            <attribute name='hil_approvedon'/>
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_performastatus' operator='eq' value='4' />
                              <condition attribute='hil_isperformainvoicecreated' operator='ne' value='1' />
                            </filter>
                            <link-entity name='account' from='accountid' to='hil_franchisee' visible='false' link-type='outer' alias='ab'>
                              <attribute name='hil_vendorcode' />
                              <attribute name='hil_state' />
                              <attribute name='hil_salesoffice' />
                              <attribute name='hil_district' />
                              <attribute name='hil_city' />
                              <attribute name='hil_category' />
                              <attribute name='hil_branch' />
                              <attribute name='hil_area' />
                              <attribute name='accountnumber' />
                            </link-entity>
                            <link-entity name='hil_claimperiod' from='hil_claimperiodid' to='hil_fiscalmonth' visible='false' link-type='outer' alias='ac'>
                                  <attribute name='hil_todate' />
                                  <attribute name='hil_fromdate' />
                            </link-entity>
                          </entity>
                        </fetch>";

                entityCollection = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                ClaimInvoiceCollection.Entities.AddRange(entityCollection.Entities.ToArray());
                if (entityCollection.Entities.Count == 5000)
                {
                    pageNo++;
                    goto hil_claimheader;
                }
                obj.GenerateAnnexureUrl(_service, ClaimInvoiceCollection);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error is " + ex.Message);
            }
        }
    }
}
