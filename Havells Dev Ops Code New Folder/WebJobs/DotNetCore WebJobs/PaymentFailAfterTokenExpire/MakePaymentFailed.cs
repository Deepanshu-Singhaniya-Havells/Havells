using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace PaymentFailAfterTokenExpire
{
    public class MakePaymentFailed
    {
        private readonly ServiceClient _service;
        public MakePaymentFailed(ServiceClient service)
        {
            _service = service;
        }
        public void UpdateStatusAfterTokenExpire()
        {
            string fetchPaymentReceipts = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='hil_paymentreceipt'>
                                        <attribute name='hil_paymentreceiptid' />
                                        <attribute name='hil_transactionid' />
                                        <attribute name='createdon' />
                                        <attribute name='hil_tokenexpireson' />
                                        <attribute name='overriddencreatedon' />
                                        <order attribute='hil_transactionid' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='hil_paymentstatus' operator='eq' value='1' />
                                          <condition attribute='createdon' operator='last-x-days' value='1' />
                                        </filter>
                                        <link-entity name='salesorder' from='salesorderid' to='hil_orderid' link-type='inner' alias='ac'>
                                          <filter type='and'>
                                            <condition attribute='hil_ordertype' operator='eq' value='{1F9E3353-0769-EF11-A670-0022486E4ABB}' />
                                            <condition attribute='hil_sellingsource' operator='not-in'>
                                              <value>{03B5A2D6-CC64-ED11-9562-6045BDAC526A}</value>
                                              <value>{608E899B-A8A3-ED11-AAD1-6045BDAD27A7}</value>
                                              <value>{668E899B-A8A3-ED11-AAD1-6045BDAD27A7}</value>
                                            </condition>
                                          </filter>
                                        </link-entity>
                                        </entity>
                                    </fetch>";

            EntityCollection paymentReceipts = _service.RetrieveMultiple(new FetchExpression(fetchPaymentReceipts));

            for (int i = 0; i < paymentReceipts.Entities.Count; i++)
            {
                DateTime tokenExpire = paymentReceipts.Entities[i].GetAttributeValue<DateTime>("hil_tokenexpireson").AddMinutes(330);
                if (DateTime.Now > tokenExpire)
                {
                    Entity updatePymentReceipt = new Entity(paymentReceipts.EntityName, paymentReceipts.Entities[i].Id);

                    updatePymentReceipt["hil_paymentstatus"] = new OptionSetValue(2); // Failed
                    _service.Update(updatePymentReceipt);
                }
            }
        }
    }
}
