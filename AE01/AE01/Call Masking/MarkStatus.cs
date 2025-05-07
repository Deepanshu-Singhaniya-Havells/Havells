using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace AE01.Call_Masking
{
    internal class MarkStatus
    {
        private IOrganizationService service;
        public MarkStatus(IOrganizationService _service)
        {
            this.service = _service;
        }

        internal void markStatusAsMissed()
        {
            string fetchNullCalls = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='phonecall'>
                                <attribute name='subject'/>
                                <attribute name='statuscode'/>
                                <attribute name='scheduleddurationminutes'/>
                                <attribute name='regardingobjectid'/>
                                <attribute name='ownerid'/>
                                <attribute name='actualdurationminutes'/>
                                <attribute name='directioncode'/>
                                <attribute name='description'/>
                                <attribute name='createdon'/>
                                <attribute name='hil_alternatenumber1'/>
                                <attribute name='hil_contactpreference'/>
                                <attribute name='hil_consumer'/>
                                <attribute name='hil_callingnumber'/>
                                <attribute name='hil_calledtonum'/>
                                <attribute name='to'/>
                                <attribute name='from'/>
                                <attribute name='scheduledend'/>
                                <attribute name='hil_appointmentconfirmed'/>
                                <attribute name='statecode'/>
                                <attribute name='activityid'/>
                                <attribute name='hil_disposition'/>
                                <order attribute='createdon' descending='true'/>
                                <filter type='and'>
                                <condition attribute='subject' operator='like' value='Xchange%'/>
                                <condition attribute='createdon' operator='today'/>
                                <condition attribute='hil_disposition' operator='null'/>
                                </filter>
                                </entity>
                                </fetch>";


            EntityCollection tempCollection = service.RetrieveMultiple(new FetchExpression(fetchNullCalls)); 

            foreach(Entity entity in tempCollection.Entities)
            {
                entity["hil_disposition"] = new OptionSetValue(8); 
                service.Update(entity);
            }

        }
    }
}
