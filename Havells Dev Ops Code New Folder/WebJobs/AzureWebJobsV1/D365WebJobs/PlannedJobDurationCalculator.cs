using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace D365WebJobs
{
    public class PlannedJobDurationCalculator
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _tracingService;
        private readonly Guid _jobId;

        public PlannedJobDurationCalculator(IOrganizationService service, Guid jobId)
        {
            _service = service;
            _jobId = jobId;
        }

        public void UpdateJobDuration()
        {
            try
            {
                var duration = GetTotalJobDurationFromJobActivity();

                if (duration != null)
                {
                    Entity jobToUpdate = new Entity("ogre_workorder", _jobId);
                    jobToUpdate["ogre_jobduration"] = duration;
                    _service.Update(jobToUpdate);
                }
                else
                {
                    //_tracingService.Trace($"Not able to calculate duration for ogre_workorder {_jobId}");
                }
            }
            catch (Exception ex)
            {
                //_tracingService.Trace(ex.ToString());
            }
        }

        private decimal? GetTotalJobDurationFromJobActivity()
        {
            Log();

            decimal? duration = null;


            string fetchXml = $@"<fetch distinct='false' mapping='logical' aggregate='true'> 
                                     <entity name='ogre_workorderactivity'> 
                                        <attribute name='ogre_duration' alias='duration' aggregate='sum'/> 
                                        <filter type='and'>
                                            <condition attribute='statecode' operator='eq' value='0' />
                                            <condition attribute='ogre_workorderid' operator='eq' value='{_jobId}' />
                                        </filter>
                                     </entity> 
                                  </fetch>";

            //_tracingService.Trace($"execute fetchxml");

            EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(fetchXml));

            if (entCol.Entities.Count > 0)
            {
                duration = Convert.ToDecimal(entCol.Entities[0].GetAttributeValue<AliasedValue>("duration").Value);

                //_tracingService.Trace($"duration is {duration}");
            }

            return duration;
        }

        private void Log()
        {
            try
            {
                string fetchXml = $@"<fetch> 
                                     <entity name='ogre_workorderactivity'>
                                        <attribute name='ogre_name'/> 
                                        <filter type='and'>
                                            <condition attribute='statecode' operator='eq' value='0' />
                                            <condition attribute='ogre_workorderid' operator='eq' value='{_jobId}' />
                                        </filter>
                                     </entity> 
                                  </fetch>";

                var entCol = _service.RetrieveMultiple(new FetchExpression(fetchXml));

                if (entCol != null && entCol.Entities != null)
                    //_tracingService.Trace($"Total {entCol.Entities.Count} record found:");

                foreach (var item in entCol.Entities)
                {
                    var name = item.GetAttributeValue<string>("ogre_name");
                    //_tracingService.Trace($"{name}");
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
