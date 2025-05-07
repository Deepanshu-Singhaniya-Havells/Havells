using System;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace Havells_Plugin.TimeOffRequest
{
   public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == msdyn_timeoffrequest.EntityLogicalName
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    CheckDuplicate(entity, service);
                    SetEndTime(entity, service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.TimeOffRequest.PreCreate.Execute" + ex.Message);
            }

            #endregion
        }
        public static void CheckDuplicate(Entity entity, IOrganizationService service)
        {
            try
            {
                msdyn_timeoffrequest TimeOffReq = entity.ToEntity<msdyn_timeoffrequest>();
                DateTime now = DateTime.Now;
                QueryExpression Query = new QueryExpression(msdyn_timeoffrequest.EntityLogicalName);
                Query.ColumnSet = new ColumnSet(true);
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, TimeOffReq.OwnerId.Id);
                Query.Criteria.AddCondition("createdon", ConditionOperator.Today);
                EntityCollection Found = service.RetrieveMultiple(Query);
                if(Found.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException("-------->>>Time entry already exist for today.<<<--------");
                }
                
                //Query.Criteria.AddCondition("msdyn_endtime", ConditionOperator.Equal, now);
                //String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //  <entity name='msdyn_timeoffrequest'>
                //    <attribute name='createdon' />
                //    <attribute name='msdyn_starttime' />
                //    <attribute name='msdyn_resource' />
                //    <attribute name='msdyn_endtime' />
                //    <attribute name='msdyn_timeoffrequestid' />
                //    <order attribute='createdon' descending='true' />
                //    <filter type='and'>
                //      <condition attribute='msdyn_starttime' operator='on' value='{0}' />
                //      <condition attribute='ownerid' operator='eq' uiname='Abhinav saini' uitype='systemuser' value='{1}' />
                //    </filter>
                //  </entity>
                //</fetch>";
                //DateTime now = DateTime.Now;
                ////2018-07-30
                //String dt = now.Year + "-" + now.Month + "-" + now.Day;
                //EntityCollection en = service.RetrieveMultiple(new FetchExpression(String.Format(fetch, dt, TimeOffReq.OwnerId.Id)));
                //if (en.Entities.Count > 0)
                //    throw new InvalidPluginExecutionException("-------->>>Time entry already exist for today.<<<--------");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        public static void SetEndTime(Entity entity, IOrganizationService service)
        {
            msdyn_timeoffrequest iRequest = entity.ToEntity<msdyn_timeoffrequest>();
            iRequest.msdyn_StartTime = iRequest.CreatedOn;
            TimeSpan iEndTime = GetEndTime(service);
            DateTime EndTime = DateTime.Now.Date + iEndTime;
            iRequest.msdyn_EndTime = EndTime;
        }
        public static TimeSpan GetEndTime(IOrganizationService service)
        {
            DateTime iEndTime = new DateTime();
            TimeSpan iTime = DateTime.Now.TimeOfDay;
            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "End Time"));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                hil_integrationconfiguration iConf = Found.Entities[0].ToEntity<hil_integrationconfiguration>();
                if(iConf.Attributes.Contains("hil_attendanceendtime"))
                {
                    iEndTime = (DateTime)iConf["hil_attendanceendtime"];
                    iTime = iEndTime.TimeOfDay;
                }
            }
            return iTime;
        }
    }
}