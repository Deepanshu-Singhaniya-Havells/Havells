using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
namespace HavellsNewPlugin.Case
{
    public class CaseAssignmentMatrixLinePreCreate : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                  && context.PrimaryEntityName.ToLower() == "hil_caseassignmentmatrixline" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    EntityReference _entRef = entity.GetAttributeValue<EntityReference>("hil_caseassignmentmatrixid");

                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_caseassignmentmatrixline'>
                    <attribute name='hil_caseassignmentmatrixlineid' />
                    <filter type='and'>
                        <condition attribute='hil_caseassignmentmatrixid' operator='eq' value='{_entRef.Id}' />
                    </filter>
                    </entity>
                    </fetch>";

                    Entity _entCaseAssignmentMatrix = service.Retrieve("hil_caseassignmentmatrix", _entRef.Id, new ColumnSet("hil_name"));
                    EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    int _rowCount = _entCol.Entities.Count + 1;
                    entity["hil_name"] = _entCaseAssignmentMatrix.GetAttributeValue<string>("hil_name") + "-" + _rowCount.ToString().PadLeft(2, '0');
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Case Pre Create Error : " + ex.Message);
            }

        }
    }
}
