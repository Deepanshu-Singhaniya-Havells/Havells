using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel;

namespace HavellsNewPlugin.Case
{
    public class CaseStateChange : IPlugin
    {
        private static readonly Guid SamparkDepartmentId = new Guid("7bf1705a-3764-ee11-8df0-6045bdaa91c3");
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
                if (context.InputParameters.Contains("EntityMoniker") &&
                    context.InputParameters["EntityMoniker"] is EntityReference)
                {
                    var entityRef = (EntityReference)context.InputParameters["EntityMoniker"];
                    var state = (OptionSetValue)context.InputParameters["State"];
                    var status = (OptionSetValue)context.InputParameters["Status"];

                    Entity Case = service.Retrieve("incident", entityRef.Id, new ColumnSet("hil_casecategory", "hil_casedepartment", "hil_assignmentmatrix", "ownerid"));
                    Guid DepartmentId = Case.GetAttributeValue<EntityReference>("hil_casedepartment").Id;
                    if (DepartmentId != SamparkDepartmentId)
                    {
                        string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='hil_grievancehandlingactivity'>
                                            <attribute name='activityid' />
                                            <attribute name='subject' />
                                            <attribute name='createdon' />
                                            <attribute name=""hil_caseassignmentmatrixlineid"" />
                                            <order attribute='subject' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='regardingobjectid' operator='eq' value='{entityRef.Id}' />
                                              <condition attribute=""statecode"" operator=""in"">
                                                <value>0</value>
                                                <value>2</value>
                                              </condition>
                                            </filter>
                                          </entity>
                                        </fetch>";
                        EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entityCollection.Entities.Count > 0)
                        {
                            throw new InvalidPluginExecutionException("Please complete all the grievance handling acitvities before resolving the case.");
                        }

                        Guid AssignmentMatrixId = Case.GetAttributeValue<EntityReference>("hil_assignmentmatrix").Id;
                        Entity AssignmentMatrix = service.Retrieve("hil_caseassignmentmatrix", AssignmentMatrixId, new ColumnSet("hil_resolveby", "hil_spoc"));
                        Guid CaseOwnerID = Case.GetAttributeValue<EntityReference>("ownerid").Id;
                        int ResolveBy = AssignmentMatrix.GetAttributeValue<OptionSetValue>("hil_resolveby").Value;

                        if (ResolveBy == 1) // SPOC
                        {
                            Guid SPOCID = AssignmentMatrix.GetAttributeValue<EntityReference>("hil_spoc").Id;
                            if (context.InitiatingUserId != SPOCID)
                            {
                                throw new InvalidPluginExecutionException("Access Denied! You are not Authorized to resolve the case");
                            }
                        }
                        else if (ResolveBy == 2) // Assignee
                        {
                            if (context.InitiatingUserId != CaseOwnerID)
                            {
                                throw new InvalidPluginExecutionException("Access Denied! You are not Authorized to resolve the case");
                            }
                        }
                    }
                }
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("An FaultException occurred in the SetStatePlugin plug-in odf Enginner Shift.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex);
            }
        }
    }
}
