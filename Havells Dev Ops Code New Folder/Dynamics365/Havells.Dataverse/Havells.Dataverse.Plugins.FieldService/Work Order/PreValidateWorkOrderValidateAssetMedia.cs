using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Havells.Dataverse.Plugins.FieldService.Work_Order
{
    public class PreValidateWorkOrderValidateAssetMedia : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("msdyn_substatus"))
                        ProcessRequest(entity, _tracingService, service);
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
        public void ProcessRequest(Entity entity, ITracingService _tracingService,IOrganizationService service)
        {
            string status_Closed = "1727fa6c-fa0f-e911-a94e-000d3af060a1"; 
            string status_EstimateApproved = "5290be10-6047-eb11-bb23-000d3af053ad"; 
            string status_EstimateGiven = "5090be10-6047-eb11-bb23-000d3af053ad"; 
            string status_PendingforGasCharge = "5490be10-6047-eb11-bb23-000d3af053ad"; 
            string status_PendingorProductDelivery = "8cb8d23b-35df-ed11-8847-6045bdac5897"; 
            EntityReference _jobSubStatus = entity.GetAttributeValue<EntityReference>("msdyn_substatus");
            Entity _entJob = null;

            if (_jobSubStatus.Id.Equals(new Guid(status_Closed)))
            {
                string fetchPermittedCount = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_workorderid' />
                <attribute name='msdyn_customerasset' />
                <attribute name='hil_isocr' />
                <order attribute='msdyn_name' descending='false' />
                <filter type='and'>
                    <condition attribute='msdyn_workorderid' operator='eq' value='{entity.Id}' />
                </filter>
                <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' link-type='inner' alias='am'>
                    <filter type='and'>
                        <condition attribute='createdon' operator='on-or-after' value='2024-10-17' />
                        <condition attribute='hil_source' operator='eq' value='4' />
                    </filter>
                    <link-entity name='product' from='productid' to='hil_productsubcategory' link-type='inner' alias='an'>
                        <attribute name='hil_assetmediacount' />
                        <filter type='and'>
                            <condition attribute='hil_assetmediacount' operator='gt' value='0' />
                        </filter>
                    </link-entity>
                </link-entity>
                </entity>
                </fetch>";
                EntityCollection permittedCountColl = service.RetrieveMultiple(new FetchExpression(fetchPermittedCount));
                if (permittedCountColl.Entities.Count > 0)
                {
                    bool _isOCR = permittedCountColl.Entities[0].Contains("hil_isocr") ? permittedCountColl.Entities[0].GetAttributeValue<bool>("hil_isocr") : false;
                    if (!_isOCR)
                    {
                        int permittedCount = (int)permittedCountColl.Entities[0].GetAttributeValue<AliasedValue>("an.hil_assetmediacount").Value;
                        EntityReference _entCustomerAsser = permittedCountColl.Entities[0].GetAttributeValue<EntityReference>("msdyn_customerasset");
                        string fetchactualCount = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='annotation'>
                        <attribute name='annotationid' />
                        <order attribute='subject' descending='false' />
                        <filter type='and'>
                        <filter type='or'>
                            <condition attribute='notetext' operator='like' value='%https://d365storagesa.blob.core.windows.net%'/>
                            <condition attribute='isdocument' operator='eq' value='1' />
                        </filter>
                        </filter>
                        <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='objectid' link-type='inner' alias='as'>
                            <filter type='and'>
                                <condition attribute='msdyn_customerassetid' operator='eq' value='{_entCustomerAsser.Id}' />
                            </filter>
                        </link-entity>
                        </entity>
                        </fetch>";

                        int actualCount = service.RetrieveMultiple(new FetchExpression(fetchactualCount)).Entities.Count;

                        if (actualCount < permittedCount)
                        {
                            throw new Exception($"Please ensure to upload the relevant documents (Invoice Copy, Unit Serial Number, Product Photograph) on customer asset ({_entCustomerAsser.Name}) before marking the Job as 'Closed'");
                        }
                    }
                }
            }
            else if (_jobSubStatus.Id.Equals(new Guid(status_EstimateApproved)) || _jobSubStatus.Id.Equals(new Guid(status_EstimateGiven)))
            {
                _entJob = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("msdyn_customerasset"));
                if (!_entJob.Contains("msdyn_customerasset"))
                {
                    throw new Exception($"Please create Incident & Asset before selecting this Job Sub Status.");
                }
            }
            else if (_jobSubStatus.Id.Equals(new Guid(status_PendingforGasCharge)))
            {
                _entJob = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_productcategory"));
                Guid _prodCatgAirConditioner = new Guid("D51EDD9D-16FA-E811-A94C-000D3AF0694E");
                Guid _prodCatgRefrigerator = new Guid("2DD99DA1-16FA-E811-A94C-000D3AF06091");

                if (_entJob.Contains("hil_productcategory"))
                {
                    EntityReference _entRefProdCatg = _entJob.GetAttributeValue<EntityReference>("hil_productcategory");
                    if(_entRefProdCatg.Id != _prodCatgAirConditioner && _entRefProdCatg.Id != _prodCatgRefrigerator)
                    {
                        throw new Exception($"This is applicable for Air Conditioner & Refrigerator only.");
                    }
                }
            }
            else if (_jobSubStatus.Id.Equals(new Guid(status_PendingorProductDelivery)))
            {
                _entJob = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_callsubtype"));
                Guid _callSubTypeInstallation = new Guid("e3129d79-3c0b-e911-a94e-000d3af06cd4");

                if (_entJob.Contains("hil_callsubtype"))
                {
                    EntityReference _entRefCallSubtype = _entJob.GetAttributeValue<EntityReference>("hil_callsubtype");
                    if (_entRefCallSubtype.Id != _callSubTypeInstallation)
                    {
                        throw new Exception($"This is applicable for Installation Call Sub Type only.");
                    }
                }

            }
        }
    }
}
