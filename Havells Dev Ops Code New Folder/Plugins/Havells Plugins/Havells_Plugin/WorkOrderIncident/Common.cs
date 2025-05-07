using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;


namespace Havells_Plugin.WorkOrderIncident
{
    class Common
    {
        public static readonly string _restrictSpareRoleName = "Restrict Spare Part Prepopulating";
        public static Guid SetModeldetail(Entity entity, IOrganizationService service)
        {
            try
            {
                Guid Model = new Guid();
                msdyn_workorderincident WoIncident = entity.ToEntity<msdyn_workorderincident>();
                if (WoIncident.msdyn_CustomerAsset != null)
                {
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        var obj = (from _Asset in orgContext.CreateQuery<msdyn_customerasset>()
                                   join _model in orgContext.CreateQuery<Product>()
                                   on _Asset.msdyn_Product.Id equals _model.ProductId.Value
                                   where _Asset.msdyn_customerassetId.Value == WoIncident.msdyn_CustomerAsset.Id
                                   select new
                                   {
                                       _Asset.msdyn_Product,
                                       _model
                                   }).Take(1);
                        foreach (var iobj in obj)
                        {
                            if (iobj.msdyn_Product != null)
                            {
                                msdyn_workorderincident UpWOIncident = new msdyn_workorderincident();
                                UpWOIncident.msdyn_workorderincidentId = WoIncident.msdyn_workorderincidentId.Value;
                                UpWOIncident.hil_modelcode = iobj.msdyn_Product;
                                UpWOIncident.hil_modelname = iobj.msdyn_Product.Name;
                                if(iobj._model!=null && iobj._model.Description!=null)
                                UpWOIncident["hil_modeldescription"] = iobj._model.Description;
                                Model = iobj.msdyn_Product.Id;
                                service.Update(UpWOIncident);
                            }
                        }
                    }
                }
                return Model;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.SetModeldetail" + ex.Message);
            }
        }


        #region SetJobClass
        public static void getandSetJobClass(Entity entity, IPluginExecutionContext context, IOrganizationService service)
        {
            try
            {
                //tracingService.Trace("2");
                //tracingService.Trace("2.1" + context.MessageName.ToUpper());
                if (context.MessageName.ToUpper() == "UPDATE")
                {
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    //tracingService.Trace("3");
                    if (entity.Contains("hil_quantity") || entity.Contains("msdyn_customerasset"))
                    {
                       //tracingService.Trace("4");
                        entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                        getJobClass(entity, service);
                    }
                }
                else if (context.MessageName.ToUpper() == "CREATE")
                {
                    if (entity.Contains("hil_quantity") || entity.Contains("msdyn_customerasset"))
                    {
                        //tracingService.Trace("5");
                        getJobClass(entity, service);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.Common.getandSetJobClass" + ex.Message);
            }
        }

        public static void getJobClass(Entity entity, IOrganizationService service)
        {
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                   
                    msdyn_workorderincident woIncident = entity.ToEntity<msdyn_workorderincident>();
                    //tracingService.Trace("6 is warranty null" + woIncident.hil_warrantystatus);
                    if (woIncident.msdyn_CustomerAsset != null)
                    { 
                        if (woIncident.hil_warrantystatus != null  && woIncident.hil_Quantity!=null)
                    {
                        //tracingService.Trace("7 WarrantyStatus"+ woIncident.hil_warrantystatus.Value);
                            if (
                                woIncident.msdyn_WorkOrder != null && woIncident.msdyn_CustomerAsset != null
                                && woIncident.hil_warrantystatus.Value == ((int)hil_inventorytype.InWarranty)
                                )
                            {
                                Int32 iJobClassValue = -1;
                                #region GetJobClasstoSet
                                var obj = from _woIncident in orgContext.CreateQuery<msdyn_workorderincident>()
                                          join _Asset in orgContext.CreateQuery<msdyn_customerasset>()
                                          on _woIncident.msdyn_CustomerAsset.Id equals _Asset.msdyn_customerassetId.Value
                                          join _ProductCategory in orgContext.CreateQuery<Product>()
                                          on _Asset.hil_ProductCategory.Id equals _ProductCategory.ProductId.Value
                                          where _woIncident.msdyn_workorderincidentId.Value == woIncident.Id
                                          select new
                                          {
                                              _ProductCategory.hil_MaximumThreshold,
                                              _ProductCategory.hil_MinimumThreshold,
                                              _woIncident.hil_Quantity
                                          };

                                foreach (var iobj in obj)
                                {
                                    //tracingService.Trace("8 qty "+ iobj.hil_Quantity +" Min Thershold"+ iobj.hil_MinimumThreshold + " Max Thershold" + iobj.hil_MaximumThreshold);
                                    if (iobj.hil_Quantity != null && iobj.hil_MinimumThreshold != null && iobj.hil_MaximumThreshold != null)
                                    {
                                        //tracingService.Trace("8" + iobj.hil_Quantity.Value);
                                        if (iobj.hil_Quantity.Value <= iobj.hil_MinimumThreshold.Value)
                                        {
                                            iJobClassValue = 1;
                                        }
                                        else if (iobj.hil_MinimumThreshold.Value < iobj.hil_Quantity.Value && iobj.hil_MaximumThreshold.Value >= iobj.hil_Quantity.Value)
                                        {
                                            iJobClassValue = 2;
                                        }
                                        else if (iobj.hil_Quantity.Value > iobj.hil_MaximumThreshold.Value)
                                        {
                                            iJobClassValue = 3;
                                        }
                                    }
                                }
                                #endregion
                                if (iJobClassValue != -1)
                                {
                                    setJobClass(woIncident.msdyn_WorkOrder.Id, iJobClassValue, service);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.Common.getJobClass" + ex.Message);
            }
        }

        public static void setJobClass(Guid fsWorkOrderId, Int32 iJobClassValue, IOrganizationService service)
        {
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    Int32 JobClass = -1;
                    #region CheckWorkOrderJobClass
                    var obj = from _WorkOrder in orgContext.CreateQuery<msdyn_workorder>()
                              where _WorkOrder.msdyn_workorderId.Value == fsWorkOrderId
                              select new { _WorkOrder.hil_JobClass };
                    foreach (var iobj in obj)
                    {
                        if (iobj.hil_JobClass != null)
                        { JobClass = iobj.hil_JobClass.Value; }
                    }
                    #endregion

                    if (JobClass != -1) //if job class already exist on wo
                    {
                        if (iJobClassValue > JobClass)//check if new job class is greater than existing
                        {
                            msdyn_workorder upWo = new msdyn_workorder();
                            upWo.msdyn_workorderId = fsWorkOrderId;
                            upWo.hil_JobClass = new OptionSetValue(iJobClassValue);
                            service.Update(upWo);
                        }
                    }
                    else //If Job class is blank on workorder 
                    {
                        msdyn_workorder upWo = new msdyn_workorder();
                        upWo.msdyn_workorderId = fsWorkOrderId;
                        upWo.hil_JobClass = new OptionSetValue(iJobClassValue);
                        service.Update(upWo);
                    }

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.Common.setJobClass" + ex.Message);
            }
        }

        #endregion
    }
}
