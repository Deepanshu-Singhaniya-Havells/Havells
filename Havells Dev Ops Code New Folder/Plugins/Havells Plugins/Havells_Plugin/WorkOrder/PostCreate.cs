using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Havells_Plugin.WorkOrder;

namespace Plugins.ServiceTicket
{
    public class PostCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity serviceTicket = (Entity)context.InputParameters["Target"];
                    msdyn_workorder Stck = serviceTicket.ToEntity<msdyn_workorder>();
                    PopulateBrand(serviceTicket, service);
                    SetKKGCode(serviceTicket, service);

                    #region Not in use
                    //CheckParentJobTagged(service, serviceTicket);
                    //SetWarrantyStatus(service, Stck);

                    //Entity jobExtension = new Entity("hil_jobsextension");
                    //jobExtension["hil_jobs"] = serviceTicket.ToEntityReference();
                    //jobExtension["hil_name"] = serviceTicket.GetAttributeValue<string>("msdyn_name");
                    //service.Create(jobExtension);
                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Plugin.ServiceTicket.ServiceTicket_PostCreate.Execute: " + ex.Message);
            }
        }
        public static void PopulateBrand(Entity entity, IOrganizationService service)
        {
            msdyn_workorder Jobs = entity.ToEntity<msdyn_workorder>();
            msdyn_workorder Jobs1 = new msdyn_workorder();
            Jobs1.Id = Jobs.Id;// (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Jobs.Id, new ColumnSet(false));
            if (Jobs.hil_Productcategory != null)
            {
                //EntityReference PdtCgry = Jobs.hil_Productcategory;
                Product Pdt = new Product();
                Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Jobs.hil_Productcategory.Id, new ColumnSet("hil_brandidentifier"));
                if (Pdt.hil_BrandIdentifier != null)
                {
                    Jobs1.hil_Brand = (OptionSetValue)Pdt.hil_BrandIdentifier;
                    service.Update(Jobs1);
                }
            }
        }
        public static void SetKKGCode(Entity entity, IOrganizationService service)
        {
            msdyn_workorder _Wo = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet("hil_productsubcategory", "hil_kkgcode", "hil_customerref", "msdyn_name"));
            if(_Wo.hil_ProductSubcategory != null)
            {
                EntityReference PdtSCg = _Wo.hil_ProductSubcategory;
                Product SubCat = (Product)service.Retrieve(Product.EntityLogicalName, PdtSCg.Id, new ColumnSet("hil_kkgcoderequired"));
                if(SubCat.hil_kkgcoderequired != null)
                {
                    int _OpVal = SubCat.hil_kkgcoderequired.Value;
                    if(_OpVal == 1)
                    {
                        string OTP = string.Empty;
                        OTP = GenerateOTP(service);
                        if(OTP != string.Empty)
                        {
                            //if(_Wo.hil_CustomerRef != null)
                            //{
                            //    EntityReference Cont = _Wo.hil_CustomerRef;
                            //    if(Cont.LogicalName == Contact.EntityLogicalName)
                            //    {
                            //        Contact _Cnt = (Contact)service.Retrieve(Contact.EntityLogicalName, Cont.Id, new ColumnSet("mobilephone", "fullname"));
                            //        if(_Cnt.MobilePhone != null)
                            //        {
                            //            string Mob = _Cnt.MobilePhone;
                            //            string Message = "Dear" + _Cnt.FullName + "!Your KKG Code for Service Request ID "
                            //                            + _Wo.msdyn_name + " is " + OTP + " Share KKG Code to Technician only after Satisfactory " +
                            //                            "Completion of your Service Request";


                            //            string Subject = "HV-OTP";
                            //            Helper.SendSMS(service, Mob, Message, Subject, _Cnt.Id);
                            //        }
                            //    }
                            //}
                            //msdyn_workorder _WoUpdate = new msdyn_workorder();
                            //_WoUpdate.Id = _Wo.Id;
                            // ---------------------START -------------------------------
                            // Added by Kuldeep Khare to Encrypt KKG Code 25/Dec/2019
                            //_WoUpdate.hil_KKGOTP = OTP;

                            //_WoUpdate.hil_KKGOTP = Common.GetEncryptedValue(OTP, service);
                            //_WoUpdate.hil_KKGOTP = Common.Base64Encode(OTP);
                            // -------------------- END ---------------------------------
                            //service.Update(_WoUpdate);
                            //Entity OTPEntity = new Entity("hil_yammersettings");
                            //OTPEntity["hil_name"] = _Wo.hil_CustomerRef.Name;
                            //OTPEntity["hil_settings"] = OTP;
                            //OTPEntity["hil_wo"] = new EntityReference(msdyn_workorder.EntityLogicalName, _Wo.Id);
                            //OTPEntity["hil_jobid"] = _Wo.msdyn_name;
                            //service.Create(OTPEntity);
                        }
                    }
                }
            }
        }
        public static string GenerateOTP(IOrganizationService service)
        {
            string OTPAlpha = string.Empty;
            string OTPNum = string.Empty;
            string OTP = string.Empty;
            int iOTPLengthNum = 4;
            //int iOTPLengthChar = 2;
            //string[] saAllowedCharactersAlpha = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N",
                                                 //"O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
            string[] saAllowedCharactersNum = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            //OTPAlpha = Havells_Plugin.Consumer_App_Bridge.PreCreate.GenerateRandomOTP(iOTPLengthChar, saAllowedCharactersAlpha);
            OTPNum = Havells_Plugin.Consumer_App_Bridge.PreCreate.GenerateRandomOTP(iOTPLengthNum, saAllowedCharactersNum);
            OTP = OTPAlpha + OTPNum;
            return OTP;
        }
        public static void CheckParentJobTagged(IOrganizationService service, Entity Jobs)
        {
            msdyn_workorder ThisJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Jobs.Id, new ColumnSet("msdyn_parentworkorder", "msdyn_customerasset", "hil_callsubtype"));
            if(ThisJob.msdyn_ParentWorkOrder != null)
            {
                DateTime Current = DateTime.Now;
                Current = Current.AddDays(-90);
                EntityReference ParentWo = ThisJob.msdyn_ParentWorkOrder;
                msdyn_workorder ParJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, ParentWo.Id, new ColumnSet("statuscode", "msdyn_customerasset", "hil_callsubtype", "createdon"));
                if((ParJob.statuscode.Value != 2) && (ParJob.CreatedOn < Current) && (ParJob.msdyn_CustomerAsset != ThisJob.msdyn_CustomerAsset) && (ParJob.hil_CallSubType != ThisJob.hil_CallSubType))
                {
                    throw new InvalidPluginExecutionException("Tagged Parent Job Invalid");
                }
            }
        }
        public static void SetWarrantyStatus(IOrganizationService service, msdyn_workorder Wo1)
        {
            if(Wo1.hil_PurchaseDate != null && Wo1.hil_ProductSubcategory != null)
            {
                msdyn_workorder Wo = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Wo1.Id, new ColumnSet(false));
                DateTime date1 = (DateTime)Wo1.hil_PurchaseDate;
                DateTime date2 = DateTime.Now;
                int Diff = ((date2.Year - date1.Year) * 12) + (date2.Month - date1.Month) * 30 + date2.Day - date1.Day;
                QueryByAttribute Qry = new QueryByAttribute(hil_warrantytemplate.EntityLogicalName);
                Qry.ColumnSet = new ColumnSet(true);
                Qry.AddAttributeValue("hil_product", Wo1.hil_ProductSubcategory.Id);
                EntityCollection Found = service.RetrieveMultiple(Qry);
                if(Found.Entities.Count == 1)
                {
                    foreach(hil_warrantytemplate Wtmp in Found.Entities)
                    {
                        int Period = (int)(Wtmp.hil_WarrantyPeriod * 30);
                        if (Period >= Diff)
                        {
                            Wo.hil_WarrantyStatus = new OptionSetValue(1);
                        }
                        else
                        {
                            Wo.hil_WarrantyStatus = new OptionSetValue(2);
                        }
                    }
                }
                else if(Found.Entities.Count >= 1)
                {
                    Wo.hil_WarrantyStatus = new OptionSetValue(3);
                }
                else
                {
                    Wo.hil_WarrantyStatus = new OptionSetValue(3);
                }
                service.Update(Wo);
            }
        }

        #region KKG Code Hashing
        

        #endregion
    }
}