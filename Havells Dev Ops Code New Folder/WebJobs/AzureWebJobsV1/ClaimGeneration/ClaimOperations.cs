using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using System.Text;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using System.Net;
using System.IO;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;

namespace ClaimGeneration
{
    public class ClaimOperations
    {
        const string _fixedCompensationClaimCategory = "824A1BCC-6EE5-EA11-A817-000D3AF0501C";

        public void GenerateFixedCompensationLines(IOrganizationService _service, Guid _performaInvoiceId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;

            LinkEntity lnkEntCH = new LinkEntity
            {
                LinkFromEntityName = "hil_claimline",
                LinkToEntityName = "hil_claimheader",
                LinkFromAttributeName = "hil_claimheader",
                LinkToAttributeName = "hil_claimheaderid",
                Columns = new ColumnSet("hil_performastatus"),
                EntityAlias = "ch",
                JoinOperator = JoinOperator.Inner
            };
            LinkEntity lnkEntSO = new LinkEntity
            {
                LinkFromEntityName = "account",
                LinkToEntityName = "hil_salesoffice",
                LinkFromAttributeName = "hil_salesoffice",
                LinkToAttributeName = "hil_salesofficeid",
                Columns = new ColumnSet("hil_state"),
                EntityAlias = "so",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCP = new LinkEntity
            {
                LinkFromEntityName = "hil_claimheader",
                LinkToEntityName = "account",
                LinkFromAttributeName = "hil_franchisee",
                LinkToAttributeName = "accountid",
                Columns = new ColumnSet("hil_state", "hil_salesoffice", "ownerid"),
                EntityAlias = "cp",
                JoinOperator = JoinOperator.Inner
            };
            lnkEntCP.LinkEntities.Add(lnkEntSO);
            lnkEntCH.LinkEntities.Add(lnkEntCP);
            queryExp = new QueryExpression("hil_claimline");
            queryExp.ColumnSet = new ColumnSet("hil_franchisee", "hil_claimperiod", "hil_claimheader", "hil_claimamount");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, _performaInvoiceId);
            queryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid(_fixedCompensationClaimCategory)); //Fixed Compensation
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            queryExp.LinkEntities.Add(lnkEntCH);

            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                int _gstCatg;
                decimal _amount = entcoll.Entities[0].GetAttributeValue<Money>("hil_claimamount").Value;
                EntityReference _erCP = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_franchisee");
                EntityReference _erCM = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_claimperiod");
                EntityReference _erCH = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_claimheader");
                Guid _ownerId = ((EntityReference)entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.ownerid").Value).Id;
                int jobCount = GetJobCountClaimMonth(_service, _erCP.Id, _erCM.Id);
                if (jobCount > 0)
                {
                    EntityCollection entColProdCatg = GetProdCatgWiseJobCountClaimMonth(_service, _erCP.Id, _erCM.Id);

                    if (entcoll.Entities[0].Contains("cp.hil_state") && entcoll.Entities[0].Contains("so.hil_state"))
                    {
                        EntityReference _erCPState = ((EntityReference)entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_state").Value);
                        EntityReference _erSOState = ((EntityReference)entcoll.Entities[0].GetAttributeValue<AliasedValue>("so.hil_state").Value);
                        _gstCatg = _erCPState.Id == _erSOState.Id ? 1 : 2;
                        QueryExpression queryExpTemp;
                        EntityCollection entcollTemp;
                        string _activityCode = string.Empty;
                        string _claimCategory = string.Empty;
                        int _prodCatgJobs = 0;
                        decimal _prodCatgAmount = 0;
                        decimal _prodCatgTotalAmount = 0;
                        int _rowCount = 1;
                        foreach (Entity ent in entColProdCatg.Entities)
                        {
                            _activityCode = "";
                            _claimCategory = "";
                            _prodCatgJobs = 0;
                            _prodCatgAmount = 0;

                            _prodCatgJobs = Convert.ToInt32(ent.GetAttributeValue<AliasedValue>("rowcount").Value);

                            _prodCatgAmount = Math.Round((_amount * _prodCatgJobs) / jobCount, 0);

                            _prodCatgTotalAmount += _prodCatgAmount;
                            if (_rowCount == entColProdCatg.Entities.Count)
                            {
                                _prodCatgAmount = _prodCatgAmount + (_amount - _prodCatgTotalAmount);
                            }

                            EntityReference _erProdCatg = (EntityReference)ent.GetAttributeValue<AliasedValue>("prodcatg").Value;

                            queryExpTemp = new QueryExpression("hil_claimpostingsetup");
                            queryExpTemp.ColumnSet = new ColumnSet("hil_activitycode");
                            queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExpTemp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, new Guid(JobCallSubType.Breakdown)); //Breakdown
                            queryExpTemp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, _erProdCatg.Id);
                            queryExpTemp.Criteria.AddCondition("hil_activitygstslab", ConditionOperator.Equal, _gstCatg);
                            queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            entcollTemp = _service.RetrieveMultiple(queryExpTemp);
                            if (entcollTemp.Entities.Count > 0)
                            {
                                _activityCode = entcollTemp.Entities[0].GetAttributeValue<string>("hil_activitycode");
                            }
                            _claimCategory = "824a1bcc-6ee5-ea11-a817-000d3af0501c"; //Fixed Compensation
                            ClaimLineDTO _claimLine = new ClaimLineDTO()
                            {
                                CallSubtype = new EntityReference("hil_callsubtype", new Guid(JobCallSubType.Breakdown)), //Breakdown
                                ChannelPartner = _erCP,
                                ClaimAmount = _prodCatgAmount,
                                ClaimCatwgory = _claimCategory,
                                ClaimPeriod = _erCM,
                                OwnerId = _ownerId,
                                PerformInvoiceId = _erCH,
                                ProdCatg = _erProdCatg,
                                ActivityCode = _activityCode
                            };
                            CreateOverheadClaimLine(_service, _claimLine);
                            _rowCount += 1;
                        }

                        #region Inactivate Fixed Compensation Line
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = entcoll.Entities[0].Id,
                                LogicalName = entcoll.Entities[0].LogicalName,
                            },
                            State = new OptionSetValue(1),
                            Status = new OptionSetValue(2)
                        };
                        _service.Execute(setStateRequest);
                        #endregion
                    }
                }
            }
        }

        public void GenerateClaimOverHeads(IOrganizationService _service, Guid _performaInvoiceId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;

            LinkEntity lnkEntCC = new LinkEntity
            {
                LinkFromEntityName = "hil_claimoverheadline",
                LinkToEntityName = "product",
                LinkFromAttributeName = "hil_productcategory",
                LinkToAttributeName = "productid",
                Columns = new ColumnSet("hil_claimcategory"),
                EntityAlias = "cc",
                JoinOperator = JoinOperator.LeftOuter
            };
            queryExp = new QueryExpression("hil_claimoverheadline");
            queryExp.ColumnSet = new ColumnSet("hil_productcategory", "hil_performainvoice", "hil_callsubtype", "hil_amount");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_performainvoice", ConditionOperator.Equal, _performaInvoiceId);
            queryExp.Criteria.AddCondition("hil_productcategory", ConditionOperator.NotNull);
            queryExp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.NotNull);
            queryExp.Criteria.AddCondition("hil_amount", ConditionOperator.NotNull);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            queryExp.LinkEntities.Add(lnkEntCC);

            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                LinkEntity lnkEntSO = new LinkEntity
                {
                    LinkFromEntityName = "account",
                    LinkToEntityName = "hil_salesoffice",
                    LinkFromAttributeName = "hil_salesoffice",
                    LinkToAttributeName = "hil_salesofficeid",
                    Columns = new ColumnSet("hil_state"),
                    EntityAlias = "so",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntCP = new LinkEntity
                {
                    LinkFromEntityName = "hil_claimheader",
                    LinkToEntityName = "account",
                    LinkFromAttributeName = "hil_franchisee",
                    LinkToAttributeName = "accountid",
                    Columns = new ColumnSet("hil_state", "hil_salesoffice", "ownerid"),
                    EntityAlias = "cp",
                    JoinOperator = JoinOperator.Inner
                };
                lnkEntCP.LinkEntities.Add(lnkEntSO);

                queryExp = new QueryExpression("hil_claimheader");
                queryExp.ColumnSet = new ColumnSet("hil_franchisee", "hil_fiscalmonth");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_claimheaderid", ConditionOperator.Equal, _performaInvoiceId);
                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                queryExp.LinkEntities.Add(lnkEntCP);
                EntityCollection entcollCH = _service.RetrieveMultiple(queryExp);
                if (entcollCH.Entities.Count > 0)
                {
                    int _gstCatg;
                    EntityReference _erCP = entcollCH.Entities[0].GetAttributeValue<EntityReference>("hil_franchisee");
                    EntityReference _erCM = entcollCH.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth");
                    Guid _ownerId = ((EntityReference)entcollCH.Entities[0].GetAttributeValue<AliasedValue>("cp.ownerid").Value).Id;

                    if (entcollCH.Entities[0].Contains("cp.hil_state") && entcollCH.Entities[0].Contains("so.hil_state"))
                    {
                        EntityReference _erCPState = ((EntityReference)entcollCH.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_state").Value);
                        EntityReference _erSOState = ((EntityReference)entcollCH.Entities[0].GetAttributeValue<AliasedValue>("so.hil_state").Value);
                        _gstCatg = _erCPState.Id == _erSOState.Id ? 1 : 2;
                        QueryExpression queryExpTemp;
                        EntityCollection entcollTemp;
                        string _activityCode = string.Empty;
                        string _claimCatwgory = string.Empty;

                        foreach (Entity ent in entcoll.Entities)
                        {
                            _activityCode = "";
                            _claimCatwgory = "";
                            queryExpTemp = new QueryExpression("hil_claimpostingsetup");
                            queryExpTemp.ColumnSet = new ColumnSet("hil_activitycode");
                            queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExpTemp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_callsubtype").Id);
                            queryExpTemp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_productcategory").Id);
                            queryExpTemp.Criteria.AddCondition("hil_activitygstslab", ConditionOperator.Equal, _gstCatg);
                            queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            entcollTemp = _service.RetrieveMultiple(queryExpTemp);
                            if (entcollTemp.Entities.Count > 0)
                            {
                                _activityCode = entcollTemp.Entities[0].GetAttributeValue<string>("hil_activitycode");
                            }
                            if (ent.Contains("cc.hil_claimcategory"))
                            {
                                _claimCatwgory = ((EntityReference)ent.GetAttributeValue<AliasedValue>("cc.hil_claimcategory").Value).Id.ToString();
                            }
                            else
                            {
                                _claimCatwgory = "DEAC05B5-8E68-EC11-8943-6045BDA5EF08"; //Lumpsum Payment
                            }
                            ClaimLineDTO _claimLine = new ClaimLineDTO()
                            {
                                CallSubtype = ent.GetAttributeValue<EntityReference>("hil_callsubtype"),
                                ChannelPartner = _erCP,
                                ClaimAmount = ent.GetAttributeValue<Money>("hil_amount").Value,
                                ClaimCatwgory = _claimCatwgory,
                                ClaimPeriod = _erCM,
                                OwnerId = _ownerId,
                                PerformInvoiceId = ent.GetAttributeValue<EntityReference>("hil_performainvoice"),
                                ProdCatg = ent.GetAttributeValue<EntityReference>("hil_productcategory"),
                                ActivityCode = _activityCode
                            };
                            CreateOverheadClaimLine(_service, _claimLine);
                        }
                    }
                }
            }
        }

        public void UpdatePerformaInvoice(IOrganizationService _service, Guid _performaInvoiceId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;
            EntityReference _erSalesOffice = null;
            EntityReference _erFiscalMonth = null;
            Guid _fixedCompensationId = new Guid("824A1BCC-6EE5-EA11-A817-000D3AF0501C");
            try
            {
                if (_performaInvoiceId != Guid.Empty)
                {
                    queryExp = new QueryExpression("hil_claimheader");
                    queryExp.ColumnSet = new ColumnSet("hil_fiscalmonth");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_claimheaderid", ConditionOperator.Equal, _performaInvoiceId);
                    entcoll = _service.RetrieveMultiple(queryExp);
                    if (entcoll.Entities.Count > 0)
                    {
                        _erFiscalMonth = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth");
                    }
                }
                else
                {
                    queryExp = new QueryExpression("hil_claimperiod");
                    queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    entcoll = _service.RetrieveMultiple(queryExp);
                    if (entcoll.Entities.Count > 0)
                    {
                        _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
                    }
                }
                if (_erFiscalMonth == null)
                {
                    //Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                    return;
                }

                if (_erFiscalMonth != null)
                {
                    LinkEntity lnkEntCP = new LinkEntity
                    {
                        LinkFromEntityName = "hil_claimheader",
                        LinkToEntityName = "account",
                        LinkFromAttributeName = "hil_franchisee",
                        LinkToAttributeName = "accountid",
                        Columns = new ColumnSet("hil_salesoffice"),
                        EntityAlias = "cp",
                        JoinOperator = JoinOperator.Inner
                    };
                    queryExp = new QueryExpression("hil_claimheader");
                    queryExp.ColumnSet = new ColumnSet("hil_claimheaderid", "hil_name");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    if (_performaInvoiceId != Guid.Empty)
                        queryExp.Criteria.AddCondition("hil_claimheaderid", ConditionOperator.Equal, _performaInvoiceId);

                    queryExp.Criteria.AddCondition("hil_fiscalmonth", ConditionOperator.Equal, _erFiscalMonth.Id);

                    queryExp.LinkEntities.Add(lnkEntCP);
                    entcoll = _service.RetrieveMultiple(queryExp);
                    int _cnt = 0, _jobCount;
                    Decimal _claimAmount, _fixedCharge, _overHeadsAmount, _netClaimAmount;

                    foreach (Entity ent in entcoll.Entities)
                    {
                        try
                        {
                            _cnt += 1;
                            //Console.WriteLine("Processing Performa Invoice# " + _cnt.ToString() + " : " + ent.GetAttributeValue<string>("hil_name"));

                            _performaInvoiceId = ent.Id;
                            _jobCount = 0;
                            _claimAmount = 0;
                            _fixedCharge = 0;
                            _overHeadsAmount = 0;
                            _netClaimAmount = 0;

                            if (ent.Contains("cp.hil_salesoffice"))
                            {
                                _erSalesOffice = (EntityReference)ent.GetAttributeValue<AliasedValue>("cp.hil_salesoffice").Value;
                            }
                            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                            <entity name='hil_claimline'>
                            <attribute name='hil_jobid' />
                            <filter type='and'>
                                <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_jobid' operator='not-null' />
                            </filter>
                            </entity>
                            </fetch>";
                            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                _jobCount = entcoll.Entities.Count;
                            }
                            _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                            <entity name='hil_claimline'>
                            <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                            <filter type='and'>
                                <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_claimcategory' operator='not-in'>
                                    <value uiname='Fixed Compensation' uitype='hil_claimcategory'>{824A1BCC-6EE5-EA11-A817-000D3AF0501C}</value>
                                    <value uiname='Lumpsum Payment' uitype='hil_claimcategory'>{DEAC05B5-8E68-EC11-8943-6045BDA5EF08}</value>
                                    <value uiname='Non-Standardized Charges' uitype='hil_claimcategory'>{501D69F1-8E68-EC11-8943-6045BDA5EF08}</value>
                                    <value uiname='Reverse Logistics' uitype='hil_claimcategory'>{51B1D88E-8E68-EC11-8943-6045BDA5EF08}</value>
                                    <value uiname='Abnormal Closure' uitype='hil_claimcategory'>{8dbfb0c4-33f3-ea11-a815-000d3af05a4b}</value>
                                </condition>
                            </filter>
                            </entity>
                            </fetch>";
                            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                if (entcoll.Entities[0].Contains("claimamount") && entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                    _claimAmount = ((Money)entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                            }
                            _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                                <entity name='hil_claimline'>
                                <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                                <filter type='and'>
                                    <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_claimcategory' operator='eq' value='{" + _fixedCompensationId + @"}' />
                                </filter>
                                </entity>
                                </fetch>";
                            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                if (entcoll.Entities[0].Contains("claimamount") && entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                    _fixedCharge = ((Money)entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                            }
                            _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                                <entity name='hil_claimoverheadline'>
                                <attribute name='hil_amount' alias='claimamount' aggregate='sum'/> 
                                <filter type='and'>
                                    <condition attribute='hil_performainvoice' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                                </entity>
                                </fetch>";
                            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                if (entcoll.Entities[0].Contains("claimamount") && entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                    _overHeadsAmount = ((Money)entcoll.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                            }
                            _netClaimAmount = _claimAmount + _fixedCharge + _overHeadsAmount;

                            Entity _entPerformaInv = new Entity("hil_claimheader", _performaInvoiceId);
                            _entPerformaInv["hil_branch"] = _erSalesOffice;
                            _entPerformaInv["hil_totaljobsclosed"] = _jobCount;
                            _entPerformaInv["hil_totalclaimamount"] = _claimAmount;
                            _entPerformaInv["hil_fixedcharges"] = _fixedCharge;
                            _entPerformaInv["hil_totalclaimvalue"] = _netClaimAmount;
                            _entPerformaInv["hil_expenseoverheads"] = _overHeadsAmount;
                            _service.Update(_entPerformaInv);
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidPluginExecutionException("Havells_Plugin.ClaimOperation.updatePerformaInvoice.Execute" + ex.Message);
                            //Console.WriteLine("ERRPR!!! " + ex.Message + " Processing Performa Invoice# " + _cnt.ToString() + " : " + ent.GetAttributeValue<string>("hil_name"));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ClaimOperation.updatePerformaInvoice.Execute" + ex.Message);
            }
        }

        public void GenerateClaimSummary(IOrganizationService _service, Guid _performaInvoiceId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;

            LinkEntity lnkEntSO = new LinkEntity
            {
                LinkFromEntityName = "account",
                LinkToEntityName = "hil_salesoffice",
                LinkFromAttributeName = "hil_salesoffice",
                LinkToAttributeName = "hil_salesofficeid",
                Columns = new ColumnSet("hil_sapcode"),
                EntityAlias = "so",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCP = new LinkEntity
            {
                LinkFromEntityName = "hil_claimheader",
                LinkToEntityName = "account",
                LinkFromAttributeName = "hil_franchisee",
                LinkToAttributeName = "accountid",
                Columns = new ColumnSet("hil_salesoffice", "ownerid", "hil_vendorcode"),
                EntityAlias = "cp",
                JoinOperator = JoinOperator.Inner
            };


            lnkEntCP.LinkEntities.Add(lnkEntSO);

            queryExp = new QueryExpression("hil_claimheader");
            queryExp.ColumnSet = new ColumnSet("hil_fiscalmonth", "hil_franchisee", "hil_performastatus");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_claimheaderid", ConditionOperator.Equal, _performaInvoiceId);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            queryExp.LinkEntities.Add(lnkEntCP);

            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                string _vendorCode = string.Empty;
                string _fetchXML = string.Empty;
                if (entcoll.Entities[0].Contains("cp.hil_vendorcode") && entcoll.Entities[0].Contains("so.hil_sapcode"))
                {
                    #region Deactivating all Claim Summary Lines
                    _fetchXML = @"<fetch distinct='false' mapping='logical'>
                        <entity name='hil_claimsummary'>  
                        <attribute name='hil_claimsummaryid' />
                        <filter type='and'>
                            <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
                            <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                        </entity>
                    </fetch>";
                    EntityCollection entcollCS = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    foreach (Entity ent in entcollCS.Entities)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = ent.Id,
                                LogicalName = ent.LogicalName,
                            },
                            State = new OptionSetValue(1),
                            Status = new OptionSetValue(2)
                        };
                        _service.Execute(setStateRequest);
                    }
                    #endregion

                    int _rowCount = 0;
                    _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                      <entity name='hil_claimline'>
                        <attribute name='hil_activitycode' groupby='true' alias='activity'/>
                        <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/>
                        <filter type='and'>
                          <condition attribute='hil_claimheader' operator='eq' value='{" + _performaInvoiceId + @"}' />
                          <condition attribute='hil_activitycode' operator='not-null' />
                          <condition attribute='statecode' operator='eq' value='0' />
                          <condition attribute='hil_productcategory' operator='not-null' />
                        </filter>
                        <link-entity name='product' from='productid' to='hil_productcategory' visible='false' link-type='outer' alias='pd'>
                          <attribute name='hil_sapcode' groupby='true' alias='prodcode'/>
                        </link-entity>
                      </entity>
                    </fetch>";
                    EntityCollection entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    Entity _entClaimSum = null;
                    Money _claimAmount = null;

                    foreach (Entity ent in entcoll1.Entities)
                    {
                        _claimAmount = ((Money)ent.GetAttributeValue<AliasedValue>("claimamount").Value);
                        if (_claimAmount.Value == 0)
                            continue;

                        _entClaimSum = new Entity("hil_claimsummary");

                        _entClaimSum["hil_salesoffice"] = (EntityReference)entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_salesoffice").Value;
                        _entClaimSum["hil_performainvoiceid"] = entcoll.Entities[0].ToEntityReference();
                        _entClaimSum["hil_franchise"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_franchisee");
                        _entClaimSum["hil_claimmonth"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth");
                        _entClaimSum["hil_franchisecode"] = entcoll.Entities[0].GetAttributeValue<AliasedValue>("cp.hil_vendorcode").Value.ToString();
                        _entClaimSum["hil_claimmonthcode"] = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_fiscalmonth").Name;
                        _entClaimSum["hil_salesofficecode"] = entcoll.Entities[0].GetAttributeValue<AliasedValue>("so.hil_sapcode").Value.ToString();
                        _entClaimSum["hil_activitycode"] = ent.GetAttributeValue<AliasedValue>("activity").Value.ToString();
                        _entClaimSum["hil_productdivision"] = ent.GetAttributeValue<AliasedValue>("prodcode").Value.ToString();
                        _entClaimSum["hil_claimamount"] = _claimAmount;
                        _entClaimSum["hil_qty"] = 1;
                        _service.Create(_entClaimSum);

                        _rowCount++;
                    }

                    #region Adjusting Negative Claim amount with Poisitive
                    _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                      <entity name='hil_claimsummary'>
                        <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/>
                        <filter type='and'>
                          <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
                          <condition attribute='hil_claimamount' operator='lt' value='0' />
                        </filter>
                      </entity>
                    </fetch>";
                    entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll1.Entities.Count > 0)
                    {
                        if (entcoll1.Entities[0].Contains("claimamount"))
                        {
                            if (entcoll1.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                            {
                                Money _negativeClaimAmount = null;
                                _negativeClaimAmount = ((Money)entcoll1.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value);

                                _fetchXML = @"<fetch distinct='false' mapping='logical'>
                                    <entity name='hil_claimsummary'>
                                    <attribute name='hil_claimamount' />
                                    <filter type='and'>
                                        <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                        <condition attribute='hil_claimamount' operator='ge' value='" + (_negativeClaimAmount.Value * -1) + @"' />
                                    </filter>
                                    </entity>
                                    </fetch>";
                                entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entcoll1.Entities.Count > 0)
                                {
                                    _claimAmount = entcoll1.Entities[0].GetAttributeValue<Money>("hil_claimamount");
                                    entcoll1.Entities[0]["hil_claimamount"] = _claimAmount.Value - (_negativeClaimAmount.Value * -1);
                                    _service.Update(entcoll1.Entities[0]);
                                }

                                _fetchXML = @"<fetch distinct='false' mapping='logical'>
                                  <entity name='hil_claimsummary'>  
                                    <attribute name='hil_claimsummaryid' />
                                    <filter type='and'>
                                      <condition attribute='hil_performainvoiceid' operator='eq' value='{" + _performaInvoiceId + @"}' />
                                      <condition attribute='hil_claimamount' operator='lt' value='0' />
                                    </filter>
                                  </entity>
                                </fetch>";
                                entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                foreach (Entity ent in entcoll1.Entities)
                                {
                                    SetStateRequest setStateRequest = new SetStateRequest()
                                    {
                                        EntityMoniker = new EntityReference
                                        {
                                            Id = ent.Id,
                                            LogicalName = ent.LogicalName,
                                        },
                                        State = new OptionSetValue(1),
                                        Status = new OptionSetValue(2)
                                    };
                                    _service.Execute(setStateRequest);
                                }
                            }
                        }
                    }
                    #endregion

                }
            }
        }

        private int GetJobCountClaimMonth(IOrganizationService _service, Guid _franchiseId, Guid _claimMonthId)
        {
            QueryExpression queryExp;
            EntityCollection entcoll;

            try
            {
                queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                queryExp.ColumnSet = new ColumnSet(false);
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, _franchiseId);
                queryExp.Criteria.AddCondition("hil_tatcategory", ConditionOperator.NotNull);
                queryExp.Criteria.AddCondition("hil_isocr", ConditionOperator.NotEqual, true);
                queryExp.Criteria.AddCondition("hil_claimstatus", ConditionOperator.Equal, 4);// Claim Approved
                queryExp.Criteria.AddCondition("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1")); //Closed
                queryExp.Criteria.AddCondition("hil_fiscalmonth", ConditionOperator.Equal, _claimMonthId); //
                entcoll = _service.RetrieveMultiple(queryExp);
                return entcoll.Entities.Count;
            }
            catch
            {
                return 0;
            }
        }
        private EntityCollection GetProdCatgWiseJobCountClaimMonth(IOrganizationService _service, Guid _franchiseId, Guid _claimMonthId)
        {
            EntityCollection entcoll;

            try
            {
                string _fetchXML = @"<fetch distinct='false' mapping='logical' aggregate='true'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_workorderid' alias='rowcount' aggregate='count'/> 
                <attribute name='hil_productcategory' alias='prodcatg' groupby='true' /> 
                <filter type='and'>
                    <condition attribute='hil_isocr' operator='ne' value='1' />
                    <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                    <condition attribute='hil_claimstatus' operator='eq' value='4' />
                    <condition attribute='hil_tatcategory' operator='not-null' />
                    <condition attribute='hil_owneraccount' operator='eq' value='{" + _franchiseId + @"}' />
                    <condition attribute='hil_fiscalmonth' operator='eq' value='{" + _claimMonthId + @"}' />
                </filter>
                </entity>
                </fetch>";
                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                return entcoll;
            }
            catch
            {
                return null;
            }
        }

        private Guid ClaimLineExistsOverhead(IOrganizationService _service, ClaimLineDTO _claimLineData)
        {
            Guid _returnVal = Guid.Empty;
            QueryExpression queryExp;
            EntityCollection entcoll;

            try
            {
                queryExp = new QueryExpression("hil_claimline");
                queryExp.ColumnSet = new ColumnSet("hil_claimperiod");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                queryExp.Criteria.AddCondition("hil_franchisee", ConditionOperator.Equal, _claimLineData.ChannelPartner.Id);
                queryExp.Criteria.AddCondition("hil_claimperiod", ConditionOperator.Equal, _claimLineData.ClaimPeriod.Id);
                queryExp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, _claimLineData.ProdCatg.Id);
                queryExp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, _claimLineData.CallSubtype.Id);

                if (_claimLineData.ClaimCatwgory != null)
                    queryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid(_claimLineData.ClaimCatwgory));

                queryExp.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, _claimLineData.PerformInvoiceId.Id);
                queryExp.Criteria.AddCondition("hil_claimamount", ConditionOperator.Equal, _claimLineData.ClaimAmount);
                entcoll = _service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _returnVal = entcoll.Entities[0].Id;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return _returnVal;
        }
        private void AssignRecord(String TargetLogicalName, Guid AssigneeId, Guid TargetId, IOrganizationService service)
        {
            try
            {
                AssignRequest assign = new AssignRequest();
                assign.Assignee = new EntityReference(SystemUser.EntityLogicalName, AssigneeId); //User or team
                assign.Target = new EntityReference(TargetLogicalName, TargetId); //Record to be assigned
                service.Execute(assign);
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
        private void CreateOverheadClaimLine(IOrganizationService _service, ClaimLineDTO _claimLineData)
        {
            try
            {
                Entity entObj = new Entity("hil_claimline");
                Guid _claimLineId = ClaimLineExistsOverhead(_service, _claimLineData);
                if (_claimLineId == Guid.Empty)
                {
                    Entity _entClaimCatg = _service.Retrieve("hil_claimcategory", new Guid(_claimLineData.ClaimCatwgory), new ColumnSet("hil_claimtype"));
                    entObj["hil_franchisee"] = _claimLineData.ChannelPartner;
                    entObj["hil_claimperiod"] = _claimLineData.ClaimPeriod;
                    entObj["hil_productcategory"] = _claimLineData.ProdCatg;
                    entObj["hil_callsubtype"] = _claimLineData.CallSubtype;
                    entObj["hil_claimcategory"] = new EntityReference("hil_claimcategory", new Guid(_claimLineData.ClaimCatwgory));
                    entObj["hil_claimheader"] = _claimLineData.PerformInvoiceId;
                    if (_entClaimCatg.GetAttributeValue<EntityReference>("hil_claimtype").Name == "Penalty")
                        entObj["hil_claimamount"] = new Money(_claimLineData.ClaimAmount * -1);
                    else
                        entObj["hil_claimamount"] = new Money(_claimLineData.ClaimAmount);
                    entObj["hil_activitycode"] = _claimLineData.ActivityCode;
                    _claimLineId = _service.Create(entObj);

                    AssignRecord("hil_claimline", _claimLineData.OwnerId, _claimLineId, _service);
                }
                else
                {
                    entObj.Id = _claimLineId;
                    entObj["hil_activitycode"] = _claimLineData.ActivityCode;
                    entObj["hil_claimamount"] = new Money(_claimLineData.ClaimAmount);
                    _service.Update(entObj);
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
    public static class JobCallSubType
    {
        public const string PMS = "E2129D79-3C0B-E911-A94E-000D3AF06CD4";
        public const string Breakdown = "6560565A-3C0B-E911-A94E-000D3AF06CD4";
        public const string DealerStockRepair = "8D80346B-3C0B-E911-A94E-000D3AF06CD4";
        public const string AMCCall = "55a71a52-3c0b-e911-a94e-000d3af06cd4";
        public const string Installation = "e3129d79-3c0b-e911-a94e-000d3af06cd4";
        public const string PDI = "ce45f586-3c0b-e911-a94e-000d3af06cd4";
    }
    //public class ClaimLineDTO
    //{
    //    public Guid OwnerId { get; set; }
    //    public EntityReference PerformInvoiceId { get; set; }
    //    public decimal ClaimAmount { get; set; }
    //    public string ClaimCatwgory { get; set; }
    //    public EntityReference ClaimPeriod { get; set; }
    //    public EntityReference ChannelPartner { get; set; }
    //    public EntityReference ProdCatg { get; set; }
    //    public EntityReference CallSubtype { get; set; }
    //    public string ActivityCode { get; set; }
    //}
    public class ClaimSummaryData
    {
        public string ClaimId { get; set; }
        public string ZMONTH { get; set; }
        public string SALES_OFFICE { get; set; }
        public string VENDOR_CODE { get; set; }
        public string ACTIVITY_NUMBER { get; set; }
        public string ACTIVITY_QTY { get; set; }
        public string Amount { get; set; }
        public string DIVISION { get; set; }
    }
    public class RootClaimSummary
    {
        public List<ClaimSummaryData> LT_TABLE { get; set; }

    }
    public class Info
    {
        public string CLAIM_ID { get; set; }
        public string EBELNO { get; set; }
        public List<MSG> msgss { get; set; }

    }
    public class MSG
    {
        public string ID { get; set; }
        public string MESSAGE { get; set; }
        public int NUMBER { get; set; }
        public string TYPE { get; set; }
    }
}
