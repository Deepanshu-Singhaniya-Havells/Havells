function onload(executionContext) {
    debugger;
    try {
        setDummyAccount();
        defaultcustomer();
        makeAssociatedPdtMandatory();
        test();
        ForEmailconvertedJob();
        SetPriceList();
        RestrictAssign();
        setDefaultStatus();
        disableformControls();
        OnLoadOCR();
        var FormType = Xrm.Page.ui.getFormType();
        if (FormType != 2) {
            Xrm.Page.getControl("hil_kkgcode").setDisabled(true);
        }
        if (FormType == 1) {
            Xrm.Page.getAttribute("hil_claimstatus").setValue(0);
            Xrm.Page.getAttribute("hil_callertype").setValue(910590001);
            Xrm.Page.getAttribute("hil_callingnumber").setRequiredLevel("required");
            Xrm.Page.getControl("hil_schemecode").setDisabled(true);
        }
        if (FormType != 1) {
            Xrm.Page.getControl("hil_purchasedate").setDisabled(true);
            //lockFieldsOnceWorkDone();
        }
        if (!isControlNull("hil_productcatsubcatmapping")) {
            Xrm.Page.getAttribute("hil_productcatsubcatmapping").addOnChange(SetCatSubCat);
        }
        if (!isControlNull("hil_address")) {
            Xrm.Page.getAttribute("hil_address").addOnChange(GetAddress);
        }
        if (!isControlNull("hil_quantity")) {
            Xrm.Page.getAttribute("hil_quantity").addOnChange(QuantityOnChange);
        }
        if (!isControlNull("hil_kkgcode")) {
            Xrm.Page.getAttribute("hil_kkgcode").addOnChange(ValidateKKG);
        }
        PreFilterCustomerAsset();
        PreFilterJobReference();

        var formContext = executionContext.getFormContext();
        formContext.getControl("hil_sourceofjob").setDisabled(true);
        formContext.getControl("hil_requesttype").setDisabled(true);

        if (executionContext != null) {
            var formContext = executionContext.getFormContext();
            if (!isControlNull("hil_natureofcomplaint")) {
                formContext.getAttribute("hil_natureofcomplaint").addOnChange(PopulateCallSubType);
            }
            if (!isControlNull("hil_customerref")) {
                formContext.getAttribute("hil_customerref").addOnChange(PreFilterCustomerAsset);
            }
            if (!isControlNull("hil_productsubcategory")) {
                formContext.getAttribute("hil_productsubcategory").addOnChange(PreFilterCustomerAsset);
            }
            if (!isControlNull("hil_productcategory")) {
                formContext.getAttribute("hil_productcategory").addOnChange(PreFilterCustomerAsset);
            }
            if (!isControlNull("msdyn_customerasset")) {
                formContext.getAttribute("msdyn_customerasset").addOnChange(OnAssetChange);
            }
            if (!isControlNull("msdyn_customerasset")) {
                formContext.getAttribute("msdyn_customerasset").addOnChange(OnAssetChangeIfNotVerified);
            }
            if (!isControlNull("hil_branchheadapproval")) {
                formContext.getAttribute("hil_branchheadapproval").addOnChange(OnBranchHeadApproval);
            }
            if (!isControlNull("ownerid")) {
                formContext.getAttribute("ownerid").addOnChange(TechnicianName);
            }
            if (!isControlNull("hil_assigntobranchhead")) {
                formContext.getAttribute("hil_assigntobranchhead").addOnChange(AssignToBranchHead);
            }
            if (!isControlNull("hil_isocr")) {
                formContext.getAttribute("hil_isocr").addOnChange(OCROnChaneValidation);
            }
        }
        AllowCallCentertoModifyContactDetails();
        PreFilterSchemeCodes();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function AllowCallCentertoModifyContactDetails() {
    try {
        debugger;
        var FormType = Xrm.Page.ui.getFormType();
        var LoggedIn = Xrm.Page.context.getUserId();
        var subStatus = Xrm.Page.getAttribute("msdyn_substatus").getValue();
        var StatusName = subStatus[0].name.toUpperCase();
        if (FormType != 1) {
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/SystemUserSet(guid'" + LoggedIn + "')?$select=hil_JobCancelAuth,PositionId,position_users/Name&$expand=position_users", true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.responseText).d;
                        var positionId = result.PositionId;
                        var jobCancelAuth = result.hil_JobCancelAuth;
                        var position_users_Name = result.position_users.Name.toUpperCase();
                        if (position_users_Name == "CALL CENTER") {
                            if (StatusName != "KKG AUDIT FAILED" && StatusName != "CLOSED" && StatusName != "CANCELED") {
                                Xrm.Page.getControl("hil_mobilenumber").setDisabled(false);
                                Xrm.Page.getControl("hil_callingnumber").setDisabled(false);
                            }
                            else {
                                Xrm.Page.getControl("hil_mobilenumber").setDisabled(true);
                                Xrm.Page.getControl("hil_callingnumber").setDisabled(true);
                            }
                        }
                        else {
                            Xrm.Page.getControl("hil_mobilenumber").setDisabled(true);
                            Xrm.Page.getControl("hil_callingnumber").setDisabled(true);
                        }
                    }
                }
            };
            req.send();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}
function PreFilterJobReference() {
    debugger;
    try {
        if (Xrm.Page.getControl("msdyn_parentworkorder") != null) {
            Xrm.Page.getControl("msdyn_parentworkorder").addPreSearch(function () {
                var CustomerRef = Xrm.Page.getAttribute("hil_customerref").getValue();
                var CustomerAsset = Xrm.Page.getAttribute("msdyn_customerasset").getValue();
                var ProductCategory = Xrm.Page.getAttribute("hil_productcategory").getValue();
                var filterStr = "<filter type='and'>";
                filterStr = filterStr + "<condition attribute='msdyn_substatus' operator='not-in'>"
                filterStr = filterStr + "<value >{529676E1-CD9A-E811-A966-000D3AF06848}</value>" //WorkDone
                filterStr = filterStr + "<value >{8FC640B3-F29A-E811-A963-000D3AF06236}</value>" //Closed
                filterStr = filterStr + "</condition>"
                var LookupId = null;
                if (CustomerRef != null) {
                    LookupId = CustomerRef[0].id;
                    filterStr = filterStr + "<condition attribute='hil_customerref' operator='eq'  value='" + LookupId + "' />"
                }
                if (CustomerAsset != null) {
                    LookupId = CustomerAsset[0].id;
                    filterStr = filterStr + "<condition attribute='msdyn_customerasset' operator='eq'  value='" + LookupId + "' />"
                }
                if (ProductCategory != null) {
                    LookupId = ProductCategory[0].id;
                    filterStr = filterStr + "<condition attribute='hil_productcategory' operator='eq'  value='" + LookupId + "' />"
                }
                if (LookupId == null) {
                    filterStr = filterStr + "<condition attribute='hil_customerref' operator='eq'  value='{00000000-0000-0000-0000-000000000000}' />"
                }
                var filterStr = filterStr + "</filter>";
                Xrm.Page.getControl("msdyn_parentworkorder").addCustomFilter(filterStr);
            });
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function ReTriggerAssignmentIncomingSMS(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        var PCat = formContext.getAttribute("hil_productcategory").getValue();
        var SrvAdd = formContext.getAttribute("hil_address").getValue();
        var Source = formContext.getAttribute("hil_sourceofjob").getValue();
        var IfAlreadyAssigned = formContext.getAttribute("hil_automaticassign").getValue();
        if (PCat != null && SrvAdd != null && IfAlreadyAssigned != 1) {
            if (Source == 5 || Source == 3) {
                formContext.getAttribute("hil_automaticassign").setValue(1);
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OCROnChaneValidation() {
    try {
        var OCR = Xrm.Page.getAttribute("hil_isocr").getValue();
        var Remarks = Xrm.Page.getAttribute("hil_closureremarks").getValue();
        var CurrentUserId = Xrm.Page.context.getUserId();
        var CurrentUserName = Xrm.Page.context.getUserName();
        var StatusName = "Closed";
        var StatusId = null;
        var currentDate = new Date();
        //alert(currentDate);
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/msdyn_workordersubstatusSet?$select=msdyn_workordersubstatusId&$filter=msdyn_name eq 'Closed'&$top=1", true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;
                if (this.status === 200) {
                    var returned = JSON.parse(this.responseText).d;
                    var results = returned.results;
                    for (var i = 0; i < 1; i++) {
                        var msdyn_workordersubstatusId = results[i].msdyn_workordersubstatusId;
                        StatusId = msdyn_workordersubstatusId;
                        if (OCR == true) {
                            Xrm.Utility.confirmDialog("Are you sure you want to close this job as OCR? No Claim will be Calculated!!!",

                                function () {
                                    if (Remarks != null) {
                                        Xrm.Page.getControl("hil_isocr").setDisabled(true);
                                        var object = new Array();
                                        object[0] = new Object();
                                        object[0].id = CurrentUserId;
                                        object[0].name = CurrentUserName;
                                        object[0].entityType = "systemuser";
                                        Xrm.Page.getAttribute("ownerid").setValue(object);
                                        var object1 = new Array();
                                        object1[0] = new Object();
                                        object1[0].id = StatusId;
                                        object1[0].name = "Closed";
                                        object1[0].entityType = "msdyn_workordersubstatus";
                                        Xrm.Page.getAttribute("msdyn_substatus").setValue(object1);
                                        Xrm.Page.getAttribute("hil_jobclosuredon").setValue(currentDate);
                                        Xrm.Page.getAttribute("hil_jobclosuredon").setSubmitMode("always");
                                        Xrm.Page.getAttribute("msdyn_timeclosed").setValue(currentDate);
                                        Xrm.Page.getAttribute("msdyn_timeclosed").setSubmitMode("always");
                                        Xrm.Page.getAttribute("msdyn_substatus").setSubmitMode("always");
                                        Xrm.Page.getAttribute("ownerid").setSubmitMode("always");
                                        Xrm.Page.data.entity.save();
                                    }
                                    else {
                                        alert("Please add Closure Remarks before closing this Job");
                                        Xrm.Page.getAttribute("hil_isocr").setValue(false);
                                        Xrm.Page.getAttribute("hil_closureremarks").setRequiredLevel("required");
                                        Xrm.Page.data.entity.save();
                                    }
                                },

                                function () { });
                        }
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OnLoadOCR() {
    try {
        if (Xrm.Page.getAttribute("msdyn_substatus") != null) {
            var SubStatus = Xrm.Page.getAttribute("msdyn_substatus").getValue();
            if (SubStatus != null) {
                var StatusName = SubStatus[0].name;
                if (StatusName == "Registered" || StatusName == "Work Allocated" || StatusName == "Closed" || StatusName == "Parts Fulfilled" || StatusName == "Pending For allocation" || StatusName == "Work Initiated") {
                    if (Xrm.Page.getControl("hil_isocr") != null) {
                        Xrm.Page.getControl("hil_isocr").setVisible(true);
                    }
                }
                else {
                    if (Xrm.Page.getControl("hil_isocr") != null) {
                        Xrm.Page.getControl("hil_isocr").setVisible(false);
                    }
                }
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OnAssetChangeIfNotVerified() {
    try {
        var iAsset = Xrm.Page.getAttribute("msdyn_customerasset").getValue();
        if (iAsset != null) {
            var iAssetId = iAsset[0].id;
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/msdyn_customerassetSet(guid'" + iAssetId + "')?$select=statuscode", true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.responseText).d;
                        var statuscode = result.statuscode;
                        if (statuscode.Value == 910590000) {
                            alert("ASSET NOT APPROVED");
                            Xrm.Page.getAttribute("msdyn_customerasset").setValue(null);
                        }
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
        PreFilterJobReference();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function TechnicianName() {
    try {
        var OwnerId = Xrm.Page.getAttribute("ownerid").getValue();
        var Type = Xrm.Page.getAttribute("hil_typeofassignee").getValue();
        if (Type[0].name == "Franchise") {
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/hil_technicianSet?$select=hil_name&$filter=hil_UserID/Id eq (guid'" + OwnerId[0].id + "')", true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var returned = JSON.parse(this.responseText).d;
                        var results = returned.results;
                        for (var i = 0; i < results.length; i++) {
                            var name = results[i].hil_name;
                            Xrm.Page.getAttribute("hil_technicianname").setValue(name);
                        }
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OnBranchHeadApproval() {
    try {
        Xrm.Utility.confirmDialog("Are you Sure you Want to Approve?",

            function () {
                Xrm.Page.getAttribute("hil_branchheadapproval").setValue(1);
                Xrm.Page.data.entity.save();
                alert("Approved");
            },

            function () {
                Xrm.Page.getAttribute("hil_branchheadapproval").setValue(null);
                Xrm.Page.data.entity.save();
                //alert("Disapproved");
            });
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function ForEmailconvertedJob() {
    try {
        var FormType = Xrm.Page.ui.getFormType();
        if (FormType == 1) {
            var EmailContact = Xrm.Page.getAttribute("hil_emailcustomer").getValue();
            var RegEmail = Xrm.Page.getAttribute("hil_regardingemail").getValue();
            if (EmailContact != null && RegEmail != null) {
                var EmailId = EmailContact[0].id;
                var EmailName = EmailContact[0].name;
                var object = new Array();
                object[0] = new Object();
                object[0].id = EmailId;
                object[0].name = EmailName;
                object[0].entityType = "contact";
                Xrm.Page.getAttribute("hil_customerref").setValue(object);
                Xrm.Page.getAttribute("hil_sourceofjob").setValue(1);
                GetDataFromcontact(EmailId);
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OnAssetChange() {
    try {
        var Asset = Xrm.Page.getAttribute("msdyn_customerasset").getValue();
        if (Asset != null) {
            var AssetId = Asset[0].id;
            //alert(AssetId);
            var req2 = new XMLHttpRequest();
            req2.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/msdyn_customerassetSet(guid'" + AssetId + "')?$select=hil_InvoiceDate,hil_ModelName", true);
            req2.setRequestHeader("Accept", "application/json");
            req2.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req2.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.responseText).d;
                        var InvoiceDate = result.hil_InvoiceDate;
                        var ModelName = result.hil_ModelName;
                        var ProdSC = Xrm.Page.getAttribute("hil_productsubcategory").getValue();
                        var date1 = new Date(parseInt(InvoiceDate.substr(6)));
                        var PDate = date1;
                        //alert(PDate);
                        Xrm.Page.getAttribute("hil_purchasedate").setValue(PDate);
                        Xrm.Page.getAttribute("hil_modelname").setValue(ModelName);
                        if (PDate != null) {
                            var DateNow = new Date();
                            var NowYear = DateNow.getFullYear();
                            var NowMonth = DateNow.getMonth() + 1;
                            var PYear = PDate.getFullYear();
                            var PMonth = PDate.getMonth() + 1;
                            var MonDiff = (NowYear - PYear) * 12 + NowMonth - PMonth;
                            if (ProdSC != null) {
                                var ProdId = ProdSC[0].id;
                                ProdId = ProdId.replace("{", "");
                                ProdId = ProdId.replace("}", "");
                                var req = new XMLHttpRequest();
                                req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/hil_warrantytemplateSet?$select=hil_Product,hil_type,hil_WarrantyPeriod&$filter=hil_Product/Id eq (guid'" + ProdId + "') and hil_type/Value eq 1", true);
                                req.setRequestHeader("Accept", "application/json");
                                req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                                req.onreadystatechange = function () {
                                    if (this.readyState === 4) {
                                        this.onreadystatechange = null;
                                        if (this.status === 200) {
                                            var returned = JSON.parse(this.responseText).d;
                                            var results = returned.results;
                                            for (var i = 0; i < results.length; i++) {
                                                var type = results[i].hil_type;
                                                var typeValue = type.Value;
                                                var WarrantyPeriod = results[i].hil_WarrantyPeriod;
                                                if (MonDiff < WarrantyPeriod) {
                                                    //alert("IN WARRANTY");
                                                    Xrm.Page.getAttribute("hil_warrantystatus").setValue(1);
                                                }
                                                else {
                                                    //alert("Hi");
                                                    var SubCat = Xrm.Page.getAttribute("hil_productsubcategory").getValue();
                                                    if (SubCat != null) {
                                                        var SubCatId = SubCat[0].id;
                                                        SubCatId = SubCatId.replace("{", "");
                                                        SubCatId = SubCatId.replace("}", "");
                                                        //alert(SubCatId);
                                                        var req1 = new XMLHttpRequest();
                                                        req1.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/products(" + SubCatId + ")?$select=hil_standardvisitcharge", true);
                                                        req1.setRequestHeader("OData-MaxVersion", "4.0");
                                                        req1.setRequestHeader("OData-Version", "4.0");
                                                        req1.setRequestHeader("Accept", "application/json");
                                                        req1.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                                                        req1.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
                                                        req1.onreadystatechange = function () {
                                                            if (this.readyState === 4) {
                                                                req1.onreadystatechange = null;
                                                                if (this.status === 200) {
                                                                    var result1 = JSON.parse(this.response);
                                                                    var hil_standardvisitcharge = result1["hil_standardvisitcharge"];
                                                                    var hil_standardvisitcharge_formatted = result1["hil_standardvisitcharge@OData.Community.Display.V1.FormattedValue"];
                                                                    alert("OUT WARRANTY\nStandard Visit Charges may be applicable\nVisit Charge Amount Rs." + hil_standardvisitcharge_formatted);
                                                                    Xrm.Page.getAttribute("hil_warrantystatus").setValue(2);
                                                                }
                                                                else {
                                                                    Xrm.Utility.alertDialog(this.statusText);
                                                                }
                                                            }
                                                        };
                                                        req1.send();
                                                    }
                                                }
                                            }
                                        }
                                        else {
                                            Xrm.Utility.alertDialog(this.statusText);
                                        }
                                    }
                                };
                                req.send();
                            }
                        }
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req2.send();
        }
        else {
            Xrm.Page.getAttribute("hil_purchasedate").setValue(null);
            Xrm.Page.getAttribute("hil_warrantystatus").setValue(null);
            Xrm.Page.getAttribute("hil_modelname").setValue(null);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function PopulateCallSubType(executionContext) {
    debugger;
    try {
        var formContext = executionContext.getFormContext()
        var Noc = formContext.getAttribute("hil_natureofcomplaint").getValue();
        if (Noc != null) {
            Noc = Noc[0].id;
            Noc = Noc.replace("{", '').replace("}", '');
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/hil_natureofcomplaints?$select=_hil_callsubtype_value&$filter=hil_natureofcomplaintid eq " + Noc + "", true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var results = JSON.parse(this.response);
                        for (var i = 0; i < results.value.length; i++) {
                            var _hil_callsubtype_value = results.value[i]["_hil_callsubtype_value"];
                            var _hil_callsubtype_value_formatted = results.value[i]["_hil_callsubtype_value@OData.Community.Display.V1.FormattedValue"];
                            var _hil_callsubtype_value_lookuplogicalname = results.value[i]["_hil_callsubtype_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
                            if (_hil_callsubtype_value != null) Xrm.Page.getAttribute("hil_callsubtype").setValue([
                                {
                                    id: _hil_callsubtype_value,
                                    name: _hil_callsubtype_value_formatted,
                                    entityType: _hil_callsubtype_value_lookuplogicalname
                                }]);
                        }
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
        else {
            Xrm.Page.getAttribute("hil_callsubtype").setValue(null);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function onAddressSelect(executionContext) {
    debugger;
    try {
        var formContext = executionContext.getFormContext(); // get formContext
        var ConatctId = formContext.getAttribute("hil_customerref").getValue();
        if (ConatctId != null) {
            ConatctId = ConatctId[0].id;
            ConatctId = ConatctId.replace("{", '').replace("}", '');
            var AddressId = formContext.getAttribute("hil_address").getValue();
            if (AddressId != null) {
                AddressId = AddressId[0].id;
                AddressId = AddressId.replace("{", '').replace("}", '');
                var entity = {};
                entity["hil_customer_contact@odata.bind"] = "/contacts(" + ConatctId + ")";
                var req = new XMLHttpRequest();
                req.open("PATCH", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/hil_addresses(" + AddressId + ")", true);
                req.setRequestHeader("OData-MaxVersion", "4.0");
                req.setRequestHeader("OData-Version", "4.0");
                req.setRequestHeader("Accept", "application/json");
                req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                req.onreadystatechange = function () {
                    if (this.readyState === 4) {
                        req.onreadystatechange = null;
                        if (this.status === 204) {
                            //alert("Success");
                            //Success - No Return Data - Do Something
                        }
                        else {
                            Xrm.Utility.alertDialog(this.statusText);
                        }
                    }
                };
                req.send(JSON.stringify(entity));
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function suppress_C2D_Popout() {
    debugger;
    //    parent.Mscrm.ReadFormUtilities.handlePhoneNumberClick = function () { return; }
    //    parent.Mscrm.ReadFormUtilities.openPhoneClient = function () { return; }
    //    parent.Mscrm.Shortcuts.openPhoneWindow = function () { return; }
}

function disableformControls() {
    try {
        var FormType = Xrm.Page.ui.getFormType();
        if (FormType == 1) return;
        var SubStatus = Xrm.Page.getAttribute("msdyn_substatus").getValue();
        var SubStatusId = SubStatus[0].id;
        var substatusname = "";
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/msdyn_workordersubstatuses?$select=msdyn_name&$filter=msdyn_workordersubstatusid eq " + SubStatusId + "", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    for (var i = 0; i < results.value.length; i++) {
                        substatusname = results.value[i]["msdyn_name"];
                        debugger;
                        if (substatusname != "Work Done" && substatusname != "Closed") {
                            Xrm.Page.getControl("hil_schemecode").setVisible(false);
                        }
                        if (substatusname == "Closed" || substatusname == "KKG Audit Failed") {
                            Xrm.Page.ui.tabs.get("tab_Claims").setVisible(true);
                        } else {
                            Xrm.Page.ui.tabs.get("tab_Claims").setVisible(false);
                        }
                        if (substatusname == "Canceled" || substatusname == "Closed") {
                            Xrm.Page.data.entity.attributes.forEach(function (attribute, index) {
                                var control = Xrm.Page.getControl(attribute.getName());
                                if (control) {
                                    control.setDisabled(true)
                                }
                            });
                        }
                    }
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function GetAddress() {
    debugger;
    try {
        var Address = Xrm.Page.getAttribute("hil_address").getValue();
        if (Address != null) {
            var AddressId = Address[0].id;
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/hil_addressSet(guid'" + AddressId + "')?$select=hil_addressId,hil_AddressType,hil_Area,hil_Branch,hil_BusinessGeo,hil_CIty,hil_Customer,hil_District,hil_FullAddress,hil_name,hil_PinCode,hil_Region,hil_SalesOffice,hil_State,hil_Street1,hil_Street2,hil_Street3,hil_SubTerritory", true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.responseText).d;
                        var hil_addressId = result.hil_addressId;
                        var hil_AddressType = result.hil_AddressType;
                        var Area = result.hil_Area;
                        var AreaLook = new Array();
                        AreaLook[0] = new Object();
                        AreaLook[0].id = Area.Id;
                        AreaLook[0].name = Area.Name;
                        AreaLook[0].entityType = "hil_area";
                        if (Area.Id != null) {
                            Xrm.Page.getAttribute("hil_area").setValue(AreaLook);
                            Xrm.Page.getAttribute("hil_areatext").setValue(Area.Name);
                        }
                        var Branch = result.hil_Branch;
                        var BranchLook = new Array();
                        BranchLook[0] = new Object();
                        BranchLook[0].id = Branch.Id;
                        BranchLook[0].name = Branch.Name;
                        BranchLook[0].entityType = "hil_branch";
                        if (Branch.Id != null) {
                            Xrm.Page.getAttribute("hil_branch").setValue(BranchLook);
                            Xrm.Page.getAttribute("hil_branchtext").setValue(Branch.Name);
                        }
                        var CIty = result.hil_CIty;
                        var CityLook = new Array();
                        CityLook[0] = new Object();
                        CityLook[0].id = CIty.Id;
                        CityLook[0].name = CIty.Name;
                        CityLook[0].entityType = "hil_city";
                        if (CIty.Id != null) {
                            Xrm.Page.getAttribute("hil_city").setValue(CityLook);
                            Xrm.Page.getAttribute("hil_citytext").setValue(CIty.Name);
                        }
                        var District = result.hil_District;
                        var DistrictLook = new Array();
                        DistrictLook[0] = new Object();
                        DistrictLook[0].id = District.Id;
                        DistrictLook[0].name = District.Name;
                        DistrictLook[0].entityType = "hil_district";
                        if (District.Id != null) {
                            Xrm.Page.getAttribute("hil_district").setValue(DistrictLook);
                            Xrm.Page.getAttribute("hil_districttext").setValue(District.Name);
                        }
                        var PinCode = result.hil_PinCode;
                        var PinCodeLook = new Array();
                        PinCodeLook[0] = new Object();
                        PinCodeLook[0].id = PinCode.Id;
                        PinCodeLook[0].name = PinCode.Name;
                        PinCodeLook[0].entityType = "hil_pincode";
                        if (PinCode.Id != null) {
                            Xrm.Page.getAttribute("hil_pincode").setValue(PinCodeLook);
                            Xrm.Page.getAttribute("hil_pincodetext").setValue(PinCode.Name);
                        }
                        var Region = result.hil_Region;
                        var RegionLook = new Array();
                        RegionLook[0] = new Object();
                        RegionLook[0].id = Region.Id;
                        RegionLook[0].name = Region.Name;
                        RegionLook[0].entityType = "hil_region";
                        if (Region.Id != null) {
                            Xrm.Page.getAttribute("hil_region").setValue(RegionLook);
                            Xrm.Page.getAttribute("hil_regiontext").setValue(Region.Name);
                        }
                        var SalesOffice = result.hil_SalesOffice;
                        var SalesOfficeLook = new Array();
                        SalesOfficeLook[0] = new Object();
                        SalesOfficeLook[0].id = SalesOffice.Id;
                        SalesOfficeLook[0].name = SalesOffice.Name;
                        SalesOfficeLook[0].entityType = "hil_salesoffice";
                        if (SalesOffice.Id != null) {
                            Xrm.Page.getAttribute("hil_salesoffice").setValue(SalesOfficeLook);
                        }
                        else {
                            alert("Sales Office can't be null");
                            Xrm.Page.getAttribute("hil_address").setValue(null);
                        }
                        var State = result.hil_State;
                        var StateLook = new Array();
                        StateLook[0] = new Object();
                        StateLook[0].id = State.Id;
                        StateLook[0].name = State.Name;
                        StateLook[0].entityType = "hil_state";
                        if (State.Id != null) {
                            Xrm.Page.getAttribute("hil_state").setValue(StateLook);
                            Xrm.Page.getAttribute("hil_statetext").setValue(State.Name);
                        }
                        var Street1 = result.hil_Street1;
                        if (Street1 != null) {
                            Xrm.Page.getAttribute("msdyn_address1").setValue(Street1);
                        }
                        var Street2 = result.hil_Street2;
                        if (Street2 != null) {
                            Xrm.Page.getAttribute("msdyn_address2").setValue(Street2);
                        }
                        var Street3 = result.hil_Street3;
                        if (Street3 != null) {
                            Xrm.Page.getAttribute("msdyn_address3").setValue(Street3);
                        }
                        var FullAddress = result.hil_FullAddress;
                        if (FullAddress != null) {
                            Xrm.Page.getAttribute("hil_fulladdress").setValue(FullAddress);
                        }
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
        else {
            Xrm.Page.getAttribute("hil_fulladdress").setValue(null);
            Xrm.Page.getAttribute("msdyn_address3").setValue(null);
            Xrm.Page.getAttribute("msdyn_address2").setValue(null);
            Xrm.Page.getAttribute("msdyn_address1").setValue(null);
            Xrm.Page.getAttribute("hil_state").setValue(null);
            Xrm.Page.getAttribute("hil_region").setValue(null);
            Xrm.Page.getAttribute("hil_pincode").setValue(null);
            Xrm.Page.getAttribute("hil_district").setValue(null);
            Xrm.Page.getAttribute("hil_city").setValue(null);
            Xrm.Page.getAttribute("hil_branch").setValue(null);
            Xrm.Page.getAttribute("hil_pincodetext").setValue(null);
            Xrm.Page.getAttribute("hil_statetext").setValue(null);
            Xrm.Page.getAttribute("hil_branchtext").setValue(null);
            Xrm.Page.getAttribute("hil_regiontext").setValue(null);
            Xrm.Page.getAttribute("hil_districttext").setValue(null);
            Xrm.Page.getAttribute("hil_citytext").setValue(null);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function RbnBtnWOPOCreate() {
    try {
        var Flag = Xrm.Page.getAttribute("hil_flagpo").getValue();
        if (Flag == false) {
            Xrm.Page.getAttribute("hil_flagpo").setValue(true);
            Xrm.Page.data.entity.save();
        }
        else {
            Xrm.Page.getAttribute("hil_flagpo").setValue(false);
            Xrm.Page.data.entity.save();
        }
        //CallAction();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function setDefaultStatus() {
    try {
        var FormType = Xrm.Page.ui.getFormType();
        if (FormType != 1) return;
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/msdyn_workordersubstatuses?$select=msdyn_name,msdyn_systemstatus,msdyn_workordersubstatusid&$filter=msdyn_defaultsubstatus eq true", true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    for (var i = 0; i < results.value.length; i++) {
                        var msdyn_name = results.value[i]["msdyn_name"];
                        var msdyn_systemstatus = results.value[i]["msdyn_systemstatus"];
                        var msdyn_systemstatus_formatted = results.value[i]["msdyn_systemstatus@OData.Community.Display.V1.FormattedValue"];
                        var msdyn_workordersubstatusid = results.value[i]["msdyn_workordersubstatusid"];
                        Xrm.Page.getAttribute("msdyn_systemstatus").setValue(msdyn_systemstatus);
                        var lookupobj = new Array();
                        lookupobj[0] = new Object();
                        lookupobj[0].name = msdyn_name;
                        lookupobj[0].id = msdyn_workordersubstatusid;
                        lookupobj[0].entityType = "msdyn_workordersubstatus";
                        Xrm.Page.getAttribute("msdyn_substatus").setValue(lookupobj);
                        break;
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function CallAction() {
    debugger;
    try {
        var currentRecordIdString = Xrm.Page.data.entity.getId();
        var currentRecordId = currentRecordIdString.replace("{", '').replace("}", '');
        //msdyn_workorders
        Process.callAction("hil_CreatePO", [
            {
                key: "Target",
                type: Process.Type.EntityReference,
                value: new Process.EntityReference("msdyn_workorder", Xrm.Page.data.entity.getId())
            }],

            function (params) {
                // Success
                alert("success");
            },

            function (e, t) {
                // Error
                alert("error");
            });
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function HideShowAMCGrid() {
    try {
        var IfAMCAppl = Xrm.Page.getAttribute("hil_ifamceligible").getValue();
        if (IfAMCAppl == 1) {
            Xrm.Page.getControl("AMCPlans").setVisible(true);
        }
        else {
            Xrm.Page.getControl("AMCPlans").setVisible(false);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function AMCPlanspreFilterLookup() {
    debugger;
    try {
        Xrm.Page.getControl("hil_amcplan").addPreSearch(function () {
            addLookupFilter();
        });
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function addLookupFilter() {
    try {
        var co = Xrm.Page.getAttribute("hil_productcategory").getValue();
        var coId;
        if (co != null) {
            coId = co[0].id;
        }
        var fetchXml = '<filter type="and"><condition attribute="hil_materialgroup" value="' + coId + '" operator="eq" /></filter>';
        Xrm.Page.getControl("hil_amcplan").addCustomFilter(fetchXml);
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function PrefetchLookupInWorkOrder() {
    try {
        var Calltype = Xrm.Page.getAttribute("hil_calltype").getValue();
        var PdtDiv = Xrm.Page.getAttribute("hil_productcategory").getValue();
        var CallTypeName;
        var CallTypeId;
        var PdtDivName;
        var PdtDivId;
        if (Calltype != null) {
            CallTypeId = Calltype[0].id;
            CallTypeName = Calltype[0].name;
        }
        if (PdtDiv != null) {
            PdtDivId = PdtDiv[0].id;
            PdtDivName = PdtDiv[0].name;
        }
        //	alert(CallTypeId);
        //alert(PdtDivId);
        var fetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
            "<entity name='systemuser'>" +
            "<attribute name='fullname' />" +
            "<attribute name='businessunitid' />" +
            "<attribute name='title' />" +
            "<attribute name='address1_telephone1' />" +
            "<attribute name='positionid' />" +
            "<attribute name='systemuserid' />" +
            "<attribute name='internalemailaddress' />" +
            "<order attribute='fullname' descending='false' />" +
            "<link-entity name='characteristic' from='hil_assignto' to='systemuserid' link-type='inner' alias='ag'>" +
            "<filter type='and'>" +
            "<condition attribute='hil_calltype' operator='eq' uiname='" + CallTypeName + "' uitype='hil_callsubtype' value='{" + CallTypeId + "}'/>" +
            "<condition attribute='hil_productdivision' operator='eq' uiname='" + PdtDivName + "' uitype='product' value='{" + PdtDivId + "}' />" +
            "</filter>" +
            "</link-entity>" +
            "</entity>" +
            "</fetch>";
        var layoutXml = "<grid name='resultset' " +
            "object='1' " +
            "jump='name' " +
            "select='1' " +
            "icon='1' " +
            "preview='1'>" +
            "<row name='result' " +
            "id='systemuserid'>" +
            "<cell name='fullname' " +
            "width='300' />" +
            "<cell name='title' " +
            "width='150' />" +
            "<cell name='address1_telephone1' " +
            "width='100' />" +
            "<cell name='positionid' " +
            "width='100' />" +
            "disableSorting='1' />" +
            "</row>" +
            "</grid>";
        var viewId = "{C7034F4F-6F92-4DD7-BD9D-9B9C1E996380}";
        var entityName = "systemuser";
        var viewDisplayName = "Custom Lookup View";
        Xrm.Page.getControl(ownerid).addCustomView(viewId, entityName, viewDisplayName, fetchXml, layoutXml, true);
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function SetPriceList() {
    try {
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/pricelevels?$select=name,pricelevelid&$filter=name eq 'Default%20Price%20List'", true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    for (var i = 0; i < results.value.length; i++) {
                        var name = results.value[i]["name"];
                        var pricelevelid = results.value[i]["pricelevelid"];
                        if (pricelevelid != null) {
                            var SetCat = new Array();
                            SetCat[0] = new Object();
                            SetCat[0].name = name;
                            SetCat[0].id = pricelevelid;
                            SetCat[0].entityType = "pricelevel";
                            Xrm.Page.getAttribute("msdyn_pricelist").setValue(SetCat);
                        }
                        else {
                            alert("Can't find Default Price List");
                        }
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function SetCatSubCat() {
    debugger;
    try {
        var CatSubCatMap = Xrm.Page.getAttribute("hil_productcatsubcatmapping").getValue();
        if (CatSubCatMap == null) {
            Xrm.Page.getAttribute("hil_productcategory").setValue(null);
            Xrm.Page.getAttribute("hil_productsubcategory").setValue(null);
            Xrm.Page.getAttribute("hil_quantity").setValue(null);
            Xrm.Page.getAttribute("hil_maxquantity").setValue(null);
            Xrm.Page.getAttribute("hil_minquantity").setValue(null);
        }
        else if (CatSubCatMap != null) {
            CatSubCatMap = CatSubCatMap[0].id;
            CatSubCatMap = CatSubCatMap.replace("{", "").replace("}", "");
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/hil_stagingdivisonmaterialgroupmappings?$select=_hil_productcategorydivision_value,_hil_productsubcategorymg_value&$filter=hil_stagingdivisonmaterialgroupmappingid eq " + CatSubCatMap + "", true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var results = JSON.parse(this.response);
                        for (var i = 0; i < results.value.length; i++) {
                            var _hil_productcategorydivision_value = results.value[i]["_hil_productcategorydivision_value"];
                            var _hil_productcategorydivision_value_formatted = results.value[i]["_hil_productcategorydivision_value@OData.Community.Display.V1.FormattedValue"];
                            var _hil_productcategorydivision_value_lookuplogicalname = results.value[i]["_hil_productcategorydivision_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
                            var _hil_productsubcategorymg_value = results.value[i]["_hil_productsubcategorymg_value"];
                            var _hil_productsubcategorymg_value_formatted = results.value[i]["_hil_productsubcategorymg_value@OData.Community.Display.V1.FormattedValue"];
                            var _hil_productsubcategorymg_value_lookuplogicalname = results.value[i]["_hil_productsubcategorymg_value@Microsoft.Dynamics.CRM.lookuplogicalname"];
                            if (_hil_productcategorydivision_value != null) {
                                var SetCat = new Array();
                                SetCat[0] = new Object();
                                SetCat[0].name = _hil_productcategorydivision_value_formatted;
                                SetCat[0].id = _hil_productcategorydivision_value;
                                SetCat[0].entityType = _hil_productcategorydivision_value_lookuplogicalname;
                                Xrm.Page.getAttribute("hil_productcategory").setValue(SetCat);
                                //  Xrm.Page.getAttribute("hil_productsubcategory").setValue(SetSubCat);
                            }
                            if (_hil_productsubcategorymg_value != null) {
                                var SetSubCat = new Array();
                                SetSubCat[0] = new Object();
                                SetSubCat[0].name = _hil_productsubcategorymg_value_formatted;
                                SetSubCat[0].id = _hil_productsubcategorymg_value;
                                SetSubCat[0].entityType = _hil_productsubcategorymg_value_lookuplogicalname;
                                Xrm.Page.getAttribute("hil_productsubcategory").setValue(SetSubCat);
                                GetMinMaxThreshold(_hil_productsubcategorymg_value);
                            }
                            //alert("Working");
                        }
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function test() {
    try {
        var LoggedInUser = Xrm.Page.context.getUserRoles();
        //alert(LoggedInUser);
        if (LoggedInUser == "212803cb-0c76-e811-a95c-000d3af06236") {
            var OwnAcId;
            var OwnAc = Xrm.Page.getAttribute("hil_owneraccount").getValue();
            if (OwnAc != null) {
                OwnAcId = OwnAc[0].id;
                var req = new XMLHttpRequest();
                req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/AccountSet(guid'" + OwnAcId + "')?$select=CustomerTypeCode", true);
                req.setRequestHeader("Accept", "application/json");
                req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                req.onreadystatechange = function () {
                    if (this.readyState === 4) {
                        this.onreadystatechange = null;
                        if (this.status === 200) {
                            var result = JSON.parse(this.responseText).d;
                            var customerTypeCode = result.CustomerTypeCode;
                            var CustomerCode = customerTypeCode.Value;
                            if (CustomerCode != 6) {
                                PreFilterOwnerField();
                            }
                            //alert(customerTypeCode);
                            //alert(CustomerCode);
                        }
                        else {
                            Xrm.Utility.alertDialog(this.statusText);
                        }
                    }
                };
                req.send();
                //break;
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function makeAssociatedPdtMandatory() {
    try {
        var CallSubType = Xrm.Page.getAttribute("hil_callsubtype").getValue();
        if (CallSubType != null) {
            var CallSTypeName = CallSubType[0].name;
            if (CallSTypeName == "Product Registration") {
                Xrm.Page.getAttribute("msdyn_customerasset").setRequiredLevel("required");
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function setDummyAccount() {
    try {
        var CallType = Xrm.Page.getAttribute("hil_callsubtype").getValue();
        if (CallType != null) {
            var CalltypeName = CallType[0].name;
            if (CalltypeName == "PMS") {
                if (Xrm.Page.ui.tabs.get("tab_15") != null) {
                    Xrm.Page.ui.tabs.get("tab_15").setVisible(true);
                }
            }
            else {
                if (Xrm.Page.ui.tabs.get("tab_15") != null) {
                    Xrm.Page.ui.tabs.get("tab_15").setVisible(false);
                }
            }
        }
        else {
            if (Xrm.Page.ui.tabs.get("tab_15") != null) {
                Xrm.Page.ui.tabs.get("tab_15").setVisible(false);
            }
        }
        var FormType = Xrm.Page.ui.getFormType();
        if (FormType == 1) {
            var SetLookup = new Array();
            var name = "Dummy Customer";
            var result = GetDummyCustomer(name);
            if (result != null) {
                SetLookup[0] = new Object();
                SetLookup[0].name = name;
                SetLookup[0].id = result;
                SetLookup[0].entityType = "account";
                Xrm.Page.getAttribute("msdyn_serviceaccount").setValue(SetLookup);
                Xrm.Page.getAttribute("hil_sourceofjob").setValue(2);
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function GetDummyCustomer(name) {
    try {
        var ClientUrI = Xrm.Page.context.getClientUrl();
        var record = ClientUrI + "/XRMServices/2011/OrganizationData.svc/AccountSet?$select=AccountId&$filter=Name eq '" + name + "'";
        var AccountGuid = null;
        var retrieveReq = new XMLHttpRequest();
        retrieveReq.open("GET", record, false);
        retrieveReq.setRequestHeader("Accept", "application/json");
        retrieveReq.setRequestHeader("Content-Type", "application/json;charset=utf-8");
        retrieveReq.onreadystatechange = function () {
            if (retrieveReq.readyState == 4) {
                if (retrieveReq.status == 200) {
                    var retrieved = JSON.parse(retrieveReq.responseText).d;
                }
                AccountGuid = retrieved.results[0].AccountId;
            }
        };
        retrieveReq.send();
        return (AccountGuid);
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function PreFilterOwnerField() {
    try {
        var CallType = Xrm.Page.getAttribute("hil_callsubtype").getValue();
        var PdtCtgry = Xrm.Page.getAttribute("hil_productcategory").getValue();
        var Pincode = Xrm.Page.getAttribute("hil_pincode").getValue();
        if ((CallType != null) && (PdtCtgry != null) && (Pincode != null)) {
            var PinCodeId = null;
            var CallSubTypeId = null;
            var PdtCtgryId = null;
            var PinCodeName = null;
            var CallSubtypeName = null;
            var PdtCtgryName = null;
            if (CallType != null) {
                CallSubTypeId = CallType[0].id;
                CallSubtypeName = CallType[0].name;
                //alert(CallSubTypeId);
            }
            if (PdtCtgry != null) {
                DivisionId = PdtCtgry[0].id;
                DivisionName = PdtCtgry[0].name;
                //alert(DivisionId);
            }
            if (Pincode != null) {
                PinCodeId = Pincode[0].id;
                PinCodeName = Pincode[0].name;
                //alert(PinCodeId);
            }
            var fetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
                "<entity name='systemuser'>" +
                "<attribute name='fullname' />" +
                "<attribute name='systemuserid' />" +
                "<order attribute='fullname' descending='false' />" +
                "<link-entity name='hil_assignmentmatrix' from='hil_assignto' to='systemuserid' link-type='inner' alias='af'>" +
                "<filter type='and'>" +
                "<condition attribute='hil_pincode' operator='eq' uiname='" + PinCodeName + "' uitype='hil_pincode' value='" + PinCodeId + "' />" +
                "<condition attribute='hil_callsubtype' operator='eq' uiname='" + CallSubtypeName + "' uitype='hil_callsubtype' value='" + CallSubTypeId + "' />" +
                "<condition attribute='hil_division' operator='eq' uiname='" + DivisionName + "' uitype='product' value='" + DivisionId + "' />" +
                "</filter>" +
                "</link-entity>" +
                "</entity>" +
                "</fetch>";
            var layoutXml = "<grid name='resultset' " +
                "object='8' " +
                "jump='name' " +
                "select='1' " +
                "icon='1' " +
                "preview='1'>" +
                "<row name='result' " +
                "id='systemuserid'>" +
                "<cell name='fullname' " +
                "width='300' />" +
                "disableSorting='1' />" +
                "</row>" +
                "</grid>";
            var viewId = "{C7034F4F-6F92-4DD7-BD9D-9B9C1E996380}";
            var entityName = "systemuser";
            var viewDisplayName = "TestThisView";
            Xrm.Page.getControl("ownerid").addCustomView(viewId, entityName, viewDisplayName, fetchXml, layoutXml, true);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function defaultcustomer() {
    try {
        Xrm.Page.getControl("hil_customerref").addPreSearch(addFilter);
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function addFilter() {
    try {
        var customerAccountFilter = "<filter type='and'><condition attribute='accountid' operator='null' /></filter>";
        Xrm.Page.getControl("hil_customerref").addCustomFilter(customerAccountFilter, "account");
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function WoOnSave() {
    try {
        var RecType = Xrm.Page.ui.getFormType();
        if (RecType == 1) {
            var Assign = Xrm.Page.getAttribute("hil_automaticassign").setValue(1);
            Xrm.Page.data.entity.save();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function ContactOnChange() {
    debugger;
    try {
        var Customer = Xrm.Page.getAttribute("hil_customerref").getValue();
        var CustomerName;
        var CustomerId;
        var EntityName;
        if (Customer != null) {
            EntityName = Customer[0].entityType;
            //alert(EntityName);
            if (EntityName == "contact") {
                CustomerId = Customer[0].id;
                //alert(CustomerId );
                CustomerName = Customer[0].name;
                GetDataFromcontact(CustomerId);
            }
        }
        else {
            //Xrm.Page.getAttribute("hil_mobilenumber").setValue(null);
            Xrm.Page.getAttribute("hil_callingnumber").setValue(null);
            Xrm.Page.getAttribute("hil_alternate").setValue(null);
            Xrm.Page.getAttribute("hil_customername").setValue(null);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function GetDataFromcontact(ContactId) {
    try {
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/contacts?$select=address1_telephone2,address1_telephone3,fullname,mobilephone,emailaddress1&$filter=contactid eq " + ContactId + "", true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    for (var i = 0; i < results.value.length; i++) {
                        var address1_telephone2 = results.value[i]["address1_telephone2"];
                        var address1_telephone3 = results.value[i]["address1_telephone3"];
                        var fullname = results.value[i]["fullname"];
                        var mobilephone = results.value[i]["mobilephone"];
                        var Email = results.value[i]["emailaddress1"];
                        Xrm.Page.getAttribute("hil_email").setValue(Email);
                        if (Xrm.Page.ui.getFormType() == 1) //Mobile number shouldn't change if customer is changing
                        {
                            Xrm.Page.getAttribute("hil_mobilenumber").setValue(mobilephone);
                        }
                        //Xrm.Page.getAttribute("hil_callingnumber").setValue(address1_telephone2);
                        Xrm.Page.getAttribute("hil_alternate").setValue(address1_telephone3);
                        Xrm.Page.getAttribute("hil_customername").setValue(fullname);
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}
// on change of product sub category and call sub type
var City;
var EmployeeCode;

function OnChange() {
    debugger;
    try {
        var ProductSubCategory;
        var callSubType;
        var ProductSubCategoryCallSubType;
        var Region;
        var EmployeeCodeName
        var RegionBranch;
        var Owner;
        var Branch;
        var Customer;
        var OwnerBranchCityEC;
        if (Xrm.Page.getAttribute("hil_productsubcategory").getValue() != undefined && Xrm.Page.getAttribute("hil_productsubcategory").getValue() != null) {
            ProductSubCategory = Xrm.Page.getAttribute("hil_productsubcategory").getValue()[0].id;
            ProductSubCategory = ProductSubCategory.replace("{", "").replace("}", "");
        }
        if (Xrm.Page.getAttribute("hil_callsubtype").getValue() != undefined && Xrm.Page.getAttribute("hil_callsubtype").getValue() != null) {
            callSubType = Xrm.Page.getAttribute("hil_callsubtype").getValue()[0].id;
            callSubType = callSubType.replace("{", "").replace("}", "");
        }
        if (ProductSubCategory != null && callSubType != null) ProductSubCategoryCallSubType = ProductSubCategory + ":" + callSubType;
        if (Xrm.Page.getAttribute("hil_productsubcategorycallsubtype") != undefined) Xrm.Page.getAttribute("hil_productsubcategorycallsubtype").setValue(ProductSubCategoryCallSubType);
        //  on update of Region/Branch 
        if (Xrm.Page.getAttribute("hil_region").getValue() != undefined && Xrm.Page.getAttribute("hil_region").getValue() != null) {
            Region = Xrm.Page.getAttribute("hil_region").getValue()[0].id;
            Region = Region.replace("{", "").replace("}", "");
        }
        if (Xrm.Page.getAttribute("hil_branch").getValue() != undefined && Xrm.Page.getAttribute("hil_branch").getValue() != null) {
            Branch = Xrm.Page.getAttribute("hil_branch").getValue()[0].id;
            Branch = Branch.replace("{", "").replace("}", "");
        }
        if (Region != null && Branch != null) RegionBranch = Region + ":" + Branch;
        if (Xrm.Page.getAttribute("hil_regionbranch") != undefined) Xrm.Page.getAttribute("hil_regionbranch").setValue(RegionBranch);
        //on update of Branch/Engineer/City
        if (Xrm.Page.getAttribute("ownerid").getValue() != undefined && Xrm.Page.getAttribute("ownerid").getValue() != null) {
            Owner = Xrm.Page.getAttribute("ownerid").getValue()[0].id;
            Owner = Owner.replace("{", "").replace("}", "");
        } //hil_customerref
        if (Xrm.Page.getAttribute("hil_customerref").getValue() != undefined && Xrm.Page.getAttribute("hil_customerref").getValue() != null) {
            Customer = Xrm.Page.getAttribute("hil_customerref").getValue()[0].id;
            // Owner = Owner.replace("{", "").replace("}", "");
        }
        debugger;
        ContactOnChangeCity();
        GetEmployeeCode();
        OwnerBranchCityEC = Owner + ":" + Branch + ":" + City + ":" + EmployeeCode;
        Xrm.Page.getAttribute("hil_branchengineercity").setValue(OwnerBranchCityEC);
        // update employee code and name
        EmployeeCodeName = Owner + ":" + EmployeeCode;
        Xrm.Page.getAttribute("hil_emolpyeenamecode").setValue(EmployeeCodeName);
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function GetEmployeeCode() {
    try {
        var Owner = Xrm.Page.getAttribute("ownerid").getValue();
        var URL = Xrm.Page.context.getClientUrl() + "/api/data/v8.2/systemusers(guid'" + OwnerId + "')?$select=hil_employeecode";
        if (Owner != null) {
            var OwnerId = Owner[0].id;
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/SystemUserSet(guid'" + OwnerId + "')?$select=hil_EmployeeCode", true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.responseText).d;
                        var hil_EmployeeCode = result.hil_EmployeeCode;
                        //alert(hil_EmployeeCode);
                        EmployeeCode = hil_EmployeeCode;
                        //alert(hil_EmployeeCode);
                        // Xrm.Page.getAttribute("hil_branchengineercity").setValue(hil_EmployeeCode);
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function ContactOnChangeCity() {
    try {
        var Contact = Xrm.Page.getAttribute("hil_customerref").getValue();
        if (Contact != null) {
            var ContactId = Contact[0].id;
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/ContactSet(guid'" + ContactId + "')?$select=hil_city", true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.responseText).d;
                        var hil_city = result.hil_city;
                        var SetLookup = new Array();
                        SetLookup[0] = new Object();
                        SetLookup[0].name = hil_city.Name;
                        SetLookup[0].id = hil_city.Id;
                        City = hil_city.Id;
                        //alert(hil_city.Id)
                        SetLookup[0].entityType = "hil_city";
                        //Xrm.Page.getAttribute("hil_branchengineercity").setValue(hil_city.Id);
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function GetCallSubType() {
    debugger;
    try {
        var CallSTyp = Xrm.Page.getAttribute("hil_callsubtype").getValue();
        var CallSTypeID;
        var CallSTypeName;
        if (CallSTyp != null) {
            CallSTypeID = CallSTyp[0].id;
            CallSTypeName = CallSTyp[0].name;
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/hil_callsubtypeSet(guid'" + CallSTypeID + "')?$select=hil_CallType,hil_cause,hil_natureofcomplaint,hil_observation", true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.responseText).d;
                        var hil_CallType = result.hil_CallType;
                        var SetLookupCallTyp = new Array();
                        SetLookupCallTyp[0] = new Object();
                        SetLookupCallTyp[0].name = hil_CallType.Name;
                        SetLookupCallTyp[0].id = hil_CallType.Id;
                        SetLookupCallTyp[0].entityType = "hil_calltype";
                        Xrm.Page.getAttribute("hil_calltype").setValue(SetLookupCallTyp);
                        var hil_cause = result.hil_cause;
                        var SetLookupCause = new Array();
                        SetLookupCause[0] = new Object();
                        SetLookupCause[0].name = hil_cause.Name;
                        SetLookupCause[0].id = hil_cause.Id;
                        SetLookupCause[0].entityType = "msdyn_incidenttype";
                        Xrm.Page.getAttribute("msdyn_primaryincidenttype").setValue(SetLookupCause);
                        if (CallSTypeName == "Pre Sales Demo" || CallSTypeName == "Demo" || CallSTypeName == "Serial Number Change" || CallSTypeName == "Sales Call") {
                            var hil_natureofcomplaint = result.hil_natureofcomplaint;
                            var SetLookupNature = new Array();
                            SetLookupNature[0] = new Object();
                            SetLookupNature[0].name = hil_natureofcomplaint.Name;
                            SetLookupNature[0].id = hil_natureofcomplaint.Id;
                            SetLookupNature[0].entityType = "hil_natureofcomplaint";
                            Xrm.Page.getAttribute("hil_natureofcomplaint").setValue(SetLookupNature);
                        }
                        var hil_observation = result.hil_observation;
                        var SetLookupObservation = new Array();
                        SetLookupObservation[0] = new Object();
                        SetLookupObservation[0].name = hil_observation.Name;
                        SetLookupObservation[0].id = hil_observation.Id;
                        SetLookupObservation[0].entityType = "hil_observation";
                        Xrm.Page.getAttribute("hil_observation").setValue(SetLookupObservation);
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
        else {
            Xrm.Page.getAttribute("hil_observation").setValue(null);
            Xrm.Page.getAttribute("hil_natureofcomplaint").setValue(null);
            Xrm.Page.getAttribute("msdyn_primaryincidenttype").setValue(null);
            Xrm.Page.getAttribute("hil_calltype").setValue(null);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function GetWarrantyStatus() {
    try {
        var ProdSC = Xrm.Page.getAttribute("hil_productsubcategory").getValue();
        var PDate = Xrm.Page.getAttribute("hil_purchasedate").getValue();
        if (PDate != null) {
            var DateNow = new Date();
            var NowYear = DateNow.getFullYear();
            var NowMonth = DateNow.getMonth() + 1;
            var PYear = PDate.getFullYear();
            var PMonth = PDate.getMonth() + 1;
            var MonDiff = (NowYear - PYear) * 12 + NowMonth - PMonth;
            if (ProdSC != null) {
                var ProdId = ProdSC[0].id;
                ProdId = ProdId.replace("{", "");
                ProdId = ProdId.replace("}", "");
                var req = new XMLHttpRequest();
                req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/hil_warrantytemplateSet?$select=hil_Product,hil_type,hil_WarrantyPeriod&$filter=hil_Product/Id eq (guid'" + ProdId + "') and hil_type/Value eq 1", true);
                req.setRequestHeader("Accept", "application/json");
                req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                req.onreadystatechange = function () {
                    if (this.readyState === 4) {
                        this.onreadystatechange = null;
                        if (this.status === 200) {
                            var returned = JSON.parse(this.responseText).d;
                            var results = returned.results;
                            for (var i = 0; i < results.length; i++) {
                                var type = results[i].hil_type;
                                var typeValue = type.Value;
                                var WarrantyPeriod = results[i].hil_WarrantyPeriod;
                                if (MonDiff < WarrantyPeriod) {
                                    alert("IN WARRANTY");
                                    Xrm.Page.getAttribute("hil_warrantystatus").setValue(1);
                                }
                                else {
                                    alert("OUT WARRANTY\nStandard Visit Charges may be applicable");
                                    Xrm.Page.getAttribute("hil_warrantystatus").setValue(2);
                                }
                            }
                        }
                        else {
                            Xrm.Utility.alertDialog(this.statusText);
                        }
                    }
                };
                req.send();
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function RestrictFutureDate() {
    try {
        var PDate = Xrm.Page.getAttribute("hil_purchasedate").getValue();
        var SysDate = new Date();
        if (PDate != null) {
            if (PDate > SysDate) {
                alert("Future Date not Allowed");
                Xrm.Page.getAttribute("hil_purchasedate").setValue(null);
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OpenNewTask() {
    try {
        var ID = Xrm.Page.data.entity.getId();
        var Jobname = Xrm.Page.data.entity.attributes.get("msdyn_name").getValue();
        ID = ID.replace("{", "");
        ID = ID.replace("}", "");
        var formParameters = {};
        formParameters["hil_jobguid"] = ID;
        formParameters["hil_jobname"] = Jobname;
        Xrm.Utility.openQuickCreate("task", null, formParameters);
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function RestrictPastDate(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        var PDate = formContext.getAttribute("hil_preferreddate").getValue();
        var year = PDate.getFullYear();
        var month = PDate.getMonth() + 1;
        if (month < 10) {
            month = '0' + month;
        }
        var day = PDate.getDate();
        if (day < 10) {
            day = '0' + day;
        }
        var dateOnly = new Date(year, month, day);
        var dateOnly1 = month + '/' + day + '/' + year;
        var today = new Date();
        var dd = today.getDate();
        var mm = today.getMonth() + 1; //January is 0!
        var yyyy = today.getFullYear();
        if (PDate != null) {
            if (dd < 10) {
                dd = '0' + dd
            }
            if (mm < 10) {
                mm = '0' + mm
            }
            today = mm + '/' + dd + '/' + yyyy;
            //alert(dateOnly1);
            //alert(today);
            if (PDate != null) {
                if (dateOnly1 < today) {
                    alert("Past Date not Allowed");
                    formContext.getAttribute("hil_preferreddate").setValue(null);
                }
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function CreateTask() {
    debugger;
    try {
        var WoName = Xrm.Page.getAttribute("msdyn_name").getValue();
        // alert(WoName);
        var WoId = Xrm.Page.data.entity.getId();
        WoId = WoId.replace("{", "").replace("}", "");
        var entityName = Xrm.Page.data.entity.getEntityName();
        // alert(WoId);
        //alert(entityName);
        var entity = {};
        entity.RegardingObjectId = {
            Id: WoId,
            LogicalName: "msdyn_workorder"
        };
        var req = new XMLHttpRequest();
        req.open("POST", encodeURI(Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/TaskSet"), false);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;
                if (this.status === 201) {
                    var result = JSON.parse(this.responseText).d;
                    var newEntityId = result.TaskId;
                    var windowOptions = {
                        openInNewWindow: true
                    };
                    //Xrm.Utility.openEntityForm("task", newEntityId,null,windowOptions);
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send(JSON.stringify(entity));
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OpenLastTask(id) {
    debugger;
    try {
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/TaskSet?$filter=RegardingObjectId/Id eq (guid'" + workorderid + "')&$top=1&$orderby=CreatedOn desc", true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;
                if (this.status === 200) {
                    var returned = JSON.parse(this.responseText).d;
                    var results = returned.results;
                    var ActivityId = results[i].ActivityId;
                    var windowOptions = {
                        openInNewWindow: true
                    };
                    Xrm.Utility.openEntityForm("task", ActivityId, null, windowOptions);
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OpenLastTask1(id) {
    debugger;
    try {
        var req = new XMLHttpRequest();
        // req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/tasks?$filter=_regardingobjectid_value eq "+workorderid+"&$orderby=createdon desc", false);
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/tasks?$filter=_regardingobjectid_value eq " + id + "&$orderby=createdon desc", false);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\",odata.maxpagesize=1");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    for (var i = 0; i < results.value.length; i++) {
                        var activityid = results.value[i]["activityid"];
                        var windowOptions = {
                            openInNewWindow: true
                        };
                        Xrm.Utility.openEntityForm("task", activityid, null, windowOptions);
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function HideRelatedFranchisee() {
    try {
        var TypeOfAssignee = Xrm.Page.getAttribute("hil_typeofassignee").getValue();
        if (TypeOfAssignee != null) {
            var TypeName = TypeOfAssignee[0].name;
            if (TypeName == "Franchise" || TypeName == "Technician") {
                Xrm.Page.getControl("hil_owneraccount").setVisible(true);
            }
            else {
                Xrm.Page.getControl("hil_owneraccount").setVisible(false);
            }
        }
        else {
            Xrm.Page.getControl("hil_owneraccount").setVisible(false);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function GetMinMaxThreshold(_hil_productcategorydivision_value) {
    //debugger;
    try {
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/ProductSet(guid'" + _hil_productcategorydivision_value + "')?$select=hil_MaximumThreshold,hil_MinimumThreshold", true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;
                if (this.status === 200) {
                    var result = JSON.parse(this.responseText).d;
                    var MaximumThreshold = result.hil_MaximumThreshold;
                    var MinimumThreshold = result.hil_MinimumThreshold;
                    Xrm.Page.getAttribute("hil_quantity").setValue(MinimumThreshold);
                    Xrm.Page.getAttribute("hil_maxquantity").setValue(MaximumThreshold);
                    Xrm.Page.getAttribute("hil_minquantity").setValue(MinimumThreshold);
                    if (MaximumThreshold != null && MinimumThreshold != null) {
                        if (MaximumThreshold == MinimumThreshold) {
                            Xrm.Page.getControl("hil_quantity").setDisabled(true);
                        }
                        else {
                            Xrm.Page.getControl("hil_quantity").setDisabled(false);
                        }
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function QuantityOnChange() {
    try {
        var MaxQuant = Xrm.Page.getAttribute("hil_maxquantity").getValue();
        var MinQuant = Xrm.Page.getAttribute("hil_minquantity").getValue();
        var Quant = Xrm.Page.getAttribute("hil_quantity").getValue();
        if (Quant > MaxQuant || Quant < MinQuant) {
            alert("Quantity can't be more than Maximum threshold quantity " + MaxQuant);
            Xrm.Page.getAttribute("hil_quantity").setValue(null);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function PreFilterCustomerAsset() {
    debugger;
    try {
        Xrm.Page.getControl("msdyn_customerasset").addPreSearch(function () {
            var CustomerRef = Xrm.Page.getAttribute("hil_customerref").getValue();
            var ProductCategory = Xrm.Page.getAttribute("hil_productcategory").getValue();
            var filterStr = "<filter type='and'>";
            var LookupId = null;
            if (CustomerRef != null) {
                LookupId = CustomerRef[0].id;
                filterStr = filterStr + "<condition attribute='hil_customer' operator='eq'  value='" + LookupId + "' />"
            }
            if (ProductCategory != null) {
                LookupId = ProductCategory[0].id;
                filterStr = filterStr + "<condition attribute='hil_productcategory' operator='eq'  value='" + LookupId + "' />"
            }
            if (LookupId == null) {
                filterStr = filterStr + "<condition attribute='hil_customer' operator='eq'  value='{00000000-0000-0000-0000-000000000000}' />"
            }
            var filterStr = filterStr + "</filter>";
            Xrm.Page.getControl("msdyn_customerasset").addCustomFilter(filterStr);
        });
        PreFilterJobReference();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}


filterStr = filterStr + "<link-entity name='hil_servicebom' from='hil_product' to='productid' link-type='inner' alias='ad'>";
filterStr = filterStr + "<filter type='and'>";
filterStr = filterStr + "<condition attribute='hil_productsubcategory' operator='eq' value='" + prodId + "' />";
filterStr = filterStr + "</filter>";
filterStr = filterStr + "</link-entity>";

function OpenNewAppointment() {
    try {
        var JobId = Xrm.Page.data.entity.getId();
        JobId = JobId.replace("{", "");
        JobId = JobId.replace("}", "");
        var PhoneNum = Xrm.Page.getAttribute("hil_mobilenumber").getValue();
        var CallingNum = Xrm.Page.getAttribute("hil_callingnumber").getValue();
        var AlternateNum = Xrm.Page.getAttribute("hil_alternate").getValue();
        CreateAppointment(JobId, PhoneNum, CallingNum, AlternateNum);
        OpenLastAppointment(JobId);
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function CreateAppointment(JobId, PhoneNum, CallingNum, AlternateNum) {
    try {
        var entity = {};
        entity.hil_AlternateNumber = AlternateNum;
        entity.hil_CallingNumber = CallingNum;
        entity.DirectionCode = true;
        entity.PhoneNumber = PhoneNum;
        entity.RegardingObjectId = {
            Id: JobId,
            LogicalName: "msdyn_workorder"
        };
        var req = new XMLHttpRequest();
        req.open("POST", encodeURI(Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/PhoneCallSet"), true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;
                if (this.status === 201) {
                    var result = JSON.parse(this.responseText).d;
                    var newEntityId = result.PhoneCallId;
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send(JSON.stringify(entity));
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OpenLastAppointment(JobId) {
    try {
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/phonecalls?$filter=_regardingobjectid_value eq " + JobId + "&$orderby=createdon desc", true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200) {
                    var results = JSON.parse(this.response);
                    for (var i = 0; i < results.value.length; i++) {
                        var activityid = results.value[i]["activityid"];
                        var windowOptions = {
                            openInNewWindow: true
                        };
                        Xrm.Utility.openEntityForm("phonecall", activityid, null, windowOptions);
                        break;
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function CheckEstimatedPartOrService() {
    try {
        var Calc = Xrm.Page.getAttribute("hil_calculatecharges").getValue();
        if (Calc == true) {
            var WorkOrderId = Xrm.Page.data.entity.getId();
            var Present = true;
            Present = IfEstimatedPart(WorkOrderId);
            if (Present == false) {
                Present = IfEstimatedService(WorkOrderId);
                if (Present == true) {
                    alert("One or More Job Service in Estimated state. Kindly Mark Use");
                    Xrm.Page.getAttribute("hil_calculatecharges").setValue(false);
                }
            }
            else {
                alert("One or More Job Product in Estimated state. Kindly Mark Use");
                Xrm.Page.getAttribute("hil_calculatecharges").setValue(false);
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function IfEstimatedPart(WorkOrderId) {
    try {
        var Present = false;
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/msdyn_workorderproductSet?$select=msdyn_Product,msdyn_WorkOrder&$filter=msdyn_LineStatus/Value eq 690970000 and msdyn_WorkOrder/Id eq (guid'" + WorkOrderId + "')", true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;
                if (this.status === 200) {
                    var returned = JSON.parse(this.responseText).d;
                    var results = returned.results;
                    for (var i = 0; i < results.length; i++) {
                        Present = true;
                        break;
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
        return Present;
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function IfEstimatedService(WorkOrderId) {
    try {
        var Present = false;
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/msdyn_workorderserviceSet?$select=msdyn_Service,msdyn_WorkOrder,msdyn_WorkOrderIncident&$filter=msdyn_LineStatus/Value eq 690970000 and msdyn_WorkOrder/Id eq (guid'" + WorkOrderId + "')", true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;
                if (this.status === 200) {
                    var returned = JSON.parse(this.responseText).d;
                    var results = returned.results;
                    for (var i = 0; i < results.length; i++) {
                        Present = true;
                        break;
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
        return Present;
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function CheckUserRole() {
    try {
        var currentUserRoles = Xrm.Page.context.getUserRoles();
        for (var i = 0; i < currentUserRoles.length; i++) {
            var roleId = currentUserRoles[i];
            if (roleId == "4EF33651-5D88-E811-A960-000D3AF049DF") {
                Xrm.Page.ui.tabs.get("Admin").setVisible(true);
                break;
            }
            else {
                Xrm.Page.ui.tabs.get("Admin").setVisible(false);
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function RestrictAssign() {
    try {
        if (Xrm.Page.getAttribute("hil_appointmentstatus") != null) {
            var ApStatus = Xrm.Page.getAttribute("hil_appointmentstatus").getValue();
            if (ApStatus == null || ApStatus != 1) {
                var currentUserRoles = Xrm.Page.context.getUserRoles();
                for (var i = 0; i < currentUserRoles.length; i++) {
                    var roleId = currentUserRoles[i];
                    IfFranchisee(roleId);
                }
            }
            else {
                Xrm.Page.getControl("ownerid").setDisabled(false);
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function IfFranchisee(roleId) {
    try {
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/RoleSet?$select=Name&$filter=RoleId eq (guid'" + roleId + "')&$top=50", true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;
                if (this.status === 200) {
                    var returned = JSON.parse(this.responseText).d;
                    var results = returned.results;
                    for (var i = 0; i < results.length; i++) {
                        var name = results[i].Name;
                        if (name == "Franchise") {
                            Xrm.Page.getControl("ownerid").setDisabled(true);
                            Xrm.Page.ui.setFormNotification("Kindly Add Appointment Before Assigning the Job to Technician ", "WARNING")
                            break;
                        }
                        else {
                            Xrm.Page.getControl("ownerid").setDisabled(false);
                        }
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OpenAddressForm() {
    try {
        var Customer = Xrm.Page.getAttribute("hil_customerref").getValue();
        if (Customer != null) {
            var CustomerId = Customer[0].id;
            var CustomerName = Customer[0].name;
            //CustomerId = CustomerId.replace("{", "");
            //CustomerId = CustomerId.replace("}", "");
            // Set default values for the Contact form
            var formParameters = {};
            formParameters["hil_contactid"] = CustomerId;
            formParameters["hil_contactname"] = CustomerName;
            // Open the form.
            Xrm.Utility.openQuickCreate("hil_address", null, formParameters);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OpenAssetForm() {
    try {
        var Customer = Xrm.Page.getAttribute("hil_customerref").getValue();
        var PMap = Xrm.Page.getAttribute("hil_productcatsubcatmapping").getValue();
        if (Customer != null && PMap != null) {
            var CustomerId = Customer[0].id;
            var CustomerName = Customer[0].name;
            var PMapId = PMap[0].id;
            var PMapName = PMap[0].name;
            //CustomerId = CustomerId.replace("{", "");
            //CustomerId = CustomerId.replace("}", "");
            // Set default values for the Contact form
            var formParameters = {};
            formParameters["hil_xxproductcatsubcatid"] = PMapId;
            formParameters["hil_xxproductcatsubcatname"] = PMapName;
            formParameters["hil_customerguid"] = CustomerId;
            formParameters["hil_customernm"] = CustomerName;
            //formParameters["hil_customertype"] = "contact";
            // Open the form.
            Xrm.Utility.openQuickCreate("msdyn_customerasset", null, formParameters);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function SetSubStatusCancelLockJobForm() {
    debugger;
    try {
        var LoggedIn = Xrm.Page.context.getUserId();
        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/SystemUserSet(guid'" + LoggedIn + "')?$select=hil_JobCancelAuth,PositionId,position_users/Name&$expand=position_users", true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;
                if (this.status === 200) {
                    var result = JSON.parse(this.responseText).d;
                    var positionId = result.PositionId;
                    var jobCancelAuth = result.hil_JobCancelAuth;
                    var position_users_Name = result.position_users.Name;

                    //if (position_users_Name == "DSE" || position_users_Name == "Franchise" || position_users_Name == "Franchise Technician")
                    if (jobCancelAuth == null || jobCancelAuth == false || typeof jobCancelAuth == "undefined") {
                        alert("You are not authorised to cancel the Job");
                    }
                    else {
                        var ClosureRemarks = Xrm.Page.getAttribute("hil_closureremarks").getValue();
                        var JobCancelReason = Xrm.Page.getAttribute("hil_jobcancelreason").getValue();
                        if (ClosureRemarks == null) {
                            alert("Please enter closure remarks");
                        }
                        if (JobCancelReason == null) {
                            alert("Please select job cancel reason");
                        }
                        else {
                            Xrm.Utility.confirmDialog("Are you Sure you Want to Cancel this Job?",

                                function () {
                                    var req = new XMLHttpRequest();
                                    req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/msdyn_workordersubstatuses?$filter=msdyn_name eq 'Canceled'", true);
                                    req.setRequestHeader("OData-MaxVersion", "4.0");
                                    req.setRequestHeader("OData-Version", "4.0");
                                    req.setRequestHeader("Accept", "application/json");
                                    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                                    req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
                                    req.onreadystatechange = function () {
                                        if (this.readyState === 4) {
                                            req.onreadystatechange = null;
                                            if (this.status === 200) {
                                                var results = JSON.parse(this.response);
                                                for (var i = 0; i < results.value.length; i++) {
                                                    var msdyn_workordersubstatusid = results.value[i]["msdyn_workordersubstatusid"];
                                                    var lookupVal = new Array();
                                                    lookupVal[0] = new Object();
                                                    lookupVal[0].id = msdyn_workordersubstatusid;
                                                    lookupVal[0].name = "Canceled";
                                                    lookupVal[0].entityType = "msdyn_workordersubstatus";
                                                    Xrm.Page.getAttribute("msdyn_substatus").setValue(lookupVal);
                                                    var currentDATE = new Date();
                                                    //alert(currentDATE);
                                                    Xrm.Page.getAttribute("hil_jobclosureon").setValue(currentDATE);
                                                    Xrm.Page.getAttribute("msdyn_timeclosed").setValue(currentDATE);
                                                    Xrm.Page.data.entity.save();
                                                    disableformControls();
                                                    alert("Job is Cancelled");
                                                }
                                            }
                                            else {
                                                Xrm.Utility.alertDialog(this.statusText);
                                            }
                                        }
                                    };
                                    req.send();
                                },

                                function () {
                                    //alert("CANCELLED!!!");
                                });
                        }
                    }
                }
                else {
                    Xrm.Utility.alertDialog(this.statusText);
                }
            }
        };
        req.send();
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function AssignToBranchHead(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        var AssignRequest = formContext.getAttribute("hil_assigntobranchhead").getValue();
        if (AssignRequest == 1) {
            Xrm.Utility.confirmDialog("Are you sure you want to assign this job to Branch Service Head?",

                function () {
                    var Matrix = formContext.getAttribute("hil_assignmentmatrix").getValue();
                    if (Matrix != null) {
                        var MatrixId = Matrix[0].id;
                        var req = new XMLHttpRequest();
                        req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/hil_assignmentmatrixSet(guid'" + MatrixId + "')?$select=OwnerId", true);
                        req.setRequestHeader("Accept", "application/json");
                        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                        req.onreadystatechange = function () {
                            if (this.readyState === 4) {
                                this.onreadystatechange = null;
                                if (this.status === 200) {
                                    var result = JSON.parse(this.responseText).d;
                                    var owner = result.OwnerId;
                                    var object = new Array();
                                    object[0] = new Object();
                                    object[0].id = owner.Id;
                                    object[0].name = owner.Name;
                                    object[0].entityType = "systemuser";
                                    formContext.getAttribute("ownerid").setValue(object);
                                    Xrm.Page.data.entity.save();
                                }
                                else {
                                    Xrm.Utility.alertDialog(this.statusText);
                                }
                            }
                        };
                        req.send();
                    }
                    else {
                        alert("No Branch Head defined for this PinCode and Division");
                        formContext.getAttribute("ownerid").setValue(null);
                        Xrm.Page.data.entity.save();
                    }
                },

                function () { });
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function CloseTicketValidation(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        var CloseTicket = formContext.getAttribute("hil_calculatecharges").getValue();
        if (CloseTicket == true) {
            var JobQuant = formContext.getAttribute("hil_quantity").getValue();
            var IncQuantity = formContext.getAttribute("hil_incidentquantity").getValue();
            if (JobQuant > IncQuantity) {
                alert("Job Quantity can't be lesser than Incident Quantity. Please check Incident");
                formContext.getAttribute("hil_calculatecharges").setValue(false);
            }
            Scheme_Check(executionContext);
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function DisableCancelJobButtton(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        var IsJobCompleted = formContext.getAttribute("hil_calculatecharges").getValue();
        if (IsJobCompleted != null()) {
            alert(IsJobCompleted);
        }
        if (IsJobCompleted == true) {
            return true;
        }
        else {
            return false;
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function DisableButtonsonclosedjob() {
    try {
        var IsJobclosed = Xrm.Page.getAttribute("hil_closeticket").getValue();
        if (IsJobclosed == true) {
            return false;
        }
        else {
            return true;
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function IfNoActionsUsedOnJobComplete(executionContext) {
    debugger;
    try {
        var formContext = executionContext.getFormContext();
        var CloseTicket = formContext.getAttribute("hil_calculatecharges").getValue();
        var JobIncidentCount = formContext.getAttribute("hil_jobincidentcount").getValue();
        var IsOCR = formContext.getAttribute("hil_isocr").getValue();
        var id = Xrm.Page.data.entity.getId();
        id = id.replace("{", "");
        id = id.replace("}", "");
        var entityName = Xrm.Page.data.entity.getEntityName();
        var counts = 0;
        //alert(id);
        //alert(entityName);
        if (CloseTicket == true && IsOCR == false) {
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/msdyn_workorderservices?$filter=_msdyn_workorder_value eq " + id + " and  msdyn_linestatus eq 690970001", true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200) {
                        var results = JSON.parse(this.response);
                        for (var i = 0; i < results.value.length; i++) {
                            counts = counts + 1;
                            var msdyn_workorderserviceid = results.value[i]["msdyn_workorderserviceid"];
                            break;
                            //alert(counts);
                        }
                        if (counts == 0 || JobIncidentCount == null) {
                            alert("Service action is required before job completion");
                            formContext.getAttribute("hil_calculatecharges").setValue(false);
                        }
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function OnPDateChane(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        var PDate = formContext.getAttribute("hil_purchasedate").getValue();
        if (PDate != null) {
            var DateNow = new Date();
            var NowYear = DateNow.getFullYear();
            var NowMonth = DateNow.getMonth() + 1;
            var PYear = PDate.getFullYear();
            var PMonth = PDate.getMonth() + 1;
            var TimeDiff = DateNow.getTime() - PDate.getTime();
            var DayDiff = Math.ceil(TimeDiff / (1000 * 60 * 60 * 24));
            var MonDiff = DayDiff / 30.42;
            //alert("Differnece between dates  is : "+DayDiff+" days");
            //alert("Difference between months is : "+MonDiff);
            //var MonDiff = (NowYear - PYear) * 12 + NowMonth - PMonth;
            var ProdSC = formContext.getAttribute("hil_productsubcategory").getValue();
            if (ProdSC != null) {
                var ProdId = ProdSC[0].id;
                ProdId = ProdId.replace("{", "");
                ProdId = ProdId.replace("}", "");
                var req = new XMLHttpRequest();
                req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/hil_warrantytemplateSet?$select=hil_Product,hil_type,hil_WarrantyPeriod&$filter=hil_Product/Id eq (guid'" + ProdId + "') and hil_type/Value eq 1", true);
                req.setRequestHeader("Accept", "application/json");
                req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                req.onreadystatechange = function () {
                    if (this.readyState === 4) {
                        this.onreadystatechange = null;
                        if (this.status === 200) {
                            var returned = JSON.parse(this.responseText).d;
                            var results = returned.results;
                            for (var i = 0; i < results.length; i++) {
                                var type = results[i].hil_type;
                                var typeValue = type.Value;
                                var WarrantyPeriod = results[i].hil_WarrantyPeriod;
                                if (MonDiff < WarrantyPeriod) {
                                    alert("IN WARRANTY");
                                    formContext.getAttribute("hil_warrantystatus").setValue(1);
                                }
                                else {
                                    //alert("Hi");
                                    var SubCat = formContext.getAttribute("hil_productsubcategory").getValue();
                                    if (SubCat != null) {
                                        var SubCatId = SubCat[0].id;
                                        SubCatId = SubCatId.replace("{", "");
                                        SubCatId = SubCatId.replace("}", "");
                                        //alert(SubCatId);
                                        var req1 = new XMLHttpRequest();
                                        req1.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/products(" + SubCatId + ")?$select=hil_standardvisitcharge", true);
                                        req1.setRequestHeader("OData-MaxVersion", "4.0");
                                        req1.setRequestHeader("OData-Version", "4.0");
                                        req1.setRequestHeader("Accept", "application/json");
                                        req1.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                                        req1.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
                                        req1.onreadystatechange = function () {
                                            if (this.readyState === 4) {
                                                req1.onreadystatechange = null;
                                                if (this.status === 200) {
                                                    var result1 = JSON.parse(this.response);
                                                    var hil_standardvisitcharge = result1["hil_standardvisitcharge"];
                                                    var hil_standardvisitcharge_formatted = result1["hil_standardvisitcharge@OData.Community.Display.V1.FormattedValue"];
                                                    alert("OUT WARRANTY\nStandard Visit Charges may be applicable\nVisit Charge Amount Rs." + hil_standardvisitcharge_formatted);
                                                    formContext.getAttribute("hil_warrantystatus").setValue(2);
                                                }
                                                else {
                                                    Xrm.Utility.alertDialog(this.statusText);
                                                }
                                            }
                                        };
                                        req1.send();
                                    }
                                }
                            }
                        }
                        else {
                            Xrm.Utility.alertDialog(this.statusText);
                        }
                    }
                };
                req.send();
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function ValidateKKG(executionContext) {
    try {
        debugger;
        var FormType = Xrm.Page.ui.getFormType();
        if (FormType == 2) {
            if (ValidateOnlyNumeric(executionContext)) {
                var formContext = executionContext.getFormContext();
                var jobId = Xrm.Page.data.entity.getId();
                var kkgCode = formContext.getAttribute("hil_kkgcode").getValue();
                var kkgOTP;
                var _kkgOTPFetch = '<fetch mapping="logical" version="1.0" output-format="xml-platform" distinct="false"><entity name="msdyn_workorder"><attribute name="hil_kkgotp" /><filter type="and"><condition value="' + jobId + '" attribute="msdyn_workorderid" operator="eq" /></filter></entity></fetch>';
                var _result = XrmServiceToolkit.Soap.Fetch(_kkgOTPFetch);
                for (var i = 0; i < _result.length; i++) {
                    kkgOTP = _result[i].attributes.hil_kkgotp;
                }
                Xrm.Page.getControl("hil_kkgcode").clearNotification();
                if (kkgCode != null && typeof kkgCode != "undefined" && kkgOTP != null && typeof kkgOTP != "undefined" && (kkgCode != atob(kkgOTP.value) && kkgCode != kkgOTP.value)) {
                    Xrm.Page.getControl("hil_kkgcode").setNotification("Invalid KKG Code Entered.");
                }
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function ValidateKKGOnSave(executionContext) {
    try {
        debugger;
        var FormType = Xrm.Page.ui.getFormType();
        if (FormType == 2) {
            var formContext = executionContext.getFormContext();
            var jobId = Xrm.Page.data.entity.getId();
            var kkgCode = formContext.getAttribute("hil_kkgcode").getValue();
            var kkgOTP;
            var _kkgOTPFetch = '<fetch mapping="logical" version="1.0" output-format="xml-platform" distinct="false"><entity name="msdyn_workorder"><attribute name="hil_kkgotp" /><filter type="and"><condition value="' + jobId + '" attribute="msdyn_workorderid" operator="eq" /></filter></entity></fetch>';
            var _result = XrmServiceToolkit.Soap.Fetch(_kkgOTPFetch);
            for (var i = 0; i < _result.length; i++) {
                kkgOTP = _result[i].attributes.hil_kkgotp;
            }
            Xrm.Page.getControl("hil_kkgcode").clearNotification();
            if (kkgCode != null && typeof kkgCode != "undefined" && kkgOTP != null && typeof kkgOTP != "undefined" && (kkgCode != atob(kkgOTP.value) && kkgCode != kkgOTP.value)) {
                Xrm.Page.getControl("hil_kkgcode").setNotification("Invalid KKG Code Entered.");
                executionContext.getEventArgs().preventDefault();
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function ValidateKKGDisposition(executionContext) {
    try {
        debugger;
        var createChildJobValue = Xrm.Page.getAttribute("hil_createchildjob").getValue();
        var requiredLevel = Xrm.Page.getAttribute("hil_createchildjob").getRequiredLevel();
        if (createChildJobValue === false && requiredLevel == "required") {
            Xrm.Utility.alertDialog("Please select 'Create Child Job' in case of KKG Failure.");
            executionContext.getEventArgs().preventDefault();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function createChildJobClick() {
    try {
        debugger;
        var createChildJobValue = Xrm.Page.getAttribute("hil_createchildjob").getValue();
        if (createChildJobValue === true) {
            var LoggedIn = Xrm.Page.context.getUserId();
            var subStatus = Xrm.Page.data.entity.attributes.get("msdyn_substatus").getValue();
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/SystemUserSet(guid'" + LoggedIn + "')?$select=PositionId,position_users/Name&$expand=position_users", true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.responseText).d;
                        var positionId = result.PositionId;
                        var position_users_Name = result.position_users.Name;
                        if (position_users_Name == "Call Center" && subStatus[0].name == "Work Done") {
                            Xrm.Utility.confirmDialog("Are you sure to create child Job?",

                                function () {
                                    Xrm.Page.data.entity.save();
                                    Xrm.Utility.alertDialog("Child Job has been created successfully.");
                                },

                                function () {
                                    Xrm.Page.getAttribute("hil_createchildjob").setValue(false);
                                    return;
                                });
                        }
                        else if (position_users_Name != "Call Center") {
                            Xrm.Utility.alertDialog("You are not authorised to perform this action.");
                            Xrm.Page.getAttribute("hil_createchildjob").setValue(false);
                            return;
                        }
                        else if (subStatus[0].name != "Work Done") {
                            Xrm.Utility.alertDialog("KKG Disposition can only be selected in Work Done Status.");
                            Xrm.Page.getAttribute("hil_createchildjob").setValue(false);
                            return;
                        }
                    }
                    else {
                        Xrm.Utility.alertDialog(this.statusText);
                    }
                }
            };
            req.send();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function lockFieldsOnceWorkDone() {
    try {
        if (Xrm.Page.getAttribute("msdyn_substatus") != null) {
            var SubStatus = Xrm.Page.getAttribute("msdyn_substatus").getValue();
            if (SubStatus != null) {
                var StatusName = SubStatus[0].name.toUpperCase();
                if (StatusName == "WORK DONE" || StatusName == "CLOSED" || StatusName == "CANCELED") {
                    if (!isControlNull("hil_customerref")) {
                        Xrm.Page.getControl("hil_customerref").setDisabled(true);
                    }
                    if (!isControlNull("hil_productcatsubcatmapping")) {
                        Xrm.Page.getControl("hil_productcatsubcatmapping").setDisabled(true);
                    }
                    if (!isControlNull("hil_consumertype")) {
                        Xrm.Page.getControl("hil_consumertype").setDisabled(true);
                    }
                    if (!isControlNull("hil_consumercategory")) {
                        Xrm.Page.getControl("hil_consumercategory").setDisabled(true);
                    }
                    if (!isControlNull("hil_natureofcomplaint")) {
                        Xrm.Page.getControl("hil_natureofcomplaint").setDisabled(true);
                    }
                    if (!isControlNull("hil_callsubtype")) {
                        Xrm.Page.getControl("hil_callsubtype").setDisabled(true);
                    }
                    if (!isControlNull("msdyn_customerasset")) {
                        Xrm.Page.getControl("msdyn_customerasset").setDisabled(true);
                    }
                    if (!isControlNull("hil_quantity")) {
                        Xrm.Page.getControl("hil_quantity").setDisabled(true);
                    }
                    if (!isControlNull("hil_owneraccount")) {
                        Xrm.Page.getControl("hil_owneraccount").setDisabled(true);
                    }
                    if (!isControlNull("hil_typeofassignee")) {
                        Xrm.Page.getControl("hil_typeofassignee").setDisabled(true);
                    }
                    if (!isControlNull("hil_sourceofjob")) {
                        Xrm.Page.getControl("hil_sourceofjob").setDisabled(true);
                    }
                    if (!isControlNull("msdyn_serviceaccount")) {
                        Xrm.Page.getControl("msdyn_serviceaccount").setDisabled(true);
                    }
                    if (!isControlNull("hil_regardingfallback")) {
                        Xrm.Page.getControl("hil_regardingfallback").setDisabled(true);
                    }
                    if (!isControlNull("hil_salesoffice")) {
                        Xrm.Page.getControl("hil_salesoffice").setDisabled(true);
                    }
                    if (!isControlNull("hil_state")) {
                        Xrm.Page.getControl("hil_state").setDisabled(true);
                    }
                    if (!isControlNull("hil_assignmentmatrix")) {
                        Xrm.Page.getControl("hil_assignmentmatrix").setDisabled(true);
                    }
                    if (!isControlNull("hil_region")) {
                        Xrm.Page.getControl("hil_region").setDisabled(true);
                    }
                    if (!isControlNull("hil_pincode")) {
                        Xrm.Page.getControl("hil_pincode").setDisabled(true);
                    }
                    if (!isControlNull("hil_productsubcategory")) {
                        Xrm.Page.getControl("hil_productsubcategory").setDisabled(true);
                    }
                    if (!isControlNull("hil_mobilenumber")) {
                        Xrm.Page.getControl("hil_mobilenumber").setDisabled(true);
                    }
                    if (!isControlNull("hil_callingnumber")) {
                        //Xrm.Page.getControl("hil_callingnumber").setDisabled(true);
                    }
                    if (!isControlNull("hil_district")) {
                        Xrm.Page.getControl("hil_district").setDisabled(true);
                    }
                    if (!isControlNull("hil_city")) {
                        Xrm.Page.getControl("hil_city").setDisabled(true);
                    }
                    if (!isControlNull("hil_area")) {
                        Xrm.Page.getControl("hil_area").setDisabled(true);
                    }
                    if (!isControlNull("hil_branch")) {
                        Xrm.Page.getControl("hil_branch").setDisabled(true);
                    }
                    if (!isControlNull("hil_brand")) {
                        Xrm.Page.getControl("hil_brand").setDisabled(true);
                    }
                    if (!isControlNull("hil_appointmentstatus")) {
                        Xrm.Page.getControl("hil_appointmentstatus").setDisabled(true);
                    }
                    if (!isControlNull("msdyn_billingaccount")) {
                        Xrm.Page.getControl("msdyn_billingaccount").setDisabled(true);
                    }
                    if (!isControlNull("hil_address")) {
                        Xrm.Page.getControl("hil_address").setDisabled(true);
                    }
                    if (!isControlNull("hil_productcategory")) {
                        Xrm.Page.getControl("hil_productcategory").setDisabled(true);
                    }
                    if (!isControlNull("ownerid")) {
                        Xrm.Page.getControl("ownerid").setDisabled(true);
                    }
                }
            }
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function isControlNull(control) {
    try {
        if (Xrm.Page.getControl(control) == null || typeof Xrm.Page.getControl(control) == "undefined") {
            return true;
        }
        else {
            return false;
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}

function ShowMessage(alertMsg, titleMsg) {
    var alertStrings = { confirmButtonLabel: "OK", text: alertMsg, title: titleMsg };
    var alertOptions = { height: 150, width: 300 };
    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
        function success(result) {
            console.log(alertMsg);
        },
        function (error) {
            console.log(error.message);
        }
    );
}

function AMCJobClosureValidations(executionContext) {
    debugger;
    try {
        formContext = executionContext.getFormContext();

        var email = formContext.getAttribute("hil_email").getValue();
        var callSubType = formContext.getAttribute("hil_callsubtype").getValue();
        var closeticket = formContext.getAttribute("hil_closeticket").getValue();
        var message = '';
        if (callSubType != null && typeof callSubType != "undefined") {
            if (callSubType[0].name.toUpperCase().indexOf("AMC") >= 0 && closeticket) {
                if (email == null || typeof email == "undefined") {
                    message = "Please Update Consumer's Email Id this will be required to send AMC Invoice to the Consumser.";
                }
                else {
                    message = "Please validate Consumer's Email Id : " + email + ". The same will be used to send AMC Invoice to the Consumser.";
                }
                var confirmStrings = { text: message, title: "Email Verification", confirmButtonLabel: "OK", cancelButtonLabel: "Cancel & Update" };
                Xrm.Navigation.openConfirmDialog(confirmStrings, null).then(
                    function (success) {
                        if (success.confirmed) {
                            console.log("OK pressed");
                        }
                        else {
                            formContext.getAttribute("hil_closeticket").setValue(false);
                            Xrm.Page.ui.tabs.get("Tab_Summary").setFocus();
                            formContext.getControl("hil_email").setFocus();
                            executionContext.getEventArgs().preventDefault();
                        }
                    });
            }
        }
    }
    catch (error) {
        ShowMessage(error.message, "AMC Call Closure Validation.");
    }
}
function PreFilterSchemeCodes() {
    try {
        Xrm.Page.getControl("hil_schemecode").addPreSearch(function () {
            debugger;
            var createdOn = Xrm.Page.getAttribute("createdon").getValue();
            createdOn = (createdOn.getFullYear() + "-" + ((createdOn.getMonth() < 9 ? '0' : '') + (createdOn.getMonth() + 1)) + "-" + ((createdOn.getDate() < 10 ? '0' : '') + createdOn.getDate()));
            var purchaseDate = Xrm.Page.getAttribute("hil_purchasedate").getValue();
            if (purchaseDate == null || typeof purchaseDate == "undefined") {
                purchaseDate = "1900-01-01";
            }
            else {
                purchaseDate = (purchaseDate.getFullYear() + "-" + ((purchaseDate.getMonth() < 9 ? '0' : '') + (purchaseDate.getMonth() + 1)) + "-" + ((purchaseDate.getDate() < 10 ? '0' : '') + purchaseDate.getDate()));
            }
            var LookupId = null;
            var salesoffice = Xrm.Page.getAttribute("hil_salesoffice").getValue();
            var callsubtype = Xrm.Page.getAttribute("hil_callsubtype").getValue();
            var productsubcategory = Xrm.Page.getAttribute("hil_productsubcategory").getValue();
            var filterStr = "<filter type='and'>";
            filterStr = filterStr + "<condition attribute='hil_salesoffice' operator='in'>"
            filterStr = filterStr + "<value >{90503976-8FD1-EA11-A813-000D3AF0563C}</value>" //All Sales Office
            if (salesoffice != null && typeof salesoffice != "undefined") {
                LookupId = salesoffice[0].id;
                filterStr = filterStr + "<value >" + LookupId + "</value>" //Job Sales Office
            }
            filterStr = filterStr + "</condition>"
            filterStr = filterStr + "<condition attribute='hil_schemeexpirydate' operator='on-or-after' value='" + createdOn + "' />"
            filterStr = filterStr + "<condition attribute='hil_fromdate' operator='on-or-before' value='" + purchaseDate + "' />"
            filterStr = filterStr + "<condition attribute='hil_todate' operator='on-or-after' value='" + purchaseDate + "' />"
            if (callsubtype != null) {
                LookupId = callsubtype[0].id;
                filterStr = filterStr + "<condition attribute='hil_callsubtype' operator='eq' value='" + LookupId + "' />"
            }

            if (productsubcategory != null) {
                LookupId = productsubcategory[0].id;
                filterStr = filterStr + "<condition attribute='hil_productsubcategory' operator='eq' value='" + LookupId + "' />"
            }
            var filterStr = filterStr + "</filter>";
            Xrm.Page.getControl("hil_schemecode").addCustomFilter(filterStr);
        });
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}
function Scheme_Check(eContext) {
    debugger;
    var formContext = eContext.getFormContext();

    var schemecode = formContext.getAttribute("hil_schemecode").getValue();
    if (schemecode == null || typeof schemecode == "undefined") {
        var salesoffice = formContext.getAttribute("hil_salesoffice").getValue();
        if (salesoffice != null && typeof salesoffice != "undefined") {
            salesoffice = salesoffice[0].id;
        }
        var callsubtype = formContext.getAttribute("hil_callsubtype").getValue();
        var productsubcategory = formContext.getAttribute("hil_productsubcategory").getValue();
        if (callsubtype != null) {
            callsubtype = callsubtype[0].id;
        }
        if (productsubcategory != null) {
            productsubcategory = productsubcategory[0].id;
        }
        var createdOn = formContext.getAttribute("createdon").getValue();
        createdOn = (createdOn.getFullYear() + "-" + ((createdOn.getMonth() < 9 ? '0' : '') + (createdOn.getMonth() + 1)) + "-" + ((createdOn.getDate() < 10 ? '0' : '') + createdOn.getDate()));
        var purchaseDate = formContext.getAttribute("hil_purchasedate").getValue();
        if (purchaseDate == null || typeof purchaseDate == "undefined") {
            purchaseDate = "1900-01-01";
        }
        else {
            purchaseDate = (purchaseDate.getFullYear() + "-" + ((purchaseDate.getMonth() < 9 ? '0' : '') + (purchaseDate.getMonth() + 1)) + "-" + ((purchaseDate.getDate() < 10 ? '0' : '') + purchaseDate.getDate()));
        }
        var fetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='hil_schemeincentive'>" +
            "<attribute name='hil_schemeincentiveid' />" +
            "<attribute name='hil_name' />" +
            "<attribute name='createdon' />" +
            "<order attribute='hil_name' descending='false' />" +
            "<filter type='and'>" +
            "<condition attribute='hil_salesoffice' operator='in'>" +
            "<value>{90503976-8FD1-EA11-A813-000D3AF0563C}</value>" +
            (salesoffice != null ? "<value>" + salesoffice + "</value>" : "") +
            "</condition>" +
            "<condition attribute='hil_schemeexpirydate' operator='on-or-after' value='" + createdOn + "' />" +
            "<condition attribute='hil_fromdate' operator='on-or-before' value='" + purchaseDate + "' />" +
            "<condition attribute='hil_todate' operator='on-or-after' value='" + purchaseDate + "' />" +
            (callsubtype != null ? "<condition attribute='hil_callsubtype' operator='eq'  value='" + callsubtype + "' />" : "") +
            (productsubcategory != null ? "<condition attribute='hil_productsubcategory' operator='eq' value='" + productsubcategory + "' />" : "") +
            "</filter>" +
            "</entity>" +
            "</fetch>";

        fetchXml = "?fetchXml=" + encodeURIComponent(fetchXml);
        Xrm.WebApi.retrieveMultipleRecords('hil_schemeincentive', fetchXml).then(
            function success(result) {
                debugger;
                if (result.entities.length > 0) {
                    debugger;
                    var confirmStrings = {
                        text: "Please select the Scheme Code, if any Installation scheme is running for this Installation.This data will be audited & may result into penalty or rejection of Job Claim in case wrong Scheme Code selected.", title: "Confirmation Dialog", cancelButtonLabel: "Continue to Save", confirmButtonLabel: "Select Scheme Code"
                    };
                    var confirmOptions = { height: 300, width: 450 };
                    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                        function (success) {
                            if (success.confirmed) {
                                debugger;
                                formContext.getAttribute("hil_closeticket").setValue(0);
                                formContext.getControl("hil_closeticket").setFocus();
                                console.log("Dialog closed using Okay button.");
                            }
                            else {
                                debugger;
                                formContext.getAttribute("hil_closeticket").setValue(1);
                                formContext.data.entity.save();
                                console.log("Dialog closed using Cancel button or X.");
                            }
                        });
                }
            },
            function (error) {
                ShowMessage(error.message, "Error!!");
            }
        );
    }
	/*else{
		var jobId = formContext.data.entity.getId();
		 fetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
			"<entity name='hil_sawactivity'>" +
			"<attribute name='hil_sawactivityid' />" +
			"<attribute name='hil_name' />" +
			"<attribute name='createdon' />" +
			"<order attribute='hil_name' descending='false' />" +
			"<filter type='and'>" +
			"<condition attribute='hil_jobid' operator='eq' value='" + jobId + "' />" +
			"<condition attribute='hil_approvalstatus' operator='ne' value='3' />" +
			"</filter>" +
			"</entity>" +
			"</fetch>";
		fetchXml = "?fetchXml=" + encodeURIComponent(fetchXml);
		Xrm.WebApi.retrieveMultipleRecords('hil_sawactivity', fetchXml).then(
			function success(result) {
				debugger;
				if (result.entities.length > 0) {
					ShowMessage("SAW is not approved", "ERROR!!!");
					formContext.getAttribute("hil_closeticket").setValue(0);
				}
				else {
					formContext.getAttribute("hil_closeticket").setValue(1);
					formContext.data.entity.save();
				}
			},
			function (error) {
				ShowMessage(error.message, "Error!!");
			}
		);
	} */
}
function GasCharge_Check() {
    debugger;
    //var formContext = eContext.getFormContext();
    var gasChargeStatus = Xrm.Page.getAttribute("hil_isgascharged").getValue();
    if (gasChargeStatus == true || gasChargeStatus == 1) {
        ShowMessage("Gas Charge already added to this Job.", "Information.");
        return;
    }
    var confirmStrings = {
        text: "Are you sure you want to add Gas Charge?", title: "Confirmation Dialog", cancelButtonLabel: "No", confirmButtonLabel: "Yes"
    };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                debugger;
                Xrm.Page.getAttribute("hil_isgascharged").setValue(true);
                Xrm.Page.data.entity.save();
            }
            else {
                debugger;
                Xrm.Page.getAttribute("hil_isgascharged").setValue(false);
                Xrm.Page.data.entity.save();
            }
        });
}
function RbnEnableRule_GasCharge() {
    debugger;
    var gasChargeStatus = Xrm.Page.getAttribute("hil_isgascharged").getValue();
    var subStatus = Xrm.Page.data.entity.attributes.get("msdyn_substatus").getValue();
    var callSubType = Xrm.Page.data.entity.attributes.get("hil_callsubtype").getValue();
    var prodCategory = Xrm.Page.data.entity.attributes.get("hil_productcategory").getValue();
    if (subStatus != null && typeof subStatus != "undefined" && callSubType != null && typeof callSubType != "undefined" && prodCategory != null && typeof prodCategory != "undefined") {
        if ((typeof gasChargeStatus == "undefined" || gasChargeStatus == null || gasChargeStatus == false) && subStatus[0].name == "Work Done" && (callSubType[0].id == "{E2129D79-3C0B-E911-A94E-000D3AF06CD4}" || callSubType[0].id == "{6560565A-3C0B-E911-A94E-000D3AF06CD4}" || callSubType[0].id == "{8D80346B-3C0B-E911-A94E-000D3AF06CD4}") && (prodCategory[0].id == "{D51EDD9D-16FA-E811-A94C-000D3AF0694E}" || prodCategory[0].id == "{2DD99DA1-16FA-E811-A94C-000D3AF06091}")) {
            return true;
        }
        else {
            return false;
        }
    }
    else { return false; }
}

function RbnEnableRule_MarkAsUpcountry() {
    try {
        debugger;
        var FormType = Xrm.Page.ui.getFormType();
        var reviewforcountryclassification = Xrm.Page.getAttribute("hil_reviewforcountryclassification").getValue();
        var isClaimGenerated = Xrm.Page.getAttribute("hil_generateclaim").getValue();
        var countryclassification = Xrm.Page.getAttribute("hil_countryclassification").getValue();
        var _retValue = false;
        if (FormType != 1 && reviewforcountryclassification == true && countryclassification != 2 && isClaimGenerated != true) {
            _retValue = true;
        }
        return _retValue;
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}
function MarkAsUpcountry() {
    try {
        debugger;
        var FormType = Xrm.Page.ui.getFormType();
        var LoggedIn = Xrm.Page.context.getUserId();
        if (FormType != 1) {
            var req = new XMLHttpRequest();
            req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/SystemUserSet(guid'" + LoggedIn + "')?$select=hil_JobCancelAuth,PositionId,position_users/Name&$expand=position_users", true);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    this.onreadystatechange = null;
                    if (this.status === 200) {
                        var result = JSON.parse(this.responseText).d;
                        var position_users_Name = result.position_users.Name.toUpperCase();
                        if (position_users_Name == "BSH" || position_users_Name == "NSH") {
                            var confirmStrings = { text: "Are you sure you want to make this Job as Upcountry?", title: "Confirmation Dialog", cancelButtonLabel: "No", confirmButtonLabel: "Yes" };
                            var confirmOptions = { height: 200, width: 450 };
                            Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                                function (success) {
                                    if (success.confirmed) {
                                        debugger;
                                        Xrm.Page.getAttribute("hil_webclosureremarks").setValue("Upcountry Marked By BSH");
                                        Xrm.Page.getAttribute("hil_countryclassification").setValue(2);
                                        Xrm.Page.data.entity.save();
                                    }
                                });
                        }
                        else {
                            ShowMessage("You are not authorised to mark this Job as Upcountry.", "Information.");
                            return;
                        }
                    }
                }
            };
            req.send();
        }
    }
    catch (err) {
        Xrm.Utility.alertDialog(err.message);
    }
}