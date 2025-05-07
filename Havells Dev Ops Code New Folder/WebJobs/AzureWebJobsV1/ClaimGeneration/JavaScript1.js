function SetFranchiseeAccount(executionContext)
{
	try
	{
		var formContext = executionContext.getFormContext();
		var FType = formContext.ui.getFormType();
		if (FType == 1)
		{
			var iUserId = Xrm.Page.context.getUserId();
			var req = new XMLHttpRequest();
			req.open("GET", Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc/SystemUserSet(guid'" + iUserId + "')?$select=hil_Account", true);
			req.setRequestHeader("Accept", "application/json");
			req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			req.onreadystatechange = function ()
			{
				if (this.readyState === 4)
				{
					this.onreadystatechange = null;
					if (this.status === 200)
					{
						var result = JSON.parse(this.responseText).d;
						var Account = result.hil_Account;
						if (Account != null)
						{
							var object = new Array();
							object[0] = new Object();
							object[0].id = Account.Id;
							object[0].name = Account.Name;
							object[0].entityType = "account";
							formContext.getAttribute("hil_franchisee").setValue(object);
							SetFranchiseManager(executionContext, Account.Id);
						}
						else
						{
							alert("Claim Account profile is not set in your user setup.");
						}
					}
					else
					{
						Xrm.Utility.alertDialog(this.statusText);
					}
				}
			};
			req.send();
		}
	}
	catch (err)
	{
		Xrm.Utility.alertDialog(err.message);
	}
}

function SetCurrentClaimMonth(executionContext) {
	try {
		var formContext = executionContext.getFormContext();
		var FType = formContext.ui.getFormType();
		if (FType == 1) {
			debugger;
			var req = new XMLHttpRequest();
			req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/hil_claimperiods?$filter=hil_isapplicable eq true and hil_claimfreeze eq false", true);
			req.setRequestHeader("Accept", "application/json");
			req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			req.onreadystatechange = function () {
				if (this.readyState === 4) {
					this.onreadystatechange = null;
					if (this.status === 200) {
						var result = JSON.parse(this.responseText).value[0];
						if (result != null) {
							var object = new Array();
							object[0] = new Object();
							object[0].id = result.hil_claimperiodid;
							object[0].name = result.hil_name;
							object[0].entityType = "hil_claimperiod";
							formContext.getAttribute("hil_fiscalmonth").setValue(object);
						}
						else {
							alert("Claim Month is not defined.");
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


function SetFranchiseManager(executionContext,franchiseeId) {
	try {
		var formContext = executionContext.getFormContext();
		var FType = formContext.ui.getFormType();
		if (FType == 1) {
			debugger;
			var req = new XMLHttpRequest();
			req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/hil_assignmentmatrixes?$top=1&$select=ownerid&$filter=_hil_franchiseedirectengineer_value eq '" + franchiseeId + "' and statuscode eq 1", true);
			req.setRequestHeader("Accept", "application/json");
			req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			req.onreadystatechange = function () {
				if (this.readyState === 4) {
					this.onreadystatechange = null;
					if (this.status === 200) {
						var result = JSON.parse(this.responseText).value[0];
						var Manager = result.ownerid;
						if (result != null) {
							var object = new Array();
							object[0] = new Object();
							object[0].id = Manager.Id;
							object[0].name = Manager.Name;
							object[0].entityType = "systemuser";
							formContext.getAttribute("hil_branchmanager").setValue(object);
						}
						else {
							alert("Franchise Manager is not defined.");
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
							var confirmStrings = { text: "Are you sure, you want to make this Job as Up Country?", title: "Confirmation Dialog", cancelButtonLabel: "No", confirmButtonLabel: "Yes" };
							var confirmOptions = { height: 200, width: 450 };
							Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
								function (success) {
									if (success.confirmed) {
										debugger;
										Xrm.Page.getAttribute("hil_countryclassification").setValue(2);
										Xrm.Page.data.entity.save();
									}
								});
						}
						else {
							ShowMessage("You are not Authorised to mark this Job as Up Country.", "Information.");
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