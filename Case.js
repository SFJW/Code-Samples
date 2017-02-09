if (typeof (AV) == "undefined")
{ AV = {}; }

AV.MC = {};
AV.MC.Case = {
    onLoad: function () {
        AV.MC.Case.socialPaneMods();
        AV.MC.Case.filterCustomer();
        AV.MC.Case.filterDisposition();
        AV.MC.Case.showHideSensitiveInfo();
    },
    // open full phone call and task forms from social pane
    socialPaneMods: function() {
        if (Xrm.Page.ui.getFormType() > 1) {
            var options = { openInNewWindow: true };
            var stypecode = Xrm.Internal.getEntityCode(Xrm.Page.data.entity.getEntityName());
            var ptypecode = Xrm.Internal.getEntityCode("phonecall");
            var ttypecode = Xrm.Internal.getEntityCode("task");
            parent.$("#activityLabelinlineactivitybar4212").first().attr("onclick", "Xrm.Utility.openEntityForm('task', null, {_CreateFromId : '" + Xrm.Page.data.entity.getId() + "', _CreateFromType: " + stypecode + ", etc:" + ttypecode + "}, { openInNewWindow : true })");
            parent.$("#activityLabelinlineactivitybar4210").first().attr("onclick", "Xrm.Utility.openEntityForm('phonecall', null, {_CreateFromId : '" + Xrm.Page.data.entity.getId() + "', _CreateFromType: " + stypecode + ", etc:" + ptypecode + "}, { openInNewWindow : true })");
        }
    },
    // only show contacts
    filterCustomer: function () {
        var fetchXml =  '<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">' +
                        '  <entity name="contact">' +
                        '    <attribute name="fullname" />' +
                        '    <attribute name="emailaddress1" />' +
                        '    <attribute name="telephone1" />' +
                        '    <attribute name="parentcustomerid" />' +
                        '    <attribute name="address1_city" />' +
                        '    <attribute name="address1_telephone1" />' +
                        '    <order attribute="fullname" descending="false" />' +
                        '  </entity>' +
                        '</fetch>';
        var layoutXml = "<grid name='resultset' object='1' jump='fullname' select='1' icon='1' preview='1'>" +
                        "<row name='result' id='contactid'>" +
                        "<cell name='fullname' width='300' />" +
                        "<cell name='emailaddress1' width='200' />" +
                        "<cell name='telephone1' width='125' />" +
                        "<cell name='parentcustomerid' width='150' />" +
                        "<cell name='address1_city' width='100' />" +
                        "<cell name='address1_telephone1' width='125' />" +
                        "</row>" +
                        "</grid>";
        // form field
        Xrm.Page.getControl("customerid").addCustomView("{00000000-0000-0000-0000-000000000001}", "contact", "Contacts Lookup View", fetchXml, layoutXml, true);
        Xrm.Page.getControl("customerid").addPreSearch(addFilter);

        // business process flow field
        Xrm.Page.getControl("header_process_customer").addCustomView("{00000000-0000-0000-0000-000000000001}", "contact", "Contacts Lookup View", fetchXml, layoutXml, true);
        Xrm.Page.getControl("header_process_customer").addPreSearch(addFilter);

        function addFilter() {
            var customerContactFilter = "<filter type='and'><condition attribute='accountid' operator='null' /></filter>";
            Xrm.Page.getControl("customerid").addCustomFilter(customerContactFilter, "account");
            Xrm.Page.getControl("header_process_customer").addCustomFilter(customerContactFilter, "account");
        }
    },
    // show dipositions related to the case's service unit
    filterDisposition: function () {
        var service = getValue("mc_serviceunit");

        if (service != null) {
            Xrm.Page.getControl("header_process_mc_dispositionid").addPreSearch(addFilter);

            function addFilter() {
                var serviceFilter = '<filter type="and"><condition attribute="mc_serviceunit" operator="eq" uitype="mc_serviceunit" value="' + service[0].id + '" /></filter>';
                Xrm.Page.getControl("header_process_mc_dispositionid").addCustomFilter(serviceFilter, "mc_disposition");
            }
        }
    },
    showHideSensitiveInfo: function () {
        var sensitiveSection = Xrm.Page.ui.tabs.get("general").sections.get("sensitive");
        var currentUserId = Xrm.Page.context.getUserId();
        var ownerId = getValue("ownerid") ? getValue("ownerid")[0].id : "";
        var ownerType = getValue("ownerid") ? getValue("ownerid")[0].entityType : "";
        var show = false;

        // if current user is in Counseling BU - section/tab is visible, regardless of Case owner
        if (UserInBU(currentUserId, "Counseling")) {
            sensitiveSection.setVisible(true);
            show = true;
        }

        // if current user is in Disability BU - show only if the Case Owner is in Disability BU or College BU
        if (UserInBU(currentUserId, "Disability Services") && (OwnerInBU(ownerType, ownerId, "Disability Services") || OwnerInBU(ownerType, ownerId, "College"))) {
            sensitiveSection.setVisible(true);
            show = true;
        }

        // if current user is in College BU - show only if the Case Owner is in the College BU
        if (UserInBU(currentUserId, "College") && OwnerInBU(ownerType, ownerId, "College")) {
            sensitiveSection.setVisible(true);
            show = true;
        }

        // if current user is a System Administrator
        if (CheckUserRole("System Administrator")) {
            sensitiveSection.setVisible(true);
            show = true;
        }

        if (getValue("av_showtopics") != show) {
            // update (existing record)
            if (Xrm.Page.ui.getFormType() == 2) {
                setValue("av_showtopics", show);
                Xrm.Page.data.entity.save(null);
            }

            Xrm.Page.ui.refreshRibbon();
        }
    },
    associateTopic: function () {
        var topicsGrid = Xrm.Page.getControl("Topics");
        if (getValue("mc_topic") && topicsGrid != null) {
            var topicId = getValue("mc_topic")[0].id;

            SDK.REST.associateRecords(topicId, "mc_topic", "mc_N_to_N_topic_incident", Xrm.Page.data.entity.getId(), "Incident",
                function successCallback() {
                    console.log("Associated the topic with this case.");
                    setValue("mc_topic", null);
                    topicsGrid.refresh();
                },
                function errorCallback(error) {
                    console.log(error);
                });
        }
    },
    addTopics: function () {
        var serviceActivity = getValue("mc_serviceactivity");

        if (serviceActivity) {
            var serviceActivityId = cleanCurlies(serviceActivity[0].id);
            var currentId = cleanCurlies(Xrm.Page.data.entity.getId());

            var vDialogOption = { width: 800, height: 640 };

            var url = Xrm.Page.context.getClientUrl() + Mscrm.CrmUri.create("$webresource:").get_path() + "/av_AddCaseTopics.html";
            url = url.indexOf('?') > 0 ? url : url + "?";
            url += "Data=ids=" + currentId + serviceActivityId;

            if (url) {
                Xrm.Internal.openDialog(url, vDialogOption, null, null, function (returnValue) {
                    Xrm.Page.data.refresh(true);
                });

                _thread(function () {
                    var iFrame = parent.document.getElementById("InlineDialog_Iframe");

                    if (!iFrame)
                        return false;

                    if ($(iFrame.contentDocument).find(".ms-crm-RefreshDialog-Header").length > 0) {
                        return true;
                    }
                    return false;
                })

            }
        }
        else {
            alert("Please specify a service activity.");
        }
    }
};

// ownerType = schema name of owner type (SystemUser or Team)
// ownerId = user/team's ID
// BU = business unit to look for (exact string)
function OwnerInBU(ownerType, ownerId, BU) {
    var type = "";

    if (!ownerType) {
        return false;
    }
    else if (ownerType == "team") {
        type = "Team";
    }
    else if (ownerType == "systemuser") {
        type = "SystemUser";
    }

    var match = false;

    SDK.REST.retrieveRecordSync(ownerId, type, "BusinessUnitId", null,
            function successCallback(record) {
                if (record.BusinessUnitId.Id != null && record.BusinessUnitId.Name == BU) {
                    match = true;
                }
            }, function errorCallback(error) {
                console.log(error);
            });

    return match;
}

function UserInBU(userId, BU) {
    var match = false;

    SDK.REST.retrieveRecordSync(userId, "SystemUser", "BusinessUnitId", null,
            function successCallback(record) {
                if (record.BusinessUnitId.Id != null && record.BusinessUnitId.Name == BU) {
                    match = true;
                }
            }, function errorCallback(error) {
                console.log(error);
            });

    return match;
}

//Check logged in user role
function CheckUserRole(roleName) {
    var currentUserRoles = Xrm.Page.context.getUserRoles();
    for (var i = 0; i < currentUserRoles.length; i++) {
        var userRoleId = currentUserRoles[i];
        var userRoleName = GetRoleName(userRoleId);
        if (userRoleName == roleName) {
            return true;
        }
    }
    return false;
}

//Get role name based on RoleId
function GetRoleName(roleId) {
    var odataSelect = Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc" + "/" + "RoleSet?$filter=RoleId eq guid'" + roleId + "'";
    var roleName = null;
    $.ajax(
        {
            type: "GET",
            async: false,
            contentType: "application/json; charset=utf-8",
            datatype: "json",
            url: odataSelect,
            beforeSend: function (XMLHttpRequest) { XMLHttpRequest.setRequestHeader("Accept", "application/json"); },
            success: function (data, textStatus, XmlHttpRequest) {
                roleName = data.d.results[0].Name;
            },
            error: function (XmlHttpRequest, textStatus, errorThrown) { alert('OData Select Failed: ' + textStatus + errorThrown + odataSelect); }
        }
    );
    return roleName;
}

function getValue(fieldName) {
    var field = Xrm.Page.getAttribute(fieldName);

    if (field) {
        return field.getValue();
    }
    else {
        return null;
    }
}

function setValue(fieldName, value) {
    var field = Xrm.Page.getAttribute(fieldName);

    if (field) {
        field.setValue(value);
    }
}

/*
Purpose: Open filtered entity lookup to select record

Parameters
defaultType = Object Type Code of the entity which needs to be shown in the Lookup.
defaultViewId = GUID of the defaultView which will be set for the Look In section of the Lookup Dialog.
controlName: It is the name of the Lookup Control.
currentObjectType = Object Type Code of the current entity.
currentId = GUID of the current record.
rDependentAttr = Attribute on the basis of which lookup is filtered.
rId = It is the GUID of the record selected in the field that filters results (rDependentAttr).
rType = It is the Object Type Code of the field that filters results (rDependentAttr).
relationshipId = Relationship name of how the two entities are related.
*/
function openFilteredLookup(defaultType, defaultViewId, currentObjectType, currentId, rDependentAttr, rId, rType, relationshipId) {
    try {
        var url = "";
        if (currentObjectType && currentId) {
            //prepare lookup url
            //url = "/_controls/lookup/lookupinfo.aspx?AllowFilterOff=0&DefaultType=" + defaultType + "&DefaultViewId=%7b" + defaultViewId + "%7d&DisableQuickFind=0&DisableViewPicker=1&IsInlineMultiLookup=0&LookupStyle=multi&ShowNewButton=1&ShowPropButton=1&browse=false&currentObjectType=" + currentObjectType + "&currentid=%7b" + currentId + "%7d&dType=1&mrsh=false&objecttypes=" + defaultType + "&rDependAttr=" + rDependentAttr + "&rId=%7b" + rId + "%7d&rType=" + rType + "&relationshipid=" + relationshipId + "";
            url = "/_controls/lookup/lookupinfo.aspx?AllowFilterOff=0&DefaultType=" + defaultType + "&DefaultViewId=" + defaultViewId + "&DisableQuickFind=0&DisableViewPicker=0&IsInlineMultiLookup=0&LookupStyle=multi&mc_ServiceActivity=" + rId + "&ShowNewButton=1&ShowPropButton=1&browse=false&currentObjectType=" + currentObjectType + "&currentid=%7b" + currentId + "%7d&dType=1&mrsh=false&objecttypes=" + defaultType;
        }
        else {
            //prepare lookup url
            url = "/_controls/lookup/lookupinfo.aspx?AllowFilterOff=0&DefaultType=" + defaultType + "&DefaultViewId=%7b" + defaultViewId + "%7d&DisableQuickFind=0&DisableViewPicker=1&IsInlineMultiLookup=0&LookupStyle=multi&ShowNewButton=1&ShowPropButton=1&browse=false&dType=1&mrsh=false&objecttypes=" + defaultType + "&rDependAttr=" + rDependentAttr + "&rId=%7b" + rId + "%7d&rType=" + rType + "&relationshipid=" + relationshipId + "";
        }

        // "/_controls/lookup/lookupinfo.aspx?AllowFilterOff=0&DefaultType=1026&DefaultViewId=BE653221-C406-4DD1-BC80-52EC2420BDBA&DisableQuickFind=0&DisableViewPicker=0&IsInlineMultiLookup=0&LookupStyle=multi&PriceLevelId=%7bF6C868F6-473D-E611-80EB-C4346BAC339C%7d&ShowNewButton=1&ShowPropButton=1&browse=false&currentObjectType=3&currentid=%7bF0F8527C-113A-E611-80F0-C4346BACF5C0%7d&dType=1&mrsh=false&objecttypes=1026"
        
        //Set the Dialog Width and Height
        var DialogOptions = new Xrm.DialogOptions();

        //Set the Width
        DialogOptions.width = 500;

        //Set the Height
        DialogOptions.height = 550;

        //open dialog
        Xrm.Internal.openDialog(Mscrm.CrmUri.create(url).toString(), DialogOptions, null, null, function callback() { });
    } catch (e) {
        alert(e);
    }
}

function getObjectTypeCode(entityLogicalName) {
    var objectTypeCode = "";

    if (parent.Mscrm.XrmInternal.getEntityCode != null && parent.Mscrm.XrmInternal.getEntityCode != undefined) {
        objectTypeCode = parent.Mscrm.XrmInternal.getEntityCode(entityLogicalName);
    }
    else {
        objectTypeCode = parent.Mscrm.XrmInternal.prototype.getEntityCode(entityLogicalName);
    }

    return objectTypeCode;
}

function cleanCurlies(id) {
    return id.replace("{", "").replace("}", "");
}

//pseudo threads
function _thread(func) {
    // Render Spin 
    var isBusy = false;
    var processor = setInterval(function () {
        if (!isBusy) {
            if (!isBusy) {
                isBusy = true;
                var renderSuccess = func();
                if (renderSuccess) {
                    clearInterval(processor);
                }
                isBusy = false;
            }
        }
    }, 250);
}