// Create email from template
function CreateContractEmail() {
    var templateName = "Email Template";
    var temps = [];

    // find the template
    SDK.REST.retrieveMultipleRecordsSync("Template", "?$filter=Title eq '" + templateName + "'&$select=Subject,Body", 
        function successCallback(records) {
            temps.push(records);
        },
        function errorCallback(error) {
            console.log("Error retrieving email template:");
            console.log(error);
        },
        function OnComplete() {
            if (temps.length > 0 && temps[0].length > 0) {
                var contractIdClean = Xrm.Page.data.entity.getId().replace("{","").replace("}","");
                var templateSubject = temps[0][0].Subject;
                var templateBody = temps[0][0].Body;

                var contract = null;
                var POs = [];

                // get the contract
                SDK.REST.retrieveRecordSync(contractIdClean, "Ora_engineeringcontract", null, "ora_account_ora_engineeringcontract",
                    function successCallback(record) {
                        contract = record;
                    },
                    function errorCallback(error) {
                        console.log("Error retrieving engineering contract:");
                        console.log(error);
                        return;
                    });

                // get the POs
                SDK.REST.retrieveMultipleRecordsSync("Ora_engineeringpo", 
                    "?$filter=ora_engineeringcontractid/Id eq (guid'" + contractIdClean + "')&$select=Ora_Number",
                    function successCallback(records) {
                        POs.push(records);
                    },
                    function errorCallback(error) {
                        console.log("Error retrieving engineering POs:");
                        console.log(error);
                    },
                    function OnComplete() {
                        var temp = [];

                        for (var i = 0; i < POs.length; i++) {
                            for (var j = 0; j < POs[i].length; j++) {
                                temp.push(POs[i][j].Ora_Number);
                            }
                        }

                        // concatenate all POs into one comma-separated string
                        POs = temp.join(", ");
                    });

                // data comes back formatted and inside brackets <![CDATA[formatted text]]>
                var subjectStartIndex = nth_occurrence(templateSubject, '[', 2) + 1;
                var subjectStopIndex = nth_occurrence(templateSubject, ']', 1);
                var bodyStartIndex = nth_occurrence(templateBody, '[', 2) + 1;
                var bodyStopIndex = nth_occurrence(templateBody, ']', 1);
                var subject = "";
                var body = "";

                if (subjectStartIndex != -1 && subjectStopIndex != -1) {
                    subject = templateSubject.substring(subjectStartIndex, subjectStopIndex);
                }
                
                if (bodyStartIndex != -1 && bodyStopIndex != -1) {
                    body = templateBody.substring(bodyStartIndex, bodyStopIndex);
                }

                // map fields from contract + account to email
                mapValues(contract, subject, body, POs);
            }
        });
}

function mapValues(contract, subject, body, POs) {
    var entLogName = "ora_engineeringcontract"; // entity logical name; used in option set query
    //debugger;
    // subject
    if (contract.Ora_Number) {
        subject = subject.replace("$ora_engineeringcontract.ora_number$", contract.Ora_Number);
    }
    else {
        subject = subject.replace("$ora_engineeringcontract.ora_number$", "");
    }

    // body
    if (contract.ora_accountid.Id != null) {
        body = body.replace("$ora_engineeringcontract.ora_accountidname$", contract.ora_accountid.Name); // client
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_accountidname$", ""); // client
    }
    if (contract.ora_account_ora_engineeringcontract.AccountNumber) {
        body = body.replace("$account.accountnumber$", contract.ora_account_ora_engineeringcontract.AccountNumber); // client account - account number
    }
    else {
        body = body.replace("$account.accountnumber$", ""); // client account - account number
    }
    if (contract.Ora_Number) {
        body = body.replace("$ora_engineeringcontract.ora_number$", contract.Ora_Number); // project
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_number$", ""); // project
    }
    if (contract.Ora_Title) {
        body = body.replace("$ora_engineeringcontract.ora_title$", contract.Ora_Title); // description
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_title$", ""); // description
    }
    if (POs) {
        body = body.replace("$EngineeringContractPOList$", POs); // purchase order
    }
    else {
        body = body.replace("$EngineeringContractPOList$", ""); // purchase order
    }
    if (contract.Ora_Amount.Value != null) {
        body = body.replace("$ora_engineeringcontract.ora_amount$", Number(contract.Ora_Amount.Value).formatMoney(2)); // price
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_amount$", ""); // price
    }
    if (contract.Ora_Class.Value != null) {
        body = body.replace("$ora_engineeringcontract.ora_classname$", getOptionSetText(entLogName, "ora_class", contract.Ora_Class.Value)); // g/c - class label
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_classname$", ""); // g/c - class label
    }
    if (contract.Ora_BillingMethod.Value != null) {
        // shows twice
        body = body.replace(/\$ora_engineeringcontract.ora_billingmethodname\$/g, getOptionSetText(entLogName, "ora_billingmethod", contract.Ora_BillingMethod.Value)); // contract type & billing method - billing method label
    }
    else {
        body = body.replace(/$ora_engineeringcontract.ora_billingmethodname$/g, ""); // contract type & billing method - billing method label
    }
    if (contract.Ora_StartDate) {
        body = body.replace("$ora_engineeringcontract.ora_startdate$", contract.Ora_StartDate.toLocaleDateString()); // start
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_startdate$", ""); // start
    }
    if (contract.Ora_CompletionDate) {
        body = body.replace("$ora_engineeringcontract.ora_completiondate$", contract.Ora_CompletionDate.toLocaleDateString()); // complete
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_completiondate$", ""); // complete
    }
    if (contract.Ora_deliverables) {
        body = body.replace("$ora_engineeringcontract.ora_deliverables$", contract.Ora_deliverables); // deliverables
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_deliverables$", ""); // deliverables
    }
    if (contract.Ora_LineofBusiness.Value) {
        body = body.replace("$ora_engineeringcontract.ora_lineofbusinessname$", getOptionSetText(entLogName, "ora_lineofbusiness", contract.Ora_LineofBusiness.Value)); // line of business - label
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_lineofbusinessname$", ""); // line of business - label
    }
    if (contract.Ora_BillingRatio != null) {
        body = body.replace("$ora_engineeringcontract.ora_billingratio$", parseFloat(Math.round(contract.Ora_BillingRatio * 100) / 100).toFixed(2)); // billing ratio (cbr)
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_billingratio$", ""); // billing ratio (cbr)
    }
    if (contract.Ora_FixedPriceGoal != null) {
        body = body.replace("$ora_engineeringcontract.ora_fixedpricegoal$", parseFloat(Math.round(contract.Ora_FixedPriceGoal * 100) / 100).toFixed(2)); // fixed price goal (fpg)
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_fixedpricegoal$", ""); // fixed price goal (fpg)
    }
    if (contract.Ora_Terms.Value != null) {
        body = body.replace("$ora_engineeringcontract.ora_termsname$", getOptionSetText(entLogName, "ora_terms", contract.Ora_Terms.Value)); // terms - label
    }
    else {
        body = body.replace("$ora_engineeringcontract.ora_termsname$", ""); // terms - label
    }
    if (contract.Ora_engineeringcontractId) {
        var recordLink = Xrm.Page.context.getClientUrl() + "/main.aspx?etn=ora_engineeringcontract&id=" + contract.Ora_engineeringcontractId + "&pagetype=entityrecord";
        body = body.replace("$RecordURL$", '<a href="' + recordLink + '">Click here to open the contract</a>');

        // only create emails for existing contracts
        createEmail(contract, subject, body);
    }
}

function createEmail(contract, subject, body) {
    // build the email
    var email = Object();
    email.Subject = subject;
    email.Description = body;
    email.RegardingObjectId = {
        Id: contract.Ora_engineeringcontractId,
        LogicalName: "ora_engineeringcontract"
    };

    var dueDate = new Date();
    dueDate = new Date(dueDate.setHours(dueDate.getHours() + 7)); // calculation returns milliseconds; must convert to date again
    email.ScheduledEnd = dueDate; // 7 hours from now

    var parties = [];

    // from contract owner
    parties.push(createEmailParty(contract.OwnerId.Id, contract.OwnerId.LogicalName, 1));

    // send to Engineering Contracts contact
    var engConId = findEngineeringContractsContact();
    parties.push(createEmailParty(engConId, "contact", 2));

    // get the engineers
    var engineers = [];
    SDK.REST.retrieveMultipleRecordsSync("Ora_engineer",
        "?$filter=ora_contractengineerid/Id eq (guid'" + contract.Ora_engineeringcontractId + "')&$select=ora_userid",
        function successCallback(records) {
            engineers.push(records);
        },
        function errorCallback(error) {
            console.log("Error retrieving engineers:");
            console.log(error);
        },
        function OnComplete() {
            // also send to related engineers
            for (var i = 0; i < engineers.length; i++) {
                for (var j = 0; j < engineers[i].length; j++) {
                    if (engineers[i][j].ora_userid.Id != null) {
                        parties.push(createEmailParty(engineers[i][j].ora_userid.Id, "systemuser", 2));
                    }
                }
            }
        });

    // takes care of all activity parties, not just from a specific type (from, to)
    email.email_activity_parties = parties;

    SDK.REST.createRecordSync(
        email,
        "Email",
        function successCallback(record) {
            console.log("Created email.");

            // open the email
            Xrm.Utility.openEntityForm("email", record.ActivityId);
        }, function errorCallback(error) {
            console.log("Error creating email:");
            console.log(error);
        });
}

// partyId: recipient's id
// partyType: recipient's entity type
// partyParticipationType: 1 = sender, 2 = recipient
function createEmailParty(partyId, partyType, partyParticipationType) {
    var activityParty = new Object();
    activityParty.PartyId = {
        Id: partyId,
        LogicalName: partyType
    };
    activityParty.ParticipationTypeMask = {
        Value: partyParticipationType
    };

    return activityParty;
}

function findEngineeringContractsContact() {
    var email = "ora_contracts@synopsys.com";
    var contacts = [];
    var contactId = null;

    SDK.REST.retrieveMultipleRecordsSync("Contact", "?$filter=EMailAddress1 eq '" + email + "'&$select=ContactId",
        function successCallback(records) {
            contacts.push(records);
        },
        function errorCallback(error) {
            console.log("Error retrieving engineering contracts contact:");
            console.log(error);
        },
        function OnComplete() {
            if (contacts.length > 0 && contacts[0].length > 0) {
                // use the first result
                contactId = contacts[0][0].ContactId;
            }
            else {
                // create a new contact
                var contact = {
                    FirstName: "Engineering",
                    LastName: "Contracts",
                    EMailAddress1: email
                };
                SDK.REST.createRecordSync(
                    contact,
                    "Contact",
                    function successCallback(record) {
                        console.log("Created contact.");
                        contactId = record.ContactId;
                    }, function errorCallback(error) {
                        console.log("Error creating contact:");
                        console.log(error);
                    });
            }
        });

    return contactId;
}

function nth_occurrence (string, char, nth) {
    var first_index = string.indexOf(char);
    var length_up_to_first_index = first_index + 1;

    if (nth == 1) {
        return first_index;
    } else {
        var string_after_first_occurrence = string.slice(length_up_to_first_index);
        var next_occurrence = nth_occurrence(string_after_first_occurrence, char, nth - 1);

        if (next_occurrence === -1) {
            return -1;
        } else {
            return length_up_to_first_index + next_occurrence;  
        }
    }
}

// c = decimal places, d = dollar & cent separator, t = thousands separator
Number.prototype.formatMoney = function (c, d, t) {
    var n = this,
        c = isNaN(c = Math.abs(c)) ? 2 : c,
        d = d == undefined ? "." : d,
        t = t == undefined ? "," : t,
        s = n < 0 ? "-" : "",
        i = parseInt(n = Math.abs(+n || 0).toFixed(c)) + "",
        j = (j = i.length) > 3 ? j % 3 : 0;
    return s + (j ? i.substr(0, j) + t : "") + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + t) + (c ? d + Math.abs(n - i).toFixed(c).slice(2) : "");
};

function getOptionSetText(entityLogicalName, attrLogicalName, value) {
    var text = "";
    var foundValue = false;

    SDK.Metadata.RetrieveAttributeSync(entityLogicalName, attrLogicalName, null, true,
        function (result) {
            for (var i = 0; i < result.OptionSet.Options.length; i++) {
                if (!foundValue && result.OptionSet.Options[i].Value == value) {
                    text = result.OptionSet.Options[i].Label.LocalizedLabels[0].Label;
                    foundValue = true;
                }
            }
        },
        function (error) {
            console.log("Error retrieving option set text:");
            console.log(error);
        });

    return text;
}