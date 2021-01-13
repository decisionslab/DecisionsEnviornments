function getElementById(id) {
    return document.getElementById(id);
}


async function  copyCriteria() {
    const text = getElementById('tenantQueryText').value;

    const lower = text.trim().toLowerCase();

    const criteria = `WHERE CONTAINS(LOWER(c.TenantName), '${lower}') or CONTAINS(LOWER(c.AzureADPrimaryRealm), '${lower}') or CONTAINS(LOWER(c.VerifiedDomains), '${lower}')`
    console.log(criteria);

    var data = new Blob([criteria], {type : "text/plain"});

    await navigator.clipboard.writeText(criteria).catch(error => {
        console.error(error);
    });

    getElementById('message').innerHTML = `<i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i> Criteria copied...`;

    setTimeout(() => {
        getElementById('message').innerText = "";

    }, 1000);
}

getElementById('copyTenantBtn').addEventListener('click', copyCriteria);

getElementById('tenantQueryText').addEventListener("keyup", function(evt) {

    if (evt.keyCode === 13) {
        window.copyCriteria()
    }
}, false);

async function copyFeatureFlags() {
const featureFlags = 
`
    "FeatureFlags": {
        "enableDecisionVoting": true,
        "useSeparateForTasksAndDecisions": true
    },
`
	
	await navigator.clipboard.writeText(featureFlags).catch(error => {
        console.error(error);
    });
	
	getElementById('FeatureMessage').innerHTML = `<i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i> Feature flags copied...`;

    setTimeout(() => {
        getElementById(FeatureMessage).innerText = "";

    }, 1000);
	
}
	
getElementById('copyFeatureFlagsBtn').addEventListener('click', copyFeatureFlags);


async function  copyCaseSubmissionPlan() {
    const text = getElementById('caseSubmissionQueryText').value;
    
    getElementById('enteredText').innerHTML = text;
    const url = new URL(text.replace('#', ''));

    const groupId = url.searchParams.get("groupId");
    const planId = url.searchParams.get("planId");

    const plan = 
        `
            "CaseSubmissionPlanId": "${planId}",
            "CaseSubmissionPlanOwnerId": "${groupId}",
            "FeatureFlags": {
                "readOnlyPermissionsAfterPublish": true
                },
        `
    console.log(plan);

    var data = new Blob([plan], {type : "text/plain"});

    await navigator.clipboard.writeText(plan).catch(error => {
        console.error(error);
    });

    getElementById('caseSubmissionMessage').innerHTML = `<i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i> Criteria copied...`;

    setTimeout(() => {
        getElementById('caseSubmissionMessage').innerText = "";

    }, 1000);
}

getElementById('copyPlanBtn').addEventListener('click', copyCaseSubmissionPlan);

getElementById('caseSubmissionQueryText').addEventListener("keyup", function(evt) {

    if (evt.keyCode === 13) {
        window.copyCaseSubmissionPlan()
    }
}, false);


// Calender Access functions
getElementById('copyCalendarBtn').addEventListener('click', copyCalendarAccess);

//Creating object for CalenderAccess

async   function  copyCalendarAccess() {
    const tenantId = getElementById('tenantIdText').value;
    const grantedBy = getElementById('grantedByText').value;
    const grantedTo = getElementById('grantedToText').value;

    const calendarAccessObj = 
    `{
    "id": "${grantedBy}_${grantedTo}",
    "GrantedByUserId": "${grantedBy}",
    "GrantedToUserId": "${grantedTo}",
    "GrantedDate":"${new Date().toISOString()}",
    "PartitionKey": "AgendaCreationAccess_${tenantId}",
}`;

    await navigator.clipboard.writeText(calendarAccessObj).catch(error => {
        console.error(error);
    });

    getElementById('calendarAccessMessage').innerHTML = `<i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i> Criteria copied...`;
    getElementById('calendarAccessOutput').innerHTML = calendarAccessObj ;

    setTimeout(() => {
        getElementById('calendarAccessMessage').innerText = "";

    }, 1000);
}