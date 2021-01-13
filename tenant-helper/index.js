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
getElementById('copyCalenderBtn').addEventListener('click', copyCalenderAccess);

//Creating object for CalenderAccess
function CalendarAccess(tenantId,grantedBy,grantedTo)
{
    this.id = `${grantedBy}_${grantedTo}`;
    this.GrantedByUserId = grantedBy;
    this.GrantedToUserId = grantedTo;
    this.GrantedDate = new Date().toUTCString();
    this.PartitionKey = `AgendaCreationAccess_${tenantId}`

}

//for the javascript object to show output in html
function jsonFormatter(val)
{
    return `
      <div class="mt-2">
      <p>{</p>
      <p> "id": "${val.id}"</p>
      <p> "GrantedByUserId": "${val.GrantedByUserId}"</p>
      <p> "GrantedToUserId": "${val.GrantedToUserId}"</p>
      <p> "GrantedDate": "${val.GrantedDate}"</p>
       <p> "PartitionKey": "${val.PartitionKey}"</p>
       <p>}</p>
     </div>
  `;
}


async function  copyCalendarAccess() {
    const tenantId = getElementById('tenantIdText').value;
    const grantedBy = getElementById('grantedByText').value;
    const grantedTo = getElementById('grantedToText').value;

    const calendarAccessObj = new CalendarAccess(tenantId,grantedBy,grantedTo);
    const calendarAccess = JSON.stringify(calendarAccessObj);
    console.log(calendarAccessObj);
    
    await navigator.clipboard.writeText(calendarAccess).catch(error => {
        console.error(error);
    });

    getElementById('calendarAccessMessage').innerHTML = `<i class="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i> Criteria copied...`;
    getElementById('calendarAccessOutput').innerHTML = jsonFormatter(calendarAccessObj) ;

    setTimeout(() => {
        getElementById('calendarAccessMessage').innerText = "";

    }, 1000);
}