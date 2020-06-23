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
    const url = getElementById('caseSubmissionQueryText').value;

    console.log(url);
    const firstIndexOfGroupId = url.indexOf('?') + 9;
    const lastIndexOfGroupId = url.indexOf('&');

    const groupId = url.substring(firstIndexOfGroupId, lastIndexOfGroupId);
    const planId = url.substring(lastIndexOfGroupId + 8);
    console.log(groupId);
    console.log(planId);

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