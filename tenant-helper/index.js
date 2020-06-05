function getElementById(id) {
    return document.getElementById(id);
}


async function  copyCriteria() {
    const text = getElementById('tenantQueryText').value;

    const lower = text.trim().toLowerCase();

    const criteria = `WHERE CONTAINS(LOWER(c.TenantName), '${lower}') or CONTAINS(LOWER(c.AzureADPrimaryRealm), '${lower}') or CONTAINS(LOWER(c.NotificationEmails), '${lower}')`
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


