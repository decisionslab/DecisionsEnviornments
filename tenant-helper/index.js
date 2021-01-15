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


//Calander Access event binding
(function bindCalendarAccessEvent()
{
   const inputNodes = [ 
        getElementById('tenantIdText'),
        getElementById('grantedByText'),
        getElementById('grantedToText')
    ];
    
    inputNodes.forEach( item =>{
    
        item.addEventListener('keypress', e =>{
            if (e.key === 'Enter') {
                copyCalendarAccess();
              }

        });
    });

})();


// Calendar Access Integrity Checking
function isValidCalendar(tenantIdText,grantedByText,grantedToText)
{
    const error = getElementById('calenderAccessError');
    const tenantIdError = getElementById('tenantIdTextError');
    const grantedByError = getElementById('grantedByTextError');
    const grantedToError = getElementById('grantedToTextError');
    
    const icon = `<i class="ms-Icon ms-Icon--Error" aria-hidden="true"></i>`;

    error.innerHTML = '';
    tenantIdError.innerHTML ='';
    grantedByError.innerHTML = '';
    grantedToError.innerHTML = '';
    
    tenantId = safeTrim(tenantIdText.value);
    grantedBy = safeTrim(grantedByText.value);
    grantedTo = safeTrim(grantedToText.value);
   
    tenantIdText.classList.remove("border","border-solid","border-red-400");
    grantedByText.classList.remove("border","border-solid","border-red-400");
    grantedToText.classList.remove("border","border-solid","border-red-400");


    if(tenantId ==='')
    {
        tenantIdError.innerHTML= `${icon}  Tenant Id is required`;
        tenantIdText.classList.add("border","border-solid","border-red-400");
        tenantIdText.focus();
        return false;
    }else if (grantedBy ==='')
    {
        grantedByError.innerHTML= `${icon} Granted By is required`;
        grantedByText.classList.add("border","border-solid","border-red-400");
        grantedByText.focus();
        return false;
    }
    else if(grantedTo ==='')
    {
        grantedToError.innerHTML= `${icon} Granted To is required`;
        grantedToText.classList.add("border","border-solid","border-red-400");
        grantedToText.focus();
        return false;
    }
    else if(tenantId===grantedBy)
    {   
        error.innerHTML= `${icon} Granted By and Tenant Id must be Unique`;
       
        tenantIdText.classList.add("border","border-solid","border-red-400");
        grantedByText.classList.add("border","border-solid","border-red-400");
        grantedByText.focus();
        return false;
    
    }else if ( tenantId===grantedTo)
    {
        error.innerHTML= `${icon} Granted To and Tenant Id must be Unique`;
        tenantIdText.classList.add("border","border-solid","border-red-400");
        grantedToText.classList.add("border","border-solid","border-red-400");
        grantedToText.focus();

        return false;
    
    }else if(grantedBy===grantedTo)
    {
        error.innerHTML= `${icon} Granted To and Granted By must be Unique`;
        grantedToText.classList.add("border","border-solid","border-red-400");
        grantedByText.classList.add("border","border-solid","border-red-400");
        grantedToText.focus();

        return false;
    }


    return true;

}



async   function  copyCalendarAccess() {
    const tenantIdText = getElementById('tenantIdText');
    const grantedByText = getElementById('grantedByText');
    const grantedToText = getElementById('grantedToText');

    //if input is not valid return
    if(!isValidCalendar(tenantIdText,grantedByText,grantedToText)) return;

    // else construct json
    const tenantId = safeTrim(tenantIdText.value);
    const grantedBy = safeTrim(grantedByText.value);
    const grantedTo = safeTrim(grantedToText.value);


    const calendarAccessObj = 
    `{
    "id": "${grantedBy}_${grantedTo}",
    "GrantedByUserId": "${grantedBy}",
    "GrantedToUserId": "${grantedTo}",
    "GrantedDate":"${new Date().toISOString()}",
    "PartitionKey": "AgendaCreationAccess_${(tenantId)}",
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


function safeTrim(text) {
    if(!text) {
        return '';
    }
    return text.trim();
}
