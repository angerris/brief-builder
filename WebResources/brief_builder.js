/* exported
  callGetSPLocationFilesAction,
  checkingSelectedFileds,
  getListOfEmails,
  renderEmails,
  moreDetails,
  getRecordId
  createCard,
  getEmails,
  getClaim,
  search,
  cancel,
  next,
  prev,
*/

let page = 1;
let emails = [];
let sharePointList = [];
let selectedEmailsId = [];
const selectedClaimFieldList = [];
let selectedSharePointItems = [];

function cancel() {
  window.close();
}

window.onload = async function onLoad() {
  try {
    const { recordId, entityName } = getRecordId();
    const emailsData = await getEmails(recordId);
    const claim = await getClaim(recordId, entityName);
    callGetSPLocationFilesAction(recordId);
    main(emailsData, claim);
  } catch (error) {
    alert(error.message);
  }
};

function getRecordId() {
  const decodedParams = decodeURIComponent(window.location.search).split("&");
  const recordId = decodedParams[1].split("id=")[1].toLowerCase();
  const entityName = decodedParams[0].split("entityName=")[1];

  return { recordId, entityName };
}

async function getEmails(recordId) {
  let emailsFetch = `
      <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
        <entity name="email">
          <attribute name="pace_slot_display_name" />
          <attribute name="activityid" />
          <attribute name="subject" />
          <attribute name="from" />
          <attribute name="to" />
          <filter type="and">
            <condition attribute="regardingobjectid" operator="eq" 
            uitype="pace_claim" value="${recordId}" />
          </filter>
        </entity>
      </fetch>   
    `;

  emailsFetch = `?fetchXml=${encodeURIComponent(emailsFetch)}`;
  return await Xrm.WebApi.retrieveMultipleRecords("email", emailsFetch);
}

function main(results, claim) {
  const data = results.entities;
  const container = document.getElementById("container");
  const searchInp = document.querySelector("#search>input");
  const prevBtn = document.getElementById("prev");

  if (data.length) {
    emails = getListOfEmails(data);
    renderEmails(emails);

    searchInp.addEventListener("input", (event) => search(event, emails));
  } else {
    container.innerHTML = `
    <span>No emails available!</span>
    `;
  }
  prevBtn.addEventListener("click", () => prev(emails, claim));
}
function updateNextButtonText() {
  const nextBtn = document.getElementById("next");
  nextBtn.innerText = page === 3 ? "Submit" : "Next";
}

function renderEmails(emails) {
  document.getElementById("container").innerHTML = emails.length
    ? ""
    : "<span>No emails available!</span>";

  for (let i = 0; i < emails.length; i++) {
    const card = createCard(emails[i]);
    document.getElementById("container").appendChild(card);
  }
}

function getListOfEmails(data) {
  const emails = data.map((email) => {
    const emailItem = {
      from: "",
      to: "",
      subject: email.subject || "",
      displayName: email.pace_slot_display_name || "",
      isChecked: false,
      id: email.activityid
    };

    email.email_activity_parties.forEach((item) => {
      if (item.participationtypemask === 1) {
        emailItem.from =
          item["_partyid_value@OData.Community.Display.V1.FormattedValue"] ||
          "";
      } else if (item.participationtypemask === 2) {
        emailItem.to =
          item["_partyid_value@OData.Community.Display.V1.FormattedValue"] ||
          "";
      }
    });
    return emailItem;
  });

  return emails;
}

function createCard(email) {
  const emailCard = document.createElement("div");
  emailCard.className = "email-card";
  emailCard.innerHTML = `
    <input type="checkbox" class="email-checkbox" onclick=''>
    <div class="email-content">
      <div class="email-heading">From: ${email.from} <br> To: ${email.to} </div>
      <div class="email-field">
        <span class="email-label">Subject: ${email.subject}</span>
      </div>
      <div class="email-field">
        <span class="email-label">Display Name: ${email.displayName}</span>
      </div>
    </div>
    <div class="email-icon">
      <i class="fa-solid fa-angle-right"></i>
    </div>
  `;
  emailCard.querySelector(".email-checkbox").checked = email.isChecked;
  emailCard
    .querySelector(".email-checkbox")
    .addEventListener("click", (e) => (email.isChecked = e.target.checked));
  emailCard
    .querySelector(".email-icon")
    .addEventListener("click", () => moreDetails(email.id));

  return emailCard;
}

function moreDetails(emailId) {
  const pageInput = {
    pageType: "entityrecord",
    entityName: `email`,
    entityId: emailId
  };

  const navigationOptions = {
    target: 2, // Opens in a dialog
    position: 2, // Side panel
    width: {
      value: 50,
      unit: "%"
    },
    title: "Email Details"
  };

  Xrm.Navigation.navigateTo(pageInput, navigationOptions);
}

async function next() {
  selectedEmailsId = emails.filter((e) => e.isChecked).map((e) => e.id);
  if (page === 3) {
    await submitAction();
    return;
  }
  const { recordId, entityName } = getRecordId();
  if (page < 3) page++;
  if (page === 1) {
    selectedEmailsId = Array.from(document.querySelectorAll(".email-checkbox"))
      .map((checkbox, index) => ({ checked: checkbox.checked, index }))
      .filter((item) => item.checked)
      .map((item) => {
        const card = document.querySelectorAll(".email-card")[item.index];
        return card ? emails[item.index].id : null;
      })
      .filter((id) => id !== null);
  }
  updateNextButtonText();
  if (page === 2) {
    document.querySelector("#search").style.display = "none";
    document.getElementById("container").innerHTML = "";
    document.getElementById("prev").style.display = "block";

    const claim = await getClaim(recordId, entityName);
    showClaimFileds(claim);
  } else if (page === 3) {
    document.getElementById("container").innerHTML = "";
    showSharePoints(sharePointList);
  }
}

async function getClaim(claimId, entityName) {
  let claimFetch = `
      <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
        <entity name="${entityName}">
          <attribute name="ownerid" />
          <attribute name="pace_fk_insurance_policy" />
          <attribute name="pace_os_claim_closure_type" />
          <attribute name="pace_date_close_date" />
          <filter type="and">
            <condition attribute="pace_claimid" operator="eq" value="${claimId}" />
          </filter>
        </entity>
      </fetch>
    `;

  claimFetch = `?fetchXml=${encodeURIComponent(claimFetch)}`;
  const data = await Xrm.WebApi.retrieveMultipleRecords(entityName, claimFetch);

  if (data.entities.length) {
    const claim = data.entities[0];
    const currentClaim = {
      "Insurance Policy":
        claim?.[
          "_pace_fk_insurance_policy_value@OData.Community.Display.V1.FormattedValue"
        ],
      "Claim Handler":
        claim?.["_ownerid_value@OData.Community.Display.V1.FormattedValue"],
      "Close Date":
        claim?.[
          "pace_date_close_date@OData.Community.Display.V1.FormattedValue"
        ],
      "Claim Closure Type":
        claim?.[
          "pace_os_claim_closure_type@OData.Community.Display.V1.FormattedValue"
        ]
    };

    return currentClaim;
  }
}

function showClaimFileds(claim) {
  document.getElementById("container").innerHTML = `
    <div class='claim-container'></div>
  `;

  for (const [key, value] of Object.entries(claim)) {
    const field = document.createElement("div");
    field.className = "field-block";
    field.innerHTML = `
      <input type="checkbox" class="field-checkbox">
      <div class="field-content">
        <div class="field-label">${key.replace(/_/g, " ")}</div>
        <div class="field-value">${value ? value : ""}</div>
      </div>
    `;
    field
      .querySelector(".field-checkbox")
      .addEventListener("click", () => selectedClaimFields(key, value));
    field.querySelector(".field-checkbox").checked = checkingSelectedFileds(
      key,
      value
    );
    document.querySelector(".claim-container").appendChild(field);
  }
}

function selectedClaimFields(key, value) {
  const index = selectedClaimFieldList.findIndex(
    (item) => Object.keys(item)[0] === key && Object.values(item)[0] === value
  );

  if (index !== -1) {
    selectedClaimFieldList.splice(index, 1);
    return;
  }

  const obj = {};
  obj[key] = value;
  selectedClaimFieldList.push(obj);
}

function checkingSelectedFileds(key, value) {
  return selectedClaimFieldList.some(
    (item) => Object.keys(item)[0] === key && Object.values(item)[0] === value
  );
}

function prev(emails, claim) {
  if (page > 1) page--;
  updateNextButtonText();
  if (page === 1) {
    document.querySelector("#search").style.display = "block";
    document.getElementById("prev").style.display = "none";
    renderEmails(emails);
  } else if (page === 2) {
    document.getElementById("container").innerHTML = "";
    selectedEmailsId = emails
      .filter((item) => item.isChecked)
      .map((item) => item.id);
    showClaimFileds(claim);
  }
}

function search(event, emails) {
  const value = event.target.value.trim();
  const filteredEmails = emails.filter(
    (email) =>
      email.from?.includes(value) ||
      email.to?.includes(value) ||
      email.subject?.includes(value) ||
      email.displayName?.includes(value)
  );

  renderEmails(filteredEmails);
}

function showSharePoints(sharePoints) {
  const container = document.getElementById("container");

  if (!sharePoints.length) {
    container.innerHTML = `<h3>Not Found</h3>`;
  }

  sharePoints.forEach((item) => {
    const card = document.createElement("div");
    card.className = "sharePoint-card";
    card.innerHTML = `
      <input type="checkbox" class="sharePoint-checkbox">
      <div class="sharePoint-info">
        <div class="sharePoint-name">${item.Name}</div>
      </div>
    `;

    card
      .querySelector(".sharePoint-checkbox")
      .addEventListener("click", () => toggleSharePointSelection(item.Id));
    card.querySelector(".sharePoint-checkbox").checked =
      selectedSharePointItems.some((sp) => sp.Id === item.Id);

    container.appendChild(card);
  });
}

function toggleSharePointSelection(sharePointId) {
  const idx = selectedSharePointItems.findIndex((sp) => sp.Id === sharePointId);
  if (idx === -1) {
    const sp = sharePointList.find((item) => item.Id === sharePointId);
    if (sp) selectedSharePointItems.push({ Id: sp.Id, Name: sp.Name });
  } else {
    selectedSharePointItems.splice(idx, 1);
  }
}

//  ------------ call Action to get Share Points files ----------------------------
async function callGetSPLocationFilesAction(claimId) {
  const request = {
    TargetRef: {
      id: claimId,
      entityType: "pace_claim"
    },

    getMetadata: () => {
      const metadata = {
        boundParameter: null,
        parameterTypes: {
          TargetRef: {
            typeName: "mscrm.pace_claim",
            structuralProperty: 5
          }
        },
        operationType: 0, // 0 = Action
        operationName: "pace_GetSPLocationFiles"
      };
      return metadata;
    }
  };

  Xrm.WebApi.online.execute(request).then(
    (response) => {
      if (response.ok) {
        response.json().then((result) => {
          sharePointList = JSON.parse(result.FileList);
        });
      }
    },
    (error) => {
      console.log("Error calling action:", error.message);
    }
  );
}

async function submitAction() {
  const { recordId } = getRecordId();

  const data = {
    recordId: recordId,
    emailIds: selectedEmailsId,
    claims: selectedClaimFieldList,
    sharepointFiles: selectedSharePointItems.map((sp) => ({
      id: sp.Id,
      name: sp.Name
    }))
  };

  const request = {
    data: JSON.stringify(data),

    getMetadata: function () {
      return {
        boundParameter: null,
        parameterTypes: {
          data: {
            typeName: "Edm.String",
            structuralProperty: 1
          }
        },
        operationType: 0,
        operationName: "pace_BriefBuilderGenerator"
      };
    }
  };

  await Xrm.WebApi.online.execute(request);
  window.close();
}
