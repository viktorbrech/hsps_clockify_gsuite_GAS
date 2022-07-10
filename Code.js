/**
 * This is a Google Apps Script, to be attached to a Google sheet (no GCP project necessary).
 * Various sheets are assumed to exist: "email_sent", "customer_meetings", "customers", "config", "sku_prioritization"
 * Access to Calendar API needs to be added as a service.
 */

let config_map = getConfig()

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    //{name: 'Validate content (placeholder)', functionName: 'validateSheet_'},
    { name: 'Refresh email and calendar data', functionName: 'refreshSheet_' },
    { name: 'Get project/client/task IDs', functionName: 'fetchServices_' }
  ];
  spreadsheet.addMenu('Clockifyiable_Activities', menuItems);
}
function refreshSheet_() {
  writeRecentSentEmail();
  writeRecentMeetings();
}
function fetchServices_() {
  enrich_customers();
}
function enrich_customers() {
  let matched_projects = getServices();
  matchCustomerProjects(matched_projects);
}

function getServices() {
  var project_requests = []
  for (var i = 0; i < 5; i++) {
    project_requests.push(
      {
        'url': 'https://hubspot.clockify.me/api/v1/workspaces/' + config_map['clockify_workspace_id'] + '/projects?page-size=3000&archived=false&page=' + i.toString() + '&hydrated=true',
        'headers': { 'x-api-key': config_map['clockify_key'] }
      }
    )
  }
  var project_batches = UrlFetchApp.fetchAll(project_requests);
  var projects_by_hid = {};
  for (const element of getHIDs()) {
    projects_by_hid[element] = []
  }
  var task = {}
  for (outer_index = 0; outer_index < project_batches.length; outer_index++) {
    let batch_of_projects = JSON.parse(project_batches[outer_index].getContentText());
    for (index = 0; index < batch_of_projects.length; index++) {
      let project = batch_of_projects[index];
      for (task_index = 0; task_index < project["tasks"].length; task_index++) {
        task = project["tasks"][task_index]
        if (task["name"] in projects_by_hid && !projects_by_hid[task["name"]].map(x => (x["project"] == project["id"])).some(x => x)) {
          projects_by_hid[task["name"]].push({
            "client": project["clientId"],
            "sku": project["name"],
            "project": project["id"],
            "task": task["id"]
          })
        }
      }
    }
  }
  return projects_by_hid;
}

function getPriorityMap() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("sku_prioritization");
  let range = sheet.getDataRange();
  let values = range.getValues();
  let prio_col = undefined
  switch (config_map["role"]) {
    case "TC":
      prio_col = 1; break;
    case "IC":
      prio_col = 2; break;
    case "CT":
      prio_col = 3; break;
    case "ONB":
      prio_col = 4; break;
  }
  service_priorities = {};
  for (var i = 1; i < values.length; i++) {
    if (values[i][prio_col] != "") {
      service_priorities[values[i][0]] = values[i][prio_col];
    }
  }
  return service_priorities;
}

function matchCustomerProjects(matched_projects) {
  let service_priorities = getPriorityMap();
  ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getSheetByName("customers");
  range = sheet.getDataRange();
  values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    let hid = Math.trunc(values[i][1]).toString()
    let identified = false;
    let chosen_index = undefined
    for (var j = 0; j < matched_projects[hid].length; j++) {
      if (service_priorities.hasOwnProperty(matched_projects[hid][j]["sku"])) {
        let priority = service_priorities[matched_projects[hid][j]["sku"]]
        if (typeof chosen_index == "undefined" || priority > service_priorities[matched_projects[hid][chosen_index]["sku"]]) {
          chosen_index = j
          identified = true;
        } else if (priority == service_priorities[matched_projects[hid][chosen_index]["sku"]]) {
          identified = false;
          break;
        }
      }
    }
    if (identified) {
      values[i][2] = matched_projects[hid][chosen_index]["sku"];
      values[i][3] = matched_projects[hid][chosen_index]["client"];
      values[i][4] = matched_projects[hid][chosen_index]["project"];
      values[i][5] = matched_projects[hid][chosen_index]["task"];
      values[i][6] = "";
    } else {
      values[i][2] = "";
      values[i][3] = "";
      values[i][4] = "";
      values[i][5] = "";
      values[i][6] = JSON.stringify(matched_projects[hid]);
    }
  }
  range.setValues(values);
}

function getHIDs() {
  let hid_array = []
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("customers");
  // This represents ALL the data
  let range = sheet.getDataRange();
  let values = range.getValues();
  // This logs the spreadsheet in CSV format with a trailing comma
  for (var i = 1; i < values.length; i++) {
    hid_array.push(Math.trunc(values[i][1]).toString())
  }
  Logger.log(hid_array);
  return hid_array;
}

function getConfig() {
  let config_map = {}
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("config");
  // This represents ALL the data
  var range = sheet.getDataRange();
  var values = range.getValues();
  // This logs the spreadsheet in CSV format with a trailing comma
  for (var i = 0; i < values.length; i++) {
    config_map[values[i][0]] = values[i][1];
  }
  return config_map
}

function extractEmailAddresses(string) {
  // via https://www.weirdgeek.com/2019/10/regular-expression-in-google-apps-script/ and https://stackoverflow.com/questions/42407785/regex-extract-email-from-strings
  // cf. https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/match
  var regExp = new RegExp("([a-zA-Z0-9+._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)", "gi");
  var results = string.match(regExp);
  return results;
}

function writeRecentSentEmail() {
  // https://developers.google.com/apps-script/reference/gmail
  //getConfig();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("email_sent");
  sheet.clear();
  sheet.appendRow(["send_timestamp", "subject", "recipient_domains"]);
  let threads = GmailApp.search("in:sent", 0, 100);
  for (var i = 0; i < threads.length; i++) {
    let messages = threads[i].getMessages();
    for (var j = messages.length - 1; j >= 0; j--) {
      if (messages[j].getFrom().includes(config_map["sender_email"])) {
        let message_date = messages[j].getDate();
        if ((Date.now() - message_date) / (1000 * 60 * 60) < config_map["hours"]) {
          let message_subject = messages[j].getSubject()
          if (message_subject && !message_subject.includes("out of office") && !message_subject.includes("slow to respond")) {
            let message_recipients = messages[j].getTo();
            let message_cc = messages[j].getCc();
            if (message_cc.length > 0) {
              message_recipients = message_recipients + ", " + message_cc
            }
            let recipients = extractEmailAddresses(message_recipients);
            let recipient_domains = []
            for (var k = 0; k < recipients.length; k++) {
              recipient_domain = recipients[k].split("@")[1];
              if (!recipient_domains.includes(recipient_domain) && recipient_domain != "hubspot.com" && recipient_domain != "gmail.com" && !recipient_domain.includes("google.com")) {
                recipient_domains.push(recipient_domain);
              }
            }
            if (recipient_domains.length > 0) {
              sheet.appendRow([message_date.getTime(), message_subject, recipient_domains.join(";")]);
            }
          }
        }
      }
    }
  }
}

function writeRecentMeetings() {
  // https://developers.google.com/apps-script/guides/services/advanced
  // https://developers.google.com/calendar/api/v3/reference/events 
  // unfortunately couldn't use https://developers.google.com/apps-script/reference/calendar/calendar-app since it doesn't return "decline" status for an event owner
  //getConfig();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("customer_meetings");
  sheet.clear();
  sheet.appendRow(["start_timestamp", "end_timestamp", "event_summary", "recipient_domains"]);
  let calendarId = 'primary';
  let now = new Date();
  let now_minus_one_day = new Date(now.getTime() - (config_map["hours"] * 60 * 60 * 1000));
  let events = Calendar.Events.list(calendarId, {
    timeMin: now_minus_one_day.toISOString(),
    timeMax: now.toISOString(),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 100
  });
  if (events.items && events.items.length > 0) {
    for (var i = 0; i < events.items.length; i++) {
      let event = events.items[i];
      if (!event.start.date) {
        log_event = true;
        let event_domains = []
        if (event.attendees && event.attendees.length > 0) {
          for (var k = 0; k < event.attendees.length; k++) {
            let attendee = event.attendees[k]
            //Logger.log(attendee.email);
            if (attendee.self) {
              if (attendee.responseStatus == "declined") {
                log_event = false;
              }
            } else {
              let attendee_domain = attendee.email.split("@")[1]
              if (!event_domains.includes(attendee_domain) && attendee_domain != "hubspot.com" && attendee_domain != "gmail.com" && !attendee_domain.includes("google.com")) {
                event_domains.push(attendee_domain);
              }
            }
          }
        }
        if (event_domains.length == 0) {
          log_event = false
        }
        if (log_event) {
          let event_start = Date.parse(event.start.dateTime);
          let event_end = Date.parse(event.end.dateTime);
          sheet.appendRow([event_start, event_end, event.summary, event_domains.join(";")]);
        }
      }
    }
  }
}