/**
 * This is a Google Apps Script, to be deployed within a Google sheet container.
 * Various sheets are assumed to exist, e.g. "email_sent", "customer_meetings", "customers", "config"
 * Access to Calendar API needs to be added as a service.
 */

/**
 * This is basically adapted from the v1 GAS code, intended to work with a GAS-translated version of the original Python code below
 */


//////
// Globals
//////

let config_map = getConfig()

const common_tags = {
  "call": "624bb7efa5b26c4f53265358",
  "prep_followup": "624bb7d9a5b26c4f53265352",
}


let logged_intervals = get_intervals(config_map["hours"] + 4)

let headers = { 'x-api-key': config_map["clockify_key"] }

let response = jsonResponse(UrlFetchApp.fetch("https://hubspot.clockify.me/api/v1/user", { headers: headers }));
config_map["user_id"] = response["id"];
config_map["workspace_id"] = response["defaultWorkspace"];

let min_adjusted_meeting_length = 0.5 // fraction
let max_meeting_start_delay = 0.33 // fraction

let max_email_minutes = 15 // minutes
let min_email_minutes = 5 // minutes
let max_email_overlap = 3 // minutes, should be smaller than min_email_minutes

//////
// GAS events and UI elements and wrapped functions
//////

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    //{name: 'Validate content (placeholder)', functionName: 'validateSheet_'},
    { name: 'Refresh email and calendar logs', functionName: 'refreshSheet_' },
    { name: 'Write logs to Clockify', functionName: 'logActivities_' },
    { name: 'Get project/client/task IDs', functionName: 'fetchServices_' }
  ];
  spreadsheet.addMenu('Clockifyiable_Activities', menuItems);
}

function refreshSheet_() {
  writeRecentSentEmail();
  writeRecentMeetings();
}

function logActivities_() {
  log_all_activities();
}

function log_all_activities() {
  log_meetings();
  log_email();
}

function fetchServices_() {
  enrich_customers();
}

function enrich_customers() {
  let matched_projects = getServices();
  matchCustomerProjects(matched_projects);
}

//////
// Sheet interaction functions (GAS v1)
//////

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

function getServices() {
  var project_requests = []
  for (var i = 0; i < 5; i++) {
    project_requests.push(
      {
        'url': 'https://hubspot.clockify.me/api/v1/workspaces/' + config_map["workspace_id"] + '/projects?page-size=3000&archived=false&page=' + i.toString() + '&hydrated=true',
        'headers': { 'x-api-key': config_map['clockify_key'] }
      }
    )
  }
  var project_batches = UrlFetchApp.fetchAll(project_requests);
  var projects_by_hid = {};
  for (let element of getHIDs()) {
    projects_by_hid[element] = []
  }
  var task = {}
  for (let outer_index = 0; outer_index < project_batches.length; outer_index++) {
    let batch_of_projects = JSON.parse(project_batches[outer_index].getContentText());
    for (let index = 0; index < batch_of_projects.length; index++) {
      let project = batch_of_projects[index];
      for (let task_index = 0; task_index < project["tasks"].length; task_index++) {
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
  return hid_array;
}

function domainToIds(domains) {
  if (typeof domains === 'string' || domains instanceof String) {
    domains = [domains]
  }
  domains = domains.map(x => x.toLowerCase());
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("customers");
  let range = sheet.getDataRange();
  let values = range.getValues();
  let domain_map = {}
  for (var i = 1; i < values.length; i++) {
    let all_domains = values[i][0].replace(";", ",").toLowerCase().split(",")
    for (let j = 0; j < all_domains.length; j++) {
      let domain = all_domains[j].trim();
      domain_map[domain] = {
        client_id: values[i][3],
        project_id: values[i][4],
        task_id: values[i][5],
        hid: values[i][1]
      }
    }
  }
  let matched_client = "";
  let matched_project = "";
  let matched_task = "";
  let matching_success = false;
  let matched_domains = [];
  for (let domain of domains) {
    if (domain_map[domain]) {
      if (!matching_success) {
        matched_client = domain_map[domain]["client_id"]
        matched_project = domain_map[domain]["project_id"]
        matched_task = domain_map[domain]["task_id"]
        matched_hid = domain_map[domain]["hid"]
        matching_success = true;
        matched_domains.push(domain);
      } else if (domain_map[domain]["project_id"] != matched_project) {
        matching_success = false;
        break;
      } else {
        matched_domains.push(domain);
      }
    }
  }
  if (matching_success && matched_project != "" && matched_task != "") {
    return [matched_client, matched_project, matched_task, matched_hid, matched_domains.join(";")]
  } else {
    return [null, null, null, null, null]
  }
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
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("email_sent");
  sheet.clear();
  sheet.appendRow(["send_timestamp", "subject", "recipient_domains", "client_id", "project_id", "task_id", "hid", "matched_domains"]);
  let threads = GmailApp.search("in:sent", 0, 100);
  for (var i = 0; i < threads.length; i++) {
    let messages = threads[i].getMessages();
    for (var j = messages.length - 1; j >= 0; j--) {
      if (messages[j].getFrom().includes(config_map["sender_email"])) {
        let message_date = messages[j].getDate();
        if ((Date.now() - message_date) / (1000 * 60 * 60) < config_map["hours"]) {
          let message_subject = messages[j].getSubject()
          //TODO make the following exclusion strings part of the config sheet
          if (message_subject && !message_subject.includes("out of office") && !message_subject.includes("slow to respond")) {
            let message_recipients = messages[j].getTo();
            let message_cc = messages[j].getCc();
            if (message_cc.length > 0) {
              message_recipients = message_recipients + ", " + message_cc
            }
            let recipients = extractEmailAddresses(message_recipients);
            let recipient_domains = []
            for (var k = 0; k < recipients.length; k++) {
              let recipient_domain = recipients[k].split("@")[1];
              if (!recipient_domains.includes(recipient_domain) && recipient_domain != "hubspot.com" && recipient_domain != "gmail.com" && !recipient_domain.includes("google.com")) {
                recipient_domains.push(recipient_domain);
              }
            }
            let matchedIds = domainToIds(recipient_domains)
            if (matchedIds[1]) {
              sheet.appendRow([message_date.getTime(), sanitize(message_subject.toLowerCase()), recipient_domains.join(";"), matchedIds[0], matchedIds[1], matchedIds[2], matchedIds[3], matchedIds[4]]);
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
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("customer_meetings");
  sheet.clear();
  sheet.appendRow(["start_timestamp", "end_timestamp", "event_summary", "recipient_domains", "client_id", "project_id", "task_id", "hid", "matched_domains"]);
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
            if (attendee.self) {
              if (attendee.responseStatus == "declined") {
                log_event = false;
              }
            } else {
              let attendee_domain = attendee.email.split("@")[1]
              //TODO also include internal meetings
              if (!event_domains.includes(attendee_domain) && attendee_domain != "hubspot.com" && attendee_domain != "gmail.com" && !attendee_domain.includes("google.com")) {
                event_domains.push(attendee_domain);
              }
            }
          }
        }
        let matchedIds = domainToIds(event_domains);
        if (matchedIds[1] && log_event) {
          let event_start = Date.parse(event.start.dateTime);
          let event_end = Date.parse(event.end.dateTime);
          sheet.appendRow([event_start, event_end, sanitize(event.summary.toLowerCase()), event_domains.join(";"), matchedIds[0], matchedIds[1], matchedIds[2], matchedIds[3], matchedIds[4]]);
        }
      }
    }
  }
}

/**
 * This is basically a GAS rewrite of the original Python code
 */

//////
// Utility functions (v2)
//////

function jsonResponse(response) {
  return JSON.parse(response.getContentText());
}

function get_intervals(minus_x_hours = 96) {
  let lower_bound = Math.floor(Date.now() - minus_x_hours * 60 * 60 * 1000);
  var newDate = new Date();
  newDate.setTime(lower_bound);
  dateString = newDate.toUTCString();
  let page_size = 0;
  let completed = false;
  let intervals = [];
  while (page_size < 1000 && completed == false) {
    page_size += 50;
    let url = "https://hubspot.clockify.me/api/v1/workspaces/" + config_map["workspace_id"] + "/user/" + config_map["user_id"] + "/time-entries?page-size=" + page_size
    let r = UrlFetchApp.fetch(url, { headers: headers });
    my_time_entries = jsonResponse(r);
    for (let time_entry of my_time_entries) {
      time_start = Date.parse(time_entry["timeInterval"]["start"]);
      time_end = Date.parse(time_entry["timeInterval"]["end"]);
      if (time_end > lower_bound) {
        intervals.push([time_start, time_end]);
      } else {
        completed = true;
        break;
      }
    }
  }
  return intervals;
}

function sanitize(description) {
  description = description.replace(/[^a-zA-Z0-9]/g, " ");
  while (description.includes("  ")) {
    description = description.replace("  ", " ");
  }
  return description.trim();
}

function log_activity(from_timestamp, to_timestamp, description, project_id, tag_list, billable, task_id) {
  let from_isoZ = new Date(from_timestamp).toISOString();
  let to_isoZ = new Date(to_timestamp).toISOString();
  let data = {
    "start": from_isoZ,
    "end": to_isoZ,
    "billable": billable,
    "projectId": project_id.toString(),
    "tagIds": tag_list,
    "description": description,
    "taskId": task_id.toString()
  };
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(data),
    'headers': headers
  };
  let response = UrlFetchApp.fetch("https://hubspot.clockify.me/api/v1/workspaces/" + config_map["workspace_id"] + "/time-entries", options);
  if (response.getResponseCode() == 201) {
    return true;
  } else {
    Logger.log("failed to create time entry (" + description + ")");
    return false;
  }
}

function effective_meeting_times(from_timestamp, to_timestamp) {
  from_timestamp = parseInt(from_timestamp);
  to_timestamp = parseInt(to_timestamp);
  skip = false;
  original_length = to_timestamp - from_timestamp;
  latest_start_date = from_timestamp + original_length * max_meeting_start_delay;
  for (var i = 0; i < logged_intervals.length; i++) {
    if (latest_start_date > logged_intervals[i][1] > from_timestamp) {
      from_timestamp = logged_intervals[i][1];
    }
  }
  for (var i = 0; i < logged_intervals.length; i++) {
    if (to_timestamp < logged_intervals[i][0] < to_timestamp) {
      to_timestamp = logged_intervals[i][0];
    }
  }
  if ((to_timestamp - from_timestamp) < original_length * min_adjusted_meeting_length) {
    skip = true;
  }
  for (var i = 0; i < logged_intervals.length; i++) {
    if (logged_intervals[i][0] < to_timestamp && logged_intervals[i][1] > from_timestamp) {
      skip = true;
    }
  }
  if (skip) {
    return [null, null];
  } else {
    return [from_timestamp, to_timestamp];
  }
}

function effective_email_times(send_timestamp) {
  send_timestamp = parseInt(send_timestamp);
  skip = false;
  upper_bound = send_timestamp;
  for (var i = 0; i < logged_intervals.length; i++) {
    if (logged_intervals[i][1] > send_timestamp && logged_intervals[i][0] < upper_bound) {
      upper_bound = logged_intervals[i][0];
    }
  }
  upper_bound = Math.min(upper_bound, send_timestamp);
  lower_bound = upper_bound - max_email_minutes * 60 * 1000;
  for (var i = 0; i < logged_intervals.length; i++) {
    if (logged_intervals[i][0] < upper_bound && logged_intervals[i][1] > lower_bound) {
      lower_bound = logged_intervals[i][1];
    }
  }
  if ((upper_bound - lower_bound) * 60 * 1000 < min_email_minutes) {
    skip = true;
  }
  if ((send_timestamp - upper_bound) * 60 * 1000 > max_email_overlap) {
    skip = true;
  }
  // this loop may be redundant, not sure
  for (var i = 0; i < logged_intervals.length; i++) {
    if (logged_intervals[i][0] < upper_bound && logged_intervals[i][1] > lower_bound) {
      skip = true;
    }
  }
  if (skip) {
    return [null, null];
  } else {
    return [lower_bound, upper_bound];
  }
}

//////
// main interface (v2)
//////

function log_meetings(silent = false, prep_time_max = 0, post_time_max = 0) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("customer_meetings");
  let range = sheet.getDataRange();
  let values = range.getValues();

  // TODO exclude meetings everybody but yourself have declined (optional?)
  // TODO also include internal meetings
  for (var i = 1; i < values.length; i++) {
    var row = {
      "start_timestamp": values[i][0],
      "end_timestamp": values[i][1],
      "project": values[i][5],
      "event_summary": values[i][2],
      "hid": values[i][7],
      "task_id": values[i][6]
    }
    if (row["project"] && row["project"] != "") {
      let from_timestamp, to_timestamp;
      [from_timestamp, to_timestamp] = effective_meeting_times(row['start_timestamp'], row['end_timestamp']);
      if (from_timestamp && to_timestamp && row['project']) {
        var r = log_activity(from_timestamp, to_timestamp, "CALL " + row['event_summary'], row['project'], [common_tags["call"]], true, row['task_id']);
        if (r) {
          if (!silent) {
            Logger.log("Logged call (" + Math.round((to_timestamp - from_timestamp) / (1000 * 60)) + "min) " + "\"" + row['event_summary'] + "\" to " + row['hid'].toString());
          }
          logged_intervals.push([from_timestamp, to_timestamp]);
          // prep_call_time
          let prep_from, prep_to;
          [prep_from, prep_to] = effective_meeting_times(from_timestamp - prep_time_max * 1000 * 60, from_timestamp);
          if (prep_to == from_timestamp && (prep_to - prep_from) / (1000 * 60) > prep_time_max / 2) {
            r = log_activity(prep_from, prep_to, "call_PREP " + row['event_summary'], row['project'], [common_tags["prep_followup"]], true, row['task_id']);
            if (!r) {
              Logger.log("failed to log call_prep for " + row['hid'].toString());
            }
          }
          // post_call_time
          let post_from, post_to;
          [post_from, post_to] = effective_meeting_times(to_timestamp, to_timestamp + post_time_max * 1000 * 60);
          if (post_from == to_timestamp && (post_to - post_from) / (1000 * 60) > post_time_max / 2) {
            r = log_activity(post_from, post_to, "call_POST " + row['event_summary'], row['project'], [common_tags["prep_followup"]], true, row['task_id']);
            if (!r) {
              Logger.log("failed to log post_call for " + row['hid'].toString());
            }
          }
        } else {
          Logger.log("FAILED to log call \"" + row['event_summary'] + "\" to " + row['hid'].toString());
        }
      } else {
        Logger.log("Cannot log call \"" + row['event_summary'] + "\" to " + row['hid'].toString() + " (coincides with logged activity)");
      }
    }
  }
}

function log_email(silent = false) {
  // TODO truncate subject line when logging activity
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("email_sent");
  let range = sheet.getDataRange();
  let values = range.getValues();
  // TODO, exclude meetings everybody but yourself have declined (optional?)
  for (var i = 1; i < values.length; i++) {
    var row = {
      "send_timestamp": values[i][0],
      "project": values[i][4],
      "subject": values[i][1],
      "hid": values[i][6],
      "task_id": values[i][5]
    }
    if (row["project"] && row["project"] != "") {
      let effective_times = effective_email_times(row['send_timestamp']);
      if (effective_times[0] && effective_times[1] && row['project']) {
        var r = log_activity(effective_times[0], effective_times[1], "EMAIL " + row['subject'], row['project'], [common_tags["prep_followup"]], true, row['task_id']);
        if (r) {
          if (!silent) {
            Logger.log("Logged email (" + Math.round((effective_times[1] - effective_times[0]) / (1000 * 60)) + "min) " + "\"" + row['subject'] + "\" to " + row['hid'].toString());
          }
          logged_intervals.push([effective_times[0], effective_times[1]]);
        } else {
          Logger.log("FAILED to log email \"" + row['subject'] + "\" to " + row['hid'].toString());
        }
      } else {
        Logger.log("Cannot log email \"" + row['subject'] + "\" to " + row['hid'].toString() + " (coincides with logged activity)");
      }
    }
  }
}