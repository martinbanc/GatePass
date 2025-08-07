const SPREADSHEET_ID = '1KKuJ3uYkzpZsPRt0dwge0vNQDyXIJCsFWTpI2RPqj98';
const RESPONSES_SHEET_NAME = 'Responses';
const QUESTIONS_SHEET_NAME = 'Questions';
const QUESTIONS_ENTERING_SHEET_NAME = 'Questions_Entering';
const USERS_SHEET = 'Users';
const SITES_SHEET = 'Sites';
const LOCATIONS_SHEET_NAME = 'Locations';
const LOGS_SHEET = 'Action Logs';
const DYNAMIC_COL_ANSWER_SUFFIX = ' - Answer';
const MAX_PREVIEW_FILE_SIZE_BYTES = 1 * 1024 * 1024;

const ROLES = {
  SUPER_USER: 'Super user',
  REGIONAL_MANAGER: 'Regional Manager',
  SITE_SUPERVISOR: 'Site Supervisor',
  SECURITY_OFFICER: 'Security Officer'
};
const STATUS_OPTIONS = ['Pending', 'Approved', 'Rejected', 'Closed'];

/**
 * Gets site data, including new APM and timestamp fields.
 * Now fetches columns A to E.
 */
function getSites_internal() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'SITES_DATA';
  const cached = cache.get(cacheKey);
  if (cached != null) {
    return JSON.parse(cached);
  }
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SITES_SHEET);
  if (sheet.getLastRow() < 2) return [];
  const sites = sheet.getRange("A2:E" + sheet.getLastRow()).getValues().map(row => ({
    siteId: row[0],
    siteName: row[1],
    apmEmail1: row[2],
    apmEmail2: row[3],
    approverListLastUpdated: row[4]
  }));
  cache.put(cacheKey, JSON.stringify(sites), 300);
  return sites;
}

function getUsers_internal(currentUser) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ALL_USERS_DATA';
  let allUsers;
  const cached = cache.get(cacheKey);
  if (cached != null) {
    allUsers = JSON.parse(cached);
  } else {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    if (sheet.getLastRow() < 2) {
      allUsers = [];
    } else {
      allUsers = sheet.getRange("A2:G" + sheet.getLastRow()).getValues().map(row => ({
        username: row[0],
        role: row[2],
        siteId: row[3],
        fullName: row[4],
        position: row[5],
        isActive: row[6]
      }));
    }
    cache.put(cacheKey, JSON.stringify(allUsers), 300);
  }
  if (currentUser.role === ROLES.SITE_SUPERVISOR) {
    return allUsers.filter(user => user.siteId === currentUser.siteId);
  }
  return allUsers;
}

function getApprovers_internal(currentUser) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ALL_APPROVERS_DATA';
  let allApprovers;
  const cached = cache.get(cacheKey);
  if (cached != null) {
    allApprovers = JSON.parse(cached);
  } else {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Approvers');
    if (sheet.getLastRow() < 2) {
      allApprovers = [];
    } else {
      allApprovers = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().map(row => ({
        approverId: row[0],
        fullName: row[1],
        email: row[2],
        siteId: row[3]
      }));
    }
    cache.put(cacheKey, JSON.stringify(allApprovers), 300);
  }
  if (currentUser.role === ROLES.SITE_SUPERVISOR) {
    return allApprovers.filter(appr => appr.siteId === currentUser.siteId);
  }
  return allApprovers;
}

/**
 * NEW: Helper function to get APM emails for a given site.
 * @param {string} siteId The ID of the site.
 * @returns {string} A comma-separated string of APM emails.
 */
function getApmEmailsForSite_(siteId) {
  if (!siteId) return '';
  const allSites = getSites_internal();
  const site = allSites.find(s => s.siteId === siteId);
  if (!site) return '';
  const emails = [];
  if (site.apmEmail1 && site.apmEmail1.trim() !== '') emails.push(site.apmEmail1.trim());
  if (site.apmEmail2 && site.apmEmail2.trim() !== '') emails.push(site.apmEmail2.trim());
  return emails.join(',');
}


/**
 * REFACTORED: Sends an email using a more flexible options object.
 * @param {object} options Email options, including to, subject, htmlBody, and optional cc.
 */
function sendAppEmail(options) {
  try {
    const mailOptions = {
      to: options.to,
      subject: options.subject,
      htmlBody: options.htmlBody,
      name: 'Approva Gate Pass',
      noReply: true
    };
    if (options.cc && options.cc.trim() !== '') {
      mailOptions.cc = options.cc;
    }
    MailApp.sendEmail(mailOptions);
  } catch (e) {
    console.error(`Failed to send email to ${options.to}: ${e.toString()}`);
    logAction('SYSTEM_ERROR', 'EMAIL_FAILURE', `Failed to send email to ${options.to}. Subject: ${options.subject}. Error: ${e.message}`);
  }
}


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gatehouse Admin')
    .addItem('Setup Project Sheets', 'setupProject')
    .addItem('Sync Users to Properties', 'syncUsersToPropertiesService')
    .addSeparator()
    .addItem('Archive System Logs Now', 'archiveLogs')
    .addToUi();
  ui.createMenu('Gate Pass Admin')
    .addItem('Setup/Update Response Headers', 'setupResponseSheetHeaders')
    .addToUi();
}

function getAuthenticatedUser(token) {
  if (!token) return null;
  const sessionDataString = CacheService.getScriptCache().get(token);
  if (!sessionDataString) return null;
  CacheService.getScriptCache().put(token, sessionDataString, 3600);
  return JSON.parse(sessionDataString);
}

function doGet(e) {
  try {
    const favicon = "https://appscript-cdn.co.uk/GatePassProject/approva/Approva_favicon.png";
    const title = "Approva";
    const page = e.parameter.page;
    if (page === 'redirect' && e.parameter.rt) {
      const redirectToken = e.parameter.rt;
      const sessionToken = CacheService.getScriptCache().get(`rt_${redirectToken}`);

      if (sessionToken) {
        // Remove the one‑time redirect token.
        CacheService.getScriptCache().remove(`rt_${redirectToken}`);

        // Build URLs for automatic and fallback navigation.
        const baseUrl = ScriptApp.getService().getUrl();
        const landingUrl = baseUrl + '?page=Landing';
        const fallbackUrl = `${landingUrl}&token=${sessionToken}`;

        // HTML auto‑stores the session token and redirects to the Landing page.
        const html = [
          '<!DOCTYPE html>',
          '<html>',
          '  <head>',
          '    <base target="_top">',
          '    <script>',
          '      (function() {',
          '        try {',
          `          sessionStorage.setItem('sessionToken', '${sessionToken}');`,
          `          window.top.location.replace('${landingUrl}');`,
          '        } catch (err) {',
          `          window.top.location.href = '${fallbackUrl}';`,
          '        }',
          '      })();',
          '    </script>',
          '  </head>',
          '  <body>',
          '    <p>Redirecting to application…</p>',
          `    <p>If you are not redirected, <a href="${fallbackUrl}">click here</a>.</p>`,
          '  </body>',
          '</html>'
        ].join('\n');

        return HtmlService.createHtmlOutput(html)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } else {
        return HtmlService.createHtmlOutput('<h1>Redirect Expired</h1><p>Your redirect token was invalid or expired. Please close this tab and log in again.</p>');
      }
    }
    const locationIdentifier = e.parameter.location;
    if (page === 'approve') {
      const token = e.parameter.token;
      if (!token) return HtmlService.createHtmlOutput("<h1>Error</h1><p>Approval token is missing.</p>");
      try {
        const template = HtmlService.createTemplateFromFile('ApprovalPage');
        template.data = getApprovalTaskDetails(token);
        return template.evaluate().setTitle('Review Gate Pass Request').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
      } catch (err) {
        return HtmlService.createHtmlOutput(`<h1>Error</h1><p>${err.message}</p>`);
      }
    }
    if (locationIdentifier) {
      const locationInfo = getAndValidateLocation(locationIdentifier);
      if (!locationInfo) {
        return HtmlService.createHtmlOutput(`<h1>Error</h1><p>Invalid location code "<strong>${locationIdentifier}</strong>". Please check the URL.</p>`);
      }
      const formType = e.parameter.formType;
      if (!formType) {
        const template = HtmlService.createTemplateFromFile('Welcome');
        template.location = locationInfo;
        template.webAppUrl = ScriptApp.getService().getUrl();
        return template.evaluate().setTitle('Approva Gate Pass - Welcome').setFaviconUrl(favicon);
      }
      const template = HtmlService.createTemplateFromFile('GatePassForm');
      template.location = locationInfo;
      template.formType = formType;
      template.data = getDataForApp(formType, locationInfo.name);
      return template.evaluate().setTitle('Approva Gate Pass Form').setFaviconUrl(favicon).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
    }
    if (page === 'approveReset') {
      return HtmlService.createHtmlOutput(processResetApproval(e.parameter.token));
    }
    let template;
    switch (page) {
      case 'Admin':
        template = HtmlService.createTemplateFromFile('Admin');
        break;
      case 'Logs':
        template = HtmlService.createTemplateFromFile('Logs');
        break;
      case 'MaterialArchive':
        template = HtmlService.createTemplateFromFile('MaterialArchive');
        break;
      case 'Requests':
        template = HtmlService.createTemplateFromFile('Requests');
        break;
      case 'Landing':
      // Simply serve the template file without passing any data to it.
      // The client-side script will handle everything.
      template = HtmlService.createTemplateFromFile('Landing');
      break;
      default: // This handles the Login page
        template = HtmlService.createTemplateFromFile('Login');
        if (e.parameter.timedout === 'true') {
          template.message = 'Your session has expired. Please log in again.';
        } else {
          template.message = '';
        }
        break;
    }
    return template.evaluate().setTitle(title).setFaviconUrl(favicon).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  } catch (err) {
    const errorOutput = `<h1>An error occurred</h1><p>Could not load the application. Please contact an administrator.</p><pre style="white-space: pre-wrap; word-wrap: break-word;">${err.toString()}\n\n${err.stack}</pre>`;
    return HtmlService.createHtmlOutput(errorOutput);
  }
}

function userLogin(credentials) {
    try {
        const user = getUserFromProperties_(credentials.username);
        if (!user || !verifyPassword_(credentials.password, user.PasswordHash)) {
            logAction(credentials.username, 'LOGIN_FAILURE', 'Invalid credentials.');
            return {
                success: false,
                message: 'Invalid username or password.'
            };
        }
        if (user.IsTempPassword) {
            logAction(user.Username, 'LOGIN_TEMP_PASSWORD', 'User logged in with temporary password.');
            return {
                success: true,
                forceChange: true,
                username: user.Username
            };
        }

        const sessionToken = Utilities.getUuid();
        const sessionData = {
            username: user.Username,
            role: user.Role,
            siteId: user.SiteID,
            fullName: user.FullName,
            position: user.Position
        };
        CacheService.getScriptCache().put(sessionToken, JSON.stringify(sessionData), 3600);

        const redirectToken = Utilities.getUuid();
        CacheService.getScriptCache().put(`rt_${redirectToken}`, sessionToken, 60); // Store for 60 seconds

        logAction(user.Username, 'LOGIN_SUCCESS', `Role: ${user.Role}`);
        
        return {
            success: true,
            forceChange: false,
            redirectToken: redirectToken 
        };
    } catch (e) {
        // This block will catch any unexpected errors and report them back
        console.error('userLogin failed with error: ' + e.toString());
        logAction(credentials.username, 'LOGIN_FAILURE', 'System error: ' + e.message);
        return {
            success: false,
            message: 'A system error occurred. Please contact an administrator.'
        };
    }
}

function forceChangePassword(data) {
    const user = findUserRow_(data.username);
    if (!user) throw new Error("User not found.");
    if (user.IsTempPassword !== true) {
        throw new Error("This user does not have a temporary password flag set.");
    }
    if (data.newPassword.length < 8) {
        throw new Error("Password must be at least 8 characters long.");
    }
    const newPasswordHash = hashPassword_(data.newPassword);
    const usersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
    usersSheet.getRange(user.rowIndex, 2).setValue(newPasswordHash);
    usersSheet.getRange(user.rowIndex, 9).setValue(false);
    logAction(data.username, 'FORCE_PASSWORD_CHANGE', 'User successfully changed their temporary password.');
    syncUsersToPropertiesService();
    
    const loginResult = userLogin({
        username: data.username,
        password: data.newPassword
    });

    if (loginResult.success && loginResult.redirectToken) {
        return {
            success: true,
            message: 'Password updated successfully.',
            redirectToken: loginResult.redirectToken
        };
    }
    return {
        success: false,
        message: 'Password updated, but auto-login failed. Please log in manually.'
    };
}

function userLogout(token) {
  if (!token) return;
  const sessionDataString = CacheService.getScriptCache().get(token);
  if (sessionDataString) {
    const session = JSON.parse(sessionDataString);
    logAction(session.username, 'LOGOUT', 'User logged out.');
    CacheService.getScriptCache().remove(token);
  }
}

function getLandingPageData(token) {
  const user = getAuthenticatedUser(token);
  if (!user) {
    throw new Error('Authentication failed. Please log in again.');
  }
  return {
    user: user,
    ROLES: ROLES
  };
}

function getAdminPageData(token) {
  const user = getAuthenticatedUser(token);
  checkAdminPermissions_(user);
  const sites = getSites_internal();
  const users = getUsers_internal(user);
  const approvers = getApprovers_internal(user);
  return {
    sites,
    users,
    approvers,
    currentUser: user
  };
}

function getSites(token) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  return getSites_internal();
}

function getUsers(token) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  return getUsers_internal(currentUser);
}

function getApprovers(token) {
  const user = getAuthenticatedUser(token);
  checkAdminPermissions_(user);
  return getApprovers_internal(user);
}



function getRequestsAndTimeline(token, statusFilter) { // <-- searchTerm parameter removed
  const user = getAuthenticatedUser(token);
  if (!user) throw new Error('Your session has expired. Please refresh the page to log in again.');
  const cache = CacheService.getScriptCache();
  const cacheKey = `requests_timeline_${user.username}_${statusFilter}`;

  try {
    // Check cache first
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      return {
        success: true,
        data: JSON.parse(cachedData)
      };
    }

    // If not in cache, build from scratch
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responsesSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
    if (!responsesSheet || responsesSheet.getLastRow() < 2) return {
      success: true,
      data: []
    };

    const responseData = responsesSheet.getDataRange().getValues();
    const responseHeaders = responseData.shift();
    const idCol = responseHeaders.indexOf('Submission_ID');
    const statusCol = responseHeaders.indexOf('Status');
    const locationCol = responseHeaders.indexOf('Location');
    const timestampCol = responseHeaders.indexOf('Timestamp');
    if ([idCol, statusCol, locationCol, timestampCol].includes(-1)) {
      throw new Error("A required column could not be found.");
    }

    const logsSheet = ss.getSheetByName(LOGS_SHEET);
    const logData = logsSheet.getLastRow() < 2 ? [] : logsSheet.getRange(2, 1, logsSheet.getLastRow() - 1, 4).getValues();
    const logsBySubmissionId = new Map();
    logData.forEach(logRow => {
      const details = logRow[3];
      const submissionIdMatch = details ? String(details).match(/[A-Z]{3}-\d{4,}/) : null;
      if (submissionIdMatch && logRow[0] instanceof Date) {
        const submissionId = submissionIdMatch[0];
        if (!logsBySubmissionId.has(submissionId)) {
          logsBySubmissionId.set(submissionId, []);
        }
        let event = logRow[2];
        if (event === 'UPDATE_TICKET_STATUS') event = `Status changed to '${details.split('->')[1]?.trim()}'`;
        else if (event === 'SAVE_RESPONSE') event = 'Request Submitted';
        else if (event === 'PROCESS_APPROVAL') event = `Request Processed`;
        logsBySubmissionId.get(submissionId).push({
          rawTimestamp: logRow[0],
          user: logRow[1],
          event
        });
      }
    });

    const activeStatuses = ['Pending', 'Approved', 'Active', 'Lodged'];
    const completedStatuses = ['Rejected', 'Closed'];
    const targetStatuses = (statusFilter === 'active') ? activeStatuses : completedStatuses;
    const uniqueRequests = new Map();
    responseData.forEach(row => {
      const submissionId = row[idCol];
      if (!row || row.length === 0 || !submissionId) return;
      const isCorrectSite = (user.siteId === 'all' || row[locationCol] === user.siteId);
      const hasCorrectStatus = targetStatuses.includes(row[statusCol]);
      if (isCorrectSite && hasCorrectStatus && !uniqueRequests.has(submissionId)) {
        uniqueRequests.set(submissionId, row);
      }
    });

    const sortedRequests = Array.from(uniqueRequests.values()).sort((a, b) => {
      const dateA = a[timestampCol] instanceof Date ? a[timestampCol].getTime() : 0;
      const dateB = b[timestampCol] instanceof Date ? b[timestampCol].getTime() : 0;
      return dateB - dateA;
    });

    let processedData = sortedRequests.map(row => {
      const submissionId = row[idCol];
      const nameCol = responseHeaders.indexOf("Driver's Full Name - Answer");
      const approverCol = responseHeaders.indexOf('Shell Approver - Answer');
      const fullDetails = {};
      responseHeaders.forEach((header, i) => {
        if (header && row[i] != null && row[i] !== '') {
          fullDetails[header] = row[i] instanceof Date ? Utilities.formatDate(row[i], "Europe/London", "dd/MM/yyyy HH:mm:ss") : row[i];
        }
      });
      const timelineEvents = logsBySubmissionId.get(submissionId) || [];
      timelineEvents.sort((a, b) => a.rawTimestamp.getTime() - b.rawTimestamp.getTime());
      const formattedTimeline = timelineEvents.map(item => ({
        timestamp: Utilities.formatDate(item.rawTimestamp, "Europe/London", "dd/MM/yyyy HH:mm:ss"),
        user: item.user,
        event: item.event
      }));
      return {
        id: submissionId,
        status: row[statusCol] || 'N/A',
        date: row[timestampCol] instanceof Date ? Utilities.formatDate(new Date(row[timestampCol]), "Europe/London", "dd/MM/yyyy") : 'N/A',
        name: nameCol !== -1 ? (row[nameCol] || 'N/A') : 'N/A',
        approver: approverCol !== -1 ? (row[approverCol] || 'N/A') : 'N/A',
        fullDetails: fullDetails,
        timeline: formattedTimeline
      };
    });
    
    // Put the full, unfiltered data into the cache
    cache.put(cacheKey, JSON.stringify(processedData), 120); 

    // Return the full dataset. Filtering will now happen on the client-side.
    return {
      success: true,
      data: processedData
    };
  } catch (e) {
    logAction('SYSTEM_ERROR', 'GET_REQUESTS_FAIL', `Error in getRequestsAndTimeline: ${e.message} Stack: ${e.stack}`);
    return {
      success: false,
      error: e.message
    };
  }
}


function searchMaterials(token, searchTerm) {
  try {
    // 1. Authenticate the user using the current token system
    const user = getAuthenticatedUser(token);
    if (!user) {
      throw new Error("Access Denied. Your session may have expired.");
    }

    // 2. Access the 'Materials' sheet
    const materialsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Materials');
    if (materialsSheet.getLastRow() < 2) {
      return []; // Return empty if there's no data
    }

    // 3. Read all data and map it to objects
    const data = materialsSheet.getRange(2, 1, materialsSheet.getLastRow() - 1, 6).getValues();
    const allMaterials = data.map(row => ({
      materialId: row[0] || '',
      name: row[1] || '',
      description: row[2] || '',
      tags: row[3] || '',
      driveId: row[4] || '',
      fullDetails: row[5] || '{}' // Default to an empty JSON string to prevent errors
    }));

    // 4. Apply SITE-BASED filtering based on the user's role
    const siteFilteredMaterials = allMaterials.filter(material => {
      // Super Users and Regional Managers can see materials from all sites
      if (user.role === ROLES.SUPER_USER || user.role === ROLES.REGIONAL_MANAGER) {
        return true;
      }
      
      // Other users are filtered by their assigned siteId
      try {
        const details = JSON.parse(material.fullDetails);
        // Check if the 'Location' property in the archived details matches the user's site
        return details.Location === user.siteId; 
      } catch (e) {
        // If JSON parsing fails or the 'Location' key doesn't exist, hide the item
        return false;
      }
    });

    // 5. If there's no search term, return the site-filtered list
    if (!searchTerm || searchTerm.trim() === '') {
      return siteFilteredMaterials;
    }

    // 6. Otherwise, apply the SEARCH TERM filter to the already site-filtered list
    const lowerCaseSearchTerm = searchTerm.toLowerCase().trim();
    const finalResults = siteFilteredMaterials.filter(material => {
      return (material.name.toLowerCase().includes(lowerCaseSearchTerm)) ||
             (material.description.toLowerCase().includes(lowerCaseSearchTerm)) ||
             (material.tags.toLowerCase().includes(lowerCaseSearchTerm));
    });

    return finalResults;

  } catch (e) {
    // Log any errors and re-throw them so the client-side UI can display a message
    const actor = user ? user.username : 'UNKNOWN_USER';
    logAction(actor, 'SEARCH_MATERIALS_FAIL', `Error: ${e.message}`);
    throw new Error(`Failed to search materials: ${e.message}`);
  }
}

/**
 * UPDATED: Now gets APM emails and adds them to CC when sending emails.
 */
function saveResponse(formData) {
  const actor = 'PUBLIC_FORM_SUBMISSION';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const allQuestions = getQuestions('leaving').concat(getQuestions('entering'));
  const questionMap = allQuestions.reduce((map, q) => {
    map[q.id] = q;
    return map;
  }, {});
  const submissionId = generateNewSubmissionId(formData.location);
  const timestamp = new Date();
  const itemQuestionIds = getQuestions(formData.formType).filter(q => q.category === 'Item Information').map(q => q.id);
  const parentData = {};
  parentData['Timestamp'] = timestamp;
  parentData['Submission_ID'] = submissionId;
  parentData['Status'] = formData.formType === 'entering' ? 'Lodged' : 'Pending';
  parentData['Location'] = formData.location || 'N/A';
  parentData['Form Type'] = formData.formType || 'N/A';
  for (const questionId in formData) {
    if (questionId.startsWith('photos_') || questionId.endsWith('_other_text')) continue;
    if (!questionId.match(/_(\d+)$/) && !itemQuestionIds.includes(questionId) && questionId !== 'location' && questionId !== 'formType') {
      const qDetails = questionMap[questionId];
      if (qDetails) {
        let answer = formData[questionId];
        if (answer && typeof answer === 'string' && answer.toLowerCase().startsWith('other')) {
          const otherTextValue = formData[`${questionId}_other_text`];
          if (otherTextValue) {
            answer = `Other: ${otherTextValue}`;
          }
        }
        const headerName = `${qDetails.title.trim()}${DYNAMIC_COL_ANSWER_SUFFIX}`;
        parentData[headerName] = answer;
      }
    }
  }
  const items = {};
  for (const key in formData) {
    if (key.startsWith('photos_') || key.endsWith('_other_text')) continue;
    const match = key.match(/_(\d+)$/);
    if (match) {
      const index = match[1];
      const questionId = key.substring(0, key.length - match[0].length);
      if (!items[index]) items[index] = {};
      let itemAnswer = formData[key];
      if (itemAnswer && typeof itemAnswer === 'string' && itemAnswer.toLowerCase().startsWith('other')) {
        const otherTextValue = formData[`${key}_other_text`];
        if (otherTextValue) {
          itemAnswer = `Other: ${otherTextValue}`;
        }
      }
      items[index][questionId] = itemAnswer;
    }
  }
  const rowsToAppend = [];
  if (Object.keys(items).length > 0) {
    for (const index in items) {
      const singleItemData = { ...parentData
      };
      const currentItem = items[index];
      for (const questionId in currentItem) {
        const qDetails = questionMap[questionId];
        if (qDetails) {
          const headerName = `${qDetails.title.trim()}${DYNAMIC_COL_ANSWER_SUFFIX}`;
          singleItemData[headerName] = currentItem[questionId];
        }
      }
      const photoData = formData[`photos_${index}`];
      const photoFileIds = savePhotosToDrive(photoData, submissionId);
      singleItemData['Image IDs'] = photoFileIds;
      rowsToAppend.push(headers.map(header => singleItemData[header] || ''));
    }
  } else {
    rowsToAppend.push(headers.map(header => parentData[header] || ''));
  }
  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  }
  logAction(actor, 'SAVE_RESPONSE', `New pass ${submissionId} submitted for location: ${formData.location}`);
  
  const apmEmails = getApmEmailsForSite_(formData.location);

  if (formData.formType === 'leaving') {
    const approverName = formData['shell_approver'];
    if (approverName && !approverName.toLowerCase().startsWith('other')) {
      const approversSheet = ss.getSheetByName('Approvers');
      const approverData = approversSheet.getDataRange().getValues();
      const approver = approverData.find(row => row[1] === approverName && row[3] === formData.location);
      if (approver) {
        const approverEmail = approver[2];
        const token = Utilities.getUuid();
        const tasksSheet = ss.getSheetByName('ApprovalTasks');
        tasksSheet.appendRow([token, submissionId, approverEmail, new Date(), 'Active', '']);
        const approvalUrl = `${ScriptApp.getService().getUrl()}?page=approve&token=${token}`;
        const subject = `ACTION REQUIRED: Gate Pass Request (${submissionId})`;
        const body = `<p>A new gate pass request requires your approval.</p><p><strong>Request ID:</strong> ${submissionId}</p><p>Please review the details and take action by clicking the link below:</p><p><a href="${approvalUrl}" style="font-size:16px; font-weight:bold; padding:10px 15px; background-color:#005f90; color:white; text-decoration:none; border-radius:5px;">Review Request</a></p>`;
        sendAppEmail({
          to: approverEmail,
          cc: apmEmails,
          subject: subject,
          htmlBody: body
        });
        logAction('SYSTEM', 'SEND_APPROVAL_EMAIL', `Approval email for ${submissionId} sent to ${approverEmail}. CC: ${apmEmails || 'None'}`);
      }
    }
  } else if (formData.formType === 'entering') {
    const submitterEmail = formData['submitter_email'];
    if (submitterEmail) {
      const subject = `Confirmation: Your Gate Pass has been lodged (${submissionId})`;
      const body = `<p>Thank you for your submission.</p><p>Your Gate Pass ID is <strong>${submissionId}</strong>. This has been lodged in our system and will be awaiting arrival.</p>`;
      sendAppEmail({
        to: submitterEmail,
        cc: apmEmails,
        subject: subject,
        htmlBody: body
      });
    }
  }
  const cache = CacheService.getScriptCache();
  const allCacheKeys = cache.getAllKeys();
  const keysToRemove = allCacheKeys.filter(k => k.startsWith('requests_timeline_'));
  if (keysToRemove.length > 0) {
    cache.removeAll(keysToRemove);
  }
  return 'Form submitted successfully!';
}

function savePhotosToDrive(photos, submissionId) {
  if (!photos || !Array.isArray(photos) || photos.length === 0) {
    return '';
  }
  const parentFolderId = '1UomVP4aPrlQTWNHdl-fB3pn4t4K-dKz-';
  let parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(parentFolderId);
  } catch (e) {
    logAction('SYSTEM_ERROR', 'DRIVE_ACCESS_FAILURE', `Could not access parent folder ID ${parentFolderId}. Error: ${e.message}`);
    return '';
  }
  const subFolderName = `GatePass_${submissionId}`;
  let subFolder;
  const folders = parentFolder.getFoldersByName(subFolderName);
  if (folders.hasNext()) {
    subFolder = folders.next();
  } else {
    subFolder = parentFolder.createFolder(subFolderName);
  }
  const fileIds = [];
  photos.forEach((photo, index) => {
    try {
      const base64Data = photo.data;
      const contentType = photo.type || 'image/jpeg';
      if (!base64Data || typeof base64Data !== 'string') {
        throw new Error(`Photo data for index ${index} is missing or has the wrong format.`);
      }
      const decodedData = Utilities.base64Decode(base64Data, Utilities.Charset.UTF_8);
      const blob = Utilities.newBlob(decodedData, contentType, `${submissionId}_photo_${index + 1}.jpg`);
      const file = subFolder.createFile(blob);
      fileIds.push(file.getId());
    } catch (e) {
      logAction('SYSTEM_ERROR', 'PHOTO_SAVE_FAILURE', `Failed to save photo for ${submissionId}. Error: ${e.message}`);
    }
  });
  return fileIds.join(',');
}

function setupResponseSheetHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const responsesSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
  if (!responsesSheet) throw new Error(`Sheet '${RESPONSES_SHEET_NAME}' not found.`);
  const staticHeaders = ['Timestamp', 'Submission_ID', 'Status', 'Location', 'Form Type', 'Image IDs', 'Approval Comments'];
  const allQuestions = getQuestions('leaving').concat(getQuestions('entering'));
  const dynamicHeaders = [...new Set(allQuestions.map(q => q.title.trim() + DYNAMIC_COL_ANSWER_SUFFIX))];
  const allHeaders = [...new Set(staticHeaders.concat(dynamicHeaders))];
  responsesSheet.clear();
  responsesSheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]).setFontWeight("bold");
  responsesSheet.setFrozenRows(1);
  const timestampCol = allHeaders.indexOf('Timestamp') + 1;
  if (timestampCol > 0) {
    responsesSheet.getRange(2, timestampCol, responsesSheet.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  }
}

function getAndValidateLocation(identifier) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOCATIONS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet with name "${LOCATIONS_SHEET_NAME}" was not found.`);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().toLowerCase() === identifier.toLowerCase()) {
        return {
          name: data[i][0],
          identifier: data[i][1],
          cluster: data[i][2]
        };
      }
    }
    return null;
  } catch (err) {
    throw new Error('Failed to get location data. ' + err.message);
  }
}

function getDataForApp(formType, siteId) {
  try {
    const allQuestions = getQuestions(formType);
    const approversSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Approvers');
    let approverNames = [];
    if (approversSheet.getLastRow() > 1) {
      const approverData = approversSheet.getRange(2, 1, approversSheet.getLastRow() - 1, 4).getValues();
      approverNames = approverData
        .filter(row => row[3] === siteId)
        .map(row => row[1]);
    }
    const shellApproverQuestion = allQuestions.find(q => q.title === 'Shell Approver');
    if (shellApproverQuestion && approverNames.length > 0) {
      shellApproverQuestion.options = approverNames;
    }
    const categoryOrder = [...new Set(allQuestions.map(q => q.category))];
    const itemQuestions = allQuestions.filter(q => q.category === 'Item Information');
    const mainQuestions = allQuestions.filter(q => q.category !== 'Item Information');
    return {
      mainQuestions,
      itemQuestions,
      categoryOrder
    };
  } catch (error) {
    throw new Error("Could not retrieve initial data from the spreadsheet. " + error.message);
  }
}

function getQuestions(formType) {
  try {
    const sheetName = formType === 'entering' ? QUESTIONS_ENTERING_SHEET_NAME : QUESTIONS_SHEET_NAME;
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    return data.map(row => (row[5] ? {
      category: row[0],
      title: row[1],
      type: row[2],
      options: row[3] ? row[3].toString().split(',').map(item => item.trim()) : [],
      photos: row[4],
      id: row[5].toString()
    } : null)).filter(Boolean);
  } catch (err) {
    throw new Error(`Failed to get question data. ${err.message}`);
  }
}

function getApprovalTaskDetails(token) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasksSheet = ss.getSheetByName('ApprovalTasks');
  if (tasksSheet.getLastRow() < 2) throw new Error("Approval task not found.");
  const taskData = tasksSheet.getDataRange().getValues();
  const taskHeaders = taskData.shift();
  const taskRow = taskData.find(row => row[taskHeaders.indexOf('ApprovalToken')] === token);
  if (!taskRow) throw new Error("This approval link is invalid.");
  if (taskRow[taskHeaders.indexOf('Status')] !== 'Active') throw new Error("This approval link has already been used or expired.");
  const submissionId = taskRow[taskHeaders.indexOf('Submission_ID')];
  const responsesSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
  const responseDataWithHeaders = responsesSheet.getDataRange().getValues();
  const responseHeaders = responseDataWithHeaders.shift();
  const requestRowIndexes = [];
  responseDataWithHeaders.forEach((row, index) => {
    if (row[responseHeaders.indexOf('Submission_ID')] === submissionId) {
      requestRowIndexes.push({
        rowIndex: index + 2,
        data: row
      });
    }
  });
  if (requestRowIndexes.length === 0) throw new Error("The original request could not be found.");
  const requestDetails = {};
  responseHeaders.forEach((header, i) => {
    const value = requestRowIndexes[0].data[i];
    if (value && header) {
      requestDetails[header] = (value instanceof Date) ? Utilities.formatDate(value, "Europe/London", "dd/MM/yyyy HH:mm:ss") : value;
    }
  });
  const itemQuestionTitles = getQuestions('leaving').filter(q => q.category === 'Item Information').map(q => q.title.trim());
  requestDetails.Items = requestRowIndexes.map(indexedRow => {
    const itemDetails = {
      spreadsheetRow: indexedRow.rowIndex
    };
    const rowData = indexedRow.data;
    itemQuestionTitles.forEach(title => {
      const headerName = `${title}${DYNAMIC_COL_ANSWER_SUFFIX}`;
      const colIndex = responseHeaders.indexOf(headerName);
      if (colIndex !== -1 && rowData[colIndex]) {
        itemDetails[title] = rowData[colIndex];
      }
    });
    const imageIdIndex = responseHeaders.indexOf("Image IDs");
    const imageIds = rowData[imageIdIndex] ? rowData[imageIdIndex].toString().split(',').filter(Boolean) : [];
    itemDetails.images = imageIds;
    return itemDetails;
  }).filter(item => Object.keys(item).length > 1);
  const siteId = requestDetails['Location'];
  const approversSheet = ss.getSheetByName('Approvers');
  let otherApprovers = [];
  if (approversSheet.getLastRow() > 1) {
    const allApprovers = approversSheet.getRange(2, 1, approversSheet.getLastRow() - 1, 4).getValues();
    otherApprovers = allApprovers
      .filter(row => row[3] === siteId)
      .map(row => ({
        fullName: row[1],
        email: row[2]
      }));
  }
  delete requestDetails.images;
  return {
    success: true,
    requestDetails,
    otherApprovers,
    token
  };
}

/**
 * UPDATED: Sets up the Sites sheet with the new columns.
 */
function setupProject() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(USERS_SHEET) || ss.insertSheet(USERS_SHEET);
  if (usersSheet.getLastRow() === 0) {
    usersSheet.getRange(1, 1, 1, 9).setValues([
      ['Username', 'PasswordHash', 'Role', 'SiteID', 'FullName', 'Position', 'IsActive', 'LastLogin', 'IsTempPassword']
    ]).setFontWeight('bold');
    usersSheet.setFrozenRows(1);
  }
  const sitesSheet = ss.getSheetByName(SITES_SHEET) || ss.insertSheet(SITES_SHEET);
  if (sitesSheet.getLastRow() === 0) {
    sitesSheet.getRange(1, 1, 1, 5).setValues([
      ['SiteID', 'SiteName', 'APM_Email_1', 'APM_Email_2', 'ApproverListLastUpdated']
    ]).setFontWeight('bold');
    sitesSheet.setFrozenRows(1);
    sitesSheet.getRange(2, 1, 3, 5).setValues([
      ['Sarnia', 'Sarnia Site', '', '', ''],
      ['Geismar', 'Geismar Site', '', '', ''],
      ['Scotford', 'Scotford Site', '', '', '']
    ]);
  }
  const logsSheet = ss.getSheetByName(LOGS_SHEET) || ss.insertSheet(LOGS_SHEET);
  if (logsSheet.getLastRow() === 0) {
    logsSheet.getRange(1, 1, 1, 4).setValues([
      ['Timestamp', 'Username', 'Action', 'Details']
    ]).setFontWeight('bold');
    logsSheet.setFrozenRows(1);
  }
  const resetsSheet = ss.getSheetByName('PasswordResets') || ss.insertSheet('PasswordResets');
  if (resetsSheet.getLastRow() === 0) {
    resetsSheet.getRange(1, 1, 1, 4).setValues([
      ['Token', 'UsernameToReset', 'RequestTimestamp', 'IsUsed']
    ]).setFontWeight('bold');
    resetsSheet.setFrozenRows(1);
  }
  const materialsSheet = ss.getSheetByName('Materials') || ss.insertSheet('Materials');
  if (materialsSheet.getLastRow() === 0) {
    materialsSheet.getRange(1, 1, 1, 6).setValues([
      ['MaterialID', 'Name', 'Description', 'Tags', 'DriveID', 'Full Details']
    ]).setFontWeight('bold');
    materialsSheet.setFrozenRows(1);
  }
  const approversSheet = ss.getSheetByName('Approvers') || ss.insertSheet('Approvers');
  if (approversSheet.getLastRow() === 0) {
    approversSheet.getRange(1, 1, 1, 4).setValues([
      ['ApproverID', 'FullName', 'Email', 'SiteID']
    ]).setFontWeight('bold');
    approversSheet.setFrozenRows(1);
  }
  const tasksSheet = ss.getSheetByName('ApprovalTasks') || ss.insertSheet('ApprovalTasks');
  if (tasksSheet.getLastRow() === 0) {
    tasksSheet.getRange(1, 1, 1, 6).setValues([
      ['ApprovalToken', 'Submission_ID', 'AssignedToEmail', 'RequestTimestamp', 'Status', 'CompletedTimestamp']
    ]).setFontWeight('bold');
    tasksSheet.setFrozenRows(1);
  }
  syncUsersToPropertiesService();
}

function logAction(username, action, details) {
  try {
    const logsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOGS_SHEET);
    logsSheet.appendRow([new Date(), username, action, details]);
  } catch (e) {
    console.error("Failed to log action: " + e.toString());
  }
}

function syncUsersToPropertiesService() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const scriptProperties = PropertiesService.getScriptProperties();
  const allKeys = scriptProperties.getKeys();
  const userKeys = allKeys.filter(key => key.startsWith('*user_'));
  if (userKeys.length > 0) {
    userKeys.forEach(key => {
      scriptProperties.deleteProperty(key);
    });
  }
  const userProperties = {};
  data.forEach(row => {
    const user = {};
    headers.forEach((header, i) => user[header] = row[i]);
    if (user.IsActive) {
      const key = `*user_${user.Username.toLowerCase()}`;
      const value = [user.PasswordHash, user.Role, user.SiteID, user.FullName, user.Position, user.IsTempPassword || false].join('||');
      userProperties[key] = value;
    }
  });
  if (Object.keys(userProperties).length > 0) {
    scriptProperties.setProperties(userProperties, false);
  }
}

function hashPassword_(password) {
  const salt = Utilities.getUuid();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  return `${salt}:${Utilities.base64Encode(digest)}`;
}

function verifyPassword_(password, storedHash) {
  if (!password || !storedHash) return false;
  const [salt, originalHash] = storedHash.split(':');
  if (!salt || !originalHash) return false;
  const comparisonHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt));
  return originalHash === comparisonHash;
}

function findUserRow_(username) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const usernameIndex = headers.indexOf('Username');
  if (usernameIndex === -1) return null;
  const userRowIndex = data.findIndex(row => row[usernameIndex] && row[usernameIndex].toLowerCase() === username.toLowerCase());
  if (userRowIndex === -1) return null;
  const user = {
    rowIndex: userRowIndex + 2
  };
  headers.forEach((header, index) => user[header] = data[userRowIndex][index]);
  return user;
}

function getUserFromProperties_(username) {
  const userDataString = PropertiesService.getScriptProperties().getProperty(`*user_${username.toLowerCase()}`);
  if (!userDataString) return null;
  const parts = userDataString.split('||');
  return {
    Username: username,
    PasswordHash: parts[0],
    Role: parts[1],
    SiteID: parts[2],
    FullName: parts[3],
    Position: parts[4],
    IsTempPassword: parts[5] === 'true'
  };
}

function createSubmissionFolder(submissionFolderName) {
  const parentFolderId = '1UomVP4aPrlQTWNHdl-fB3pn4t4K-dKz-';
  let parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(parentFolderId);
  } catch (e) {
    throw new Error(`Could not access the specified Drive folder. Please check the folder ID and make sure the script has permission to access it. Folder ID: ${parentFolderId}`);
  }
  const newFolder = parentFolder.createFolder(submissionFolderName);
  return {
    url: newFolder.getUrl(),
    id: newFolder.getId()
  };
}

function findHeaderIndex_(headersArray, searchText) {
  const normalizedSearchText = searchText.toLowerCase().trim();
  for (let i = 0; i < headersArray.length; i++) {
    if (headersArray[i].toLowerCase().trim() === normalizedSearchText) return i;
  }
  return -1;
}

function updateTicketStatus(token, uniqueId, newStatus) {
  const user = getAuthenticatedUser(token);
  if (!user) throw new Error('Access Denied.');
  if (!STATUS_OPTIONS.includes(newStatus)) throw new Error('Invalid status.');
  if (user.role === ROLES.REGIONAL_MANAGER && newStatus !== 'Closed') throw new Error('Regional Managers can only Close tickets.');
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSES_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const idCol = findHeaderIndex_(data[0], 'Submission_ID');
  const statusCol = findHeaderIndex_(data[0], 'Status');
  const rowIdx = data.findIndex(row => row[idCol] == uniqueId);
  if (rowIdx === -1) throw new Error('Ticket not found.');
  sheet.getRange(rowIdx + 1, statusCol + 1).setValue(newStatus);
  logAction(user.username, 'UPDATE_TICKET_STATUS', `Ticket ${uniqueId} -> ${newStatus}`);
  const cache = CacheService.getScriptCache();
  const allCacheKeys = cache.getAllKeys();
  const keysToRemove = allCacheKeys.filter(k => k.startsWith('requests_timeline_'));
  if (keysToRemove.length > 0) {
    cache.removeAll(keysToRemove);
  }
  return {
    success: true,
    message: 'Status updated.'
  };
}

function checkAdminPermissions_(user, requireSuperUser = false) {
  if (!user) throw new Error('Access Denied. Not logged in.');
  const allowed = requireSuperUser ? [ROLES.SUPER_USER] : [ROLES.SUPER_USER, ROLES.REGIONAL_MANAGER, ROLES.SITE_SUPERVISOR];
  if (!allowed.includes(user.role)) throw new Error('Access Denied. Insufficient permissions.');
}

function getActionLogs(token, page = 1, pageSize = 50, actionFilter = 'ALL') {
  try {
    const currentUser = getAuthenticatedUser(token);
    if (!currentUser || ![ROLES.SUPER_USER, ROLES.REGIONAL_MANAGER].includes(currentUser.role)) {
      throw new Error('Access Denied. You do not have permission to view system logs.');
    }
    const logsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOGS_SHEET);
    let logData = logsSheet.getLastRow() < 2 ? [] : logsSheet.getRange(2, 1, logsSheet.getLastRow() - 1, 4).getValues();
    if (actionFilter && actionFilter !== 'ALL') {
      logData = logData.filter(row => row[2] === actionFilter);
    }
    const totalLogs = logData.length;
    if (totalLogs === 0) {
      return {
        logs: [],
        totalPages: 0,
        currentPage: 1
      };
    }
    logData.sort((a, b) => new Date(b[0]) - new Date(a[0]));
    const totalPages = Math.ceil(totalLogs / pageSize);
    const pageNum = parseInt(page, 10);
    const startIndex = (pageNum - 1) * pageSize;
    const pagedData = logData.slice(startIndex, startIndex + pageSize);
    const logs = pagedData.map(row => [
      row[0] ? Utilities.formatDate(new Date(row[0]), "Europe/London", "dd/MM/yyyy HH:mm:ss") : null,
      row[1],
      row[2],
      row[3]
    ]);
    return {
      logs,
      totalPages,
      currentPage: pageNum
    };
  } catch (e) {
    console.error("getActionLogs Error: " + e.toString());
    throw e;
  }
}

function getUniqueLogActions(token) {
  try {
    const currentUser = getAuthenticatedUser(token);
    if (!currentUser || ![ROLES.SUPER_USER, ROLES.REGIONAL_MANAGER].includes(currentUser.role)) {
      throw new Error('Access Denied.');
    }
    const logsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOGS_SHEET);
    if (logsSheet.getLastRow() < 2) {
      return [];
    }
    const actionData = logsSheet.getRange(2, 3, logsSheet.getLastRow() - 1, 1).getValues();
    const uniqueActions = [...new Set(actionData.flat())];
    return uniqueActions.filter(Boolean).sort();
  } catch (e) {
    throw e;
  }
}

function addUser(token, userData) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  if (currentUser.role === ROLES.REGIONAL_MANAGER && userData.role !== ROLES.SITE_SUPERVISOR) throw new Error('Regional Managers can only add Site Supervisors.');
  if (currentUser.role === ROLES.SITE_SUPERVISOR) {
    if (userData.role !== ROLES.SECURITY_OFFICER) throw new Error('Site Supervisors can only add Security Officers.');
    if (userData.siteId !== currentUser.siteId) throw new Error('You can only add users to your own site.');
  }
  if (findUserRow_(userData.username)) throw new Error('Username already exists.');
  const usersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
  usersSheet.appendRow([userData.username, hashPassword_(userData.username), userData.role, userData.siteId, userData.fullName, userData.position, true, null, true]);
  logAction(currentUser.username, 'ADD_USER', `Added: ${userData.username}`);
  syncUsersToPropertiesService();
  CacheService.getScriptCache().remove('ALL_USERS_DATA');
  return {
    success: true,
    message: 'User added.'
  };
}

function updateUser(token, userData) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  const userToUpdate = findUserRow_(userData.username);
  if (!userToUpdate) throw new Error('User not found.');
  if (currentUser.role === ROLES.SITE_SUPERVISOR && userToUpdate.SiteID !== currentUser.siteId) throw new Error('You cannot edit users outside your site.');
  const usersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
  usersSheet.getRange(userToUpdate.rowIndex, 3, 1, 5).setValues([
    [userData.role, userData.siteId, userData.fullName, userData.position, userData.isActive]
  ]);
  logAction(currentUser.username, 'UPDATE_USER', `Updated: ${userData.username}`);
  syncUsersToPropertiesService();
  CacheService.getScriptCache().remove('ALL_USERS_DATA');
  return {
    success: true,
    message: 'User updated.'
  };
}

function deleteUser(token, username) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser, true);
  const userToDelete = findUserRow_(username);
  if (!userToDelete) throw new Error('User not found.');
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET).getRange(userToDelete.rowIndex, 7).setValue(false);
  logAction(currentUser.username, 'DEACTIVATE_USER', `Deactivated: ${username}`);
  syncUsersToPropertiesService();
  CacheService.getScriptCache().remove('ALL_USERS_DATA');
  return {
    success: true,
    message: 'User deactivated.'
  };
}

function resetPassword(token, username) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  const userToReset = findUserRow_(username);
  if (!userToReset) throw new Error('User not found');
  const newPasswordHash = hashPassword_(username);
  const usersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
  usersSheet.getRange(userToReset.rowIndex, 2).setValue(newPasswordHash);
  usersSheet.getRange(userToReset.rowIndex, 9).setValue(true);
  const performedBy = currentUser ? currentUser.username : 'SCRIPT_EDITOR';
  logAction(performedBy, 'RESET_PASSWORD', `Reset password for: ${username}`);
  syncUsersToPropertiesService();
  return {
    success: true,
    message: `Password for ${username} has been successfully reset.`
  };
}

function resetPasswordFromEditor() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Reset User Password', 'Please enter the full username (email) of the user you wish to reset:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK && response.getResponseText()) {
    const username = response.getResponseText();
    const result = resetPassword('MANUAL_FROM_EDITOR', username);
    ui.alert(result.message);
  } else {
    ui.alert('Password reset cancelled.');
  }
}

function MANUAL_PASSWORD_RESET() {
  const usernameToReset = "martin.bancroft@uk.g4s.com";
  if (usernameToReset === "user.email@example.com" || !usernameToReset) {
    return;
  }
  const result = resetPassword('MANUAL_SCRIPT', usernameToReset);
}

function requestPasswordReset(username) {
  const userToReset = findUserRow_(username);
  if (!userToReset) throw new Error("Username not found.");
  const usersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
  const allUsers = usersSheet.getDataRange().getValues();
  const headers = allUsers.shift();
  const roleIndex = headers.indexOf('Role');
  const siteIdIndex = headers.indexOf('SiteID');
  const emailIndex = headers.indexOf('Username');
  const supervisor = allUsers.find(userRow => userRow[siteIdIndex] === userToReset.SiteID && userRow[roleIndex] === ROLES.SITE_SUPERVISOR);
  if (!supervisor) throw new Error("Could not find a Site Supervisor assigned to this user's location.");
  const supervisorEmail = supervisor[emailIndex];
  const token = Utilities.getUuid();
  const resetsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PasswordResets');
  resetsSheet.appendRow([token, username, new Date(), false]);
  const approvalUrl = `${ScriptApp.getService().getUrl()}?page=approveReset&token=${token}`;
  const subject = "Password Reset Request for Gatehouse App";
  const body = `<p>Hello,</p><p>A password reset has been requested for the user: <strong>${username}</strong>.</p><p>To approve this reset, please click the link below. The user's password will be reset to their username.</p><p><a href="${approvalUrl}" style="font-size:16px; font-weight:bold; padding:10px 15px; background-color:#005f90; color:white; text-decoration:none; border-radius:5px;">Approve Password Reset</a></p><p>If you did not expect this request, you can safely ignore this email.</p>`;
  sendAppEmail({ to: supervisorEmail, subject: subject, htmlBody: body });
  logAction(username, 'REQUEST_PASSWORD_RESET', `Reset request sent to supervisor: ${supervisorEmail}`);
  return {
    success: true,
    message: 'A reset request has been sent to the Site Supervisor for approval.'
  };
}

function processResetApproval(token) {
  if (!token) return "<h1>Error</h1><p>No approval token provided.</p>";
  const resetsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PasswordResets');
  const data = resetsSheet.getDataRange().getValues();
  const headers = data.shift();
  const tokenIndex = headers.indexOf('Token');
  const userIndex = headers.indexOf('UsernameToReset');
  const usedIndex = headers.indexOf('IsUsed');
  const requestRowIndex = data.findIndex(row => row[tokenIndex] === token);
  if (requestRowIndex === -1) return "<h1>Error</h1><p>This approval link is invalid.</p>";
  const requestRow = data[requestRowIndex];
  if (requestRow[usedIndex] === true) return "<h1>Error</h1><p>This approval link has already been used.</p>";
  const usernameToReset = requestRow[userIndex];
  try {
    resetPassword('SUPERVISOR_APPROVAL', usernameToReset);
    resetsSheet.getRange(requestRowIndex + 2, usedIndex + 1).setValue(true);
    const supervisorEmail = Session.getActiveUser().getEmail();
    logAction(supervisorEmail, 'APPROVE_PASSWORD_RESET', `Approved reset for ${usernameToReset}`);
    return `<h1>Success</h1><p>The password for user <strong>${usernameToReset}</strong> has been successfully reset.</p>`;
  } catch (e) {
    return `<h1>Error</h1><p>An error occurred: ${e.message}</p>`;
  }
}

/**
 * UPDATED: Handles new APM email fields.
 */
function addSite(token, siteData) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  const sitesSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SITES_SHEET);
  const siteIds = sitesSheet.getRange("A2:A").getValues().flat();
  if (siteIds.includes(siteData.siteId)) throw new Error('A site with this Site ID already exists.');
  sitesSheet.appendRow([siteData.siteId, siteData.siteName, siteData.apmEmail1 || '', siteData.apmEmail2 || '', '']);
  logAction(currentUser.username, 'ADD_SITE', `Added new site: ${siteData.siteId} - ${siteData.siteName}`);
  CacheService.getScriptCache().remove('SITES_DATA');
  return {
    success: true,
    message: 'Site added successfully.'
  };
}

/**
 * UPDATED: More robust update logic for new APM email fields.
 */
function updateSite(token, siteData) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sitesSheet = ss.getSheetByName(SITES_SHEET);
  const data = sitesSheet.getDataRange().getValues();
  const headers = data.shift();
  const idColIdx = headers.indexOf('SiteID');
  const nameColIdx = headers.indexOf('SiteName');
  const apm1ColIdx = headers.indexOf('APM_Email_1');
  const apm2ColIdx = headers.indexOf('APM_Email_2');

  const rowIndex = data.findIndex(row => row[idColIdx] === siteData.siteId);
  if (rowIndex === -1) throw new Error('Site not found. Cannot update.');

  const sheetRow = rowIndex + 2;
  sitesSheet.getRange(sheetRow, nameColIdx + 1).setValue(siteData.siteName);
  sitesSheet.getRange(sheetRow, apm1ColIdx + 1).setValue(siteData.apmEmail1 || '');
  sitesSheet.getRange(sheetRow, apm2ColIdx + 1).setValue(siteData.apmEmail2 || '');

  logAction(currentUser.username, 'UPDATE_SITE', `Updated site: ${siteData.siteId} - New Name: ${siteData.siteName}`);
  CacheService.getScriptCache().remove('SITES_DATA');
  return {
    success: true,
    message: 'Site updated successfully.'
  };
}

function deleteSite(token, siteId) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser, true);
  const sitesSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SITES_SHEET);
  const siteIds = sitesSheet.getRange("A2:A").getValues().flat();
  const rowIndex = siteIds.indexOf(siteId);
  if (rowIndex === -1) throw new Error('Site not found. Cannot delete.');
  sitesSheet.deleteRow(rowIndex + 2);
  logAction(currentUser.username, 'DELETE_SITE', `Deleted site: ${siteId}`);
  CacheService.getScriptCache().remove('SITES_DATA');
  return {
    success: true,
    message: 'Site deleted successfully.'
  };
}

function initialAdminSetup() {
  const usersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
  const firstUserUsername = usersSheet.getRange("A2").getValue();
  if (!firstUserUsername) {
    return;
  }
  const newPasswordHash = hashPassword_(firstUserUsername);
  usersSheet.getRange("B2").setValue(newPasswordHash);
  usersSheet.getRange("I2").setValue(true);
  logAction('SCRIPT_EDITOR', 'INITIAL_ADMIN_SETUP', `Set initial password for: ${firstUserUsername}`);
  syncUsersToPropertiesService();
}

function getImageUrls(token, imageIdsString) {
  const user = getAuthenticatedUser(token);
  if (!user) {
    throw new Error("Access denied. You must be logged in to view images.");
  }
  if (!imageIdsString || typeof imageIdsString !== 'string') {
    return [];
  }
  const ids = imageIdsString.split(',').filter(id => id.trim() !== '');
  if (ids.length === 0) {
    return [];
  }
  const urls = ids.map(id => {
    try {
      const file = DriveApp.getFileById(id.trim());
      const blob = file.getBlob();
      const contentType = blob.getContentType();
      const base64Data = Utilities.base64Encode(blob.getBytes());
      return `data:${contentType};base64,${base64Data}`;
    } catch (e) {
      logAction('SYSTEM_ERROR', 'GET_IMAGE_URL_FAIL', `Could not retrieve file ID: ${id}. Error: ${e.message}`);
      return null;
    }
  });
  return urls.filter(Boolean);
}

function getTimelineForRequest(submissionId) {
  const logsSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOGS_SHEET);
  const logData = logsSheet.getDataRange().getValues();
  logData.shift();
  return logData.filter(logRow => logRow[3] && logRow[3].includes(submissionId)).map(logRow => {
    let event = logRow[2];
    if (event === 'UPDATE_TICKET_STATUS') event = `Status changed to '${logRow[3].split('->')[1]?.trim()}'`;
    else if (event === 'SAVE_RESPONSE') event = 'Request Submitted';
    return {
      timestamp: Utilities.formatDate(new Date(logRow[0]), "Europe/London", "dd/MM/yyyy HH:mm:ss"),
      user: logRow[1],
      event
    };
  }).sort((a, b) => {
    const dateA = new Date(a.timestamp.replace(/(\d{2})\/(\d{2})\/(\d{4})/, '$2/$1/$3')).getTime();
    const dateB = new Date(b.timestamp.replace(/(\d{2})\/(\d{2})\/(\d{4})/, '$2/$1/$3')).getTime();
    return dateA - dateB;
  });
}

function generateNewSubmissionId(locationName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let newId;
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RESPONSES_SHEET_NAME);
    const idColumn = 2;
    const prefix = locationName.substring(0, 3).toUpperCase();
    let maxNum = 0;
    if (sheet.getLastRow() > 1) {
      const ids = sheet.getRange(2, idColumn, sheet.getLastRow() - 1, 1).getValues().flat();
      ids.forEach(id => {
        if (id && id.startsWith(prefix)) {
          const numPart = parseInt(id.split('-')[1], 10);
          if (numPart > maxNum) {
            maxNum = numPart;
          }
        }
      });
    }
    const nextNum = maxNum + 1;
    newId = `${prefix}-${String(nextNum).padStart(4, '0')}`;
  } finally {
    lock.releaseLock();
  }
  return newId;
}

/**
 * UPDATED: Now includes APM emails in CC for submitter notifications.
 */
function processApproval(action, token, comments, tagsByRow) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tasksSheet = ss.getSheetByName('ApprovalTasks');
    const taskData = tasksSheet.getDataRange().getValues();
    const taskHeaders = taskData.shift();
    const taskRowIndex = taskData.findIndex(row => row[taskHeaders.indexOf('ApprovalToken')] === token);
    if (taskRowIndex === -1) throw new Error("Invalid token.");
    const taskRow = taskData[taskRowIndex];
    if (taskRow[taskHeaders.indexOf('Status')] !== 'Active') throw new Error("This task has already been completed.");
    const submissionId = taskRow[taskHeaders.indexOf('Submission_ID')];
    const responsesSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
    const responseDataWithHeaders = responsesSheet.getDataRange().getValues();
    const responseHeaders = responseDataWithHeaders.shift();
    const statusColIndex = responseHeaders.indexOf('Status');
    const idColIndex = responseHeaders.indexOf('Submission_ID');
    const commentsColIndex = responseHeaders.indexOf('Approval Comments');
    for (let i = 0; i < responseDataWithHeaders.length; i++) {
      if (responseDataWithHeaders[i][idColIndex] === submissionId) {
        responsesSheet.getRange(i + 2, statusColIndex + 1).setValue(action);
        if (comments) {
          responsesSheet.getRange(i + 2, commentsColIndex + 1).setValue(comments);
        }
      }
    }
    tasksSheet.getRange(taskRowIndex + 2, taskHeaders.indexOf('Status') + 1).setValue('Completed');
    tasksSheet.getRange(taskRowIndex + 2, taskHeaders.indexOf('CompletedTimestamp') + 1).setValue(new Date());
    const approverEmail = taskRow[taskHeaders.indexOf('AssignedToEmail')];
    logAction(approverEmail, 'PROCESS_APPROVAL', `Request ${submissionId} was ${action}. Comments: ${comments || 'N/A'}`);
    if (action === 'Approved' && tagsByRow && Object.keys(tagsByRow).length > 0) {
      for (const rowNum in tagsByRow) {
        const tags = tagsByRow[rowNum];
        if (tags) {
          archiveMaterial(parseInt(rowNum, 10), tags);
          logAction(approverEmail, 'ARCHIVE_ITEM', `Item from pass ${submissionId} (Row ${rowNum}) tagged and archived with tags: ${tags}`);
        }
      }
    }
    try {
      const emailHeader = 'Your Email Address (for notifications)' + DYNAMIC_COL_ANSWER_SUFFIX;
      const emailColIndex = responseHeaders.indexOf(emailHeader);
      if (emailColIndex !== -1) {
        const firstResponseRow = responseDataWithHeaders.find(r => r[idColIndex] === submissionId);
        const submitterEmail = firstResponseRow ? firstResponseRow[emailColIndex] : null;
        const siteId = firstResponseRow ? firstResponseRow[responseHeaders.indexOf('Location')] : null;
        const apmEmails = getApmEmailsForSite_(siteId);
        if (submitterEmail && Utilities.newBlob(submitterEmail).getContentType() === 'text/plain') {
          let subject = '';
          let body = '';
          if (action === 'Approved') {
            subject = `UPDATE: Your Gate Pass Request has been Approved (${submissionId})`;
            body = `<p>Hello,</p><p>This is a notification to confirm that your Gate Pass request (ID: <strong>${submissionId}</strong>) has been <strong>approved</strong>.</p><p>You may now proceed with moving the items off-site.</p><p><b>Approver's Comments:</b> ${comments || 'None'}</p><p>Thank you.</p>`;
          } else if (action === 'Rejected') {
            subject = `UPDATE: Your Gate Pass Request has been Rejected (${submissionId})`;
            body = `<p>Hello,</p><p>This is a notification that your Gate Pass request (ID: <strong>${submissionId}</strong>) has been <strong>rejected</strong>.</p><p><b>Rejection Reason:</b> ${comments || 'No reason provided. Please contact the site for details.'}</p><p>Please do not attempt to move the items off-site. You may need to submit a new request with corrected information.</p>`;
          }
          if (subject) {
            sendAppEmail({
              to: submitterEmail,
              cc: apmEmails,
              subject: subject,
              htmlBody: body
            });
            logAction('SYSTEM', 'SUBMITTER_NOTIFICATION_SENT', `Notification sent to ${submitterEmail} for pass ${submissionId}. CC: ${apmEmails || 'None'}`);
          }
        }
      }
    } catch (e) {
      logAction('SYSTEM_ERROR', 'SUBMITTER_NOTIFICATION_FAIL', `Failed to send notification for pass ${submissionId}. Error: ${e.message}`);
    }
    const cache = CacheService.getScriptCache();
    const allCacheKeys = cache.getAllKeys();
    const keysToRemove = allCacheKeys.filter(k => k.startsWith('requests_timeline_'));
    if (keysToRemove.length > 0) {
      cache.removeAll(keysToRemove);
    }
    return {
      success: true,
      message: `Request has been successfully ${action.toLowerCase()}.`
    };
  } finally {
    lock.releaseLock();
  }
}

function archiveMaterial(rowNumber, tags) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responsesSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
    const responseHeaders = responsesSheet.getRange(1, 1, 1, responsesSheet.getLastColumn()).getValues()[0];
    const requestRow = responsesSheet.getRange(rowNumber, 1, 1, responseHeaders.length).getValues()[0];
    const submissionId = requestRow[responseHeaders.indexOf('Submission_ID')];
    const imageIds = requestRow[responseHeaders.indexOf("Image IDs")] || '';
    if (!imageIds) {
      logAction('SYSTEM', 'ARCHIVE_SKIPPED', `Archiving skipped for item in row ${rowNumber} because no images were attached.`);
      return;
    }
    const fullDetailsObject = {};
    responseHeaders.forEach((header, i) => {
      if (header && requestRow[i]) {
        fullDetailsObject[header] = requestRow[i] instanceof Date ? Utilities.formatDate(requestRow[i], "Europe/London", "dd/MM/yyyy HH:mm:ss") : requestRow[i];
      }
    });
    const itemDescHeader = "Item Description - Answer";
    const itemDesc = requestRow[responseHeaders.indexOf(itemDescHeader)] || `Item from pass ${submissionId}`;
    const materialsSheet = ss.getSheetByName('Materials');
    const materialId = `MAT-${submissionId}-R${rowNumber}`;
    const description = `Original Submission ID: ${submissionId}. Tags: ${tags.trim()}`;
    materialsSheet.appendRow([
      materialId,
      itemDesc,
      description,
      tags.trim(),
      imageIds,
      JSON.stringify(fullDetailsObject)
    ]);
  } catch (e) {
    logAction('SYSTEM_ERROR', 'ARCHIVE_FAILURE', `Failed to archive material for item in row ${rowNumber}. Error: ${e.message}`);
  }
}

/**
 * UPDATED: Includes APM emails in CC for reassignment emails.
 */
function processReassignment(token, newApproverEmail) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tasksSheet = ss.getSheetByName('ApprovalTasks');
    const taskData = tasksSheet.getDataRange().getValues();
    const taskHeaders = taskData.shift();
    const oldTaskRowIndex = taskData.findIndex(row => row[taskHeaders.indexOf('ApprovalToken')] === token);
    if (oldTaskRowIndex === -1) throw new Error("Invalid token.");
    const oldTaskRow = taskData[oldTaskRowIndex];
    if (oldTaskRow[taskHeaders.indexOf('Status')] !== 'Active') throw new Error("This task has already been completed.");
    const submissionId = oldTaskRow[taskHeaders.indexOf('Submission_ID')];
    const originalApproverEmail = oldTaskRow[taskHeaders.indexOf('AssignedToEmail')];
    tasksSheet.getRange(oldTaskRowIndex + 2, taskHeaders.indexOf('Status') + 1).setValue('Reassigned');
    tasksSheet.getRange(oldTaskRowIndex + 2, taskHeaders.indexOf('CompletedTimestamp') + 1).setValue(new Date());
    const newToken = Utilities.getUuid();
    tasksSheet.appendRow([newToken, submissionId, newApproverEmail, new Date(), 'Active', '']);
    
    // Find siteId to get APM emails
    const responsesSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
    const responseData = responsesSheet.getDataRange().getValues();
    const responseHeaders = responseData.shift();
    const resIdCol = responseHeaders.indexOf('Submission_ID');
    const resLocCol = responseHeaders.indexOf('Location');
    const requestRow = responseData.find(r => r[resIdCol] === submissionId);
    const siteId = requestRow ? requestRow[resLocCol] : null;
    const apmEmails = getApmEmailsForSite_(siteId);

    const approvalUrl = `${ScriptApp.getService().getUrl()}?page=approve&token=${newToken}`;
    const subject = `ACTION REQUIRED: Gate Pass Request Re-assigned (${submissionId})`;
    const body = `<p>A gate pass request has been re-assigned to you for approval.</p><p><strong>Request ID:</strong> ${submissionId}</p><p>Please review the details and take action by clicking the link below:</p><p><a href="${approvalUrl}" style="font-size:16px; font-weight:bold; padding:10px 15px; background-color:#005f90; color:white; text-decoration:none; border-radius:5px;">Review Request</a></p>`;
    sendAppEmail({
      to: newApproverEmail,
      cc: apmEmails,
      subject: subject,
      htmlBody: body
    });
    logAction(originalApproverEmail, 'REASSIGN_APPROVAL', `Request ${submissionId} re-assigned to ${newApproverEmail}. CC: ${apmEmails || 'None'}`);
    return {
      success: true,
      message: `Task successfully re-assigned to ${newApproverEmail}.`
    };
  } finally {
    lock.releaseLock();
  }
}

function sendHtmlEmail(recipient, subject, plainBody, htmlContent, ccEmails) {
  var footerImage = '<img src="https://appscript-cdn.co.uk/nlhsse/ookmorgenveiliger.png" style="width:50%;" alt="Footer Image">';
  var htmlBody = htmlContent + '<br><br>' + footerImage;
  var options = {
    from: 'G4S NL Complaints <noreply@uk.g4s.com>',
    name: 'G4S NL Complaints',
    noReply: true,
    htmlBody: htmlBody
  };
  if (ccEmails) {
    options.cc = ccEmails;
  }
  MailApp.sendEmail(recipient, subject, plainBody, options);
}

function archiveLogs() {
  const LOG_ARCHIVE_FOLDER_ID = '1ZzgJYL7hV_mT9dzKrnrLGNekrsUHVtCw';
  const LOG_ARCHIVE_THRESHOLD = 40000;
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logsSheet = ss.getSheetByName(LOGS_SHEET);
    if (!logsSheet) {
      throw new Error(`Log sheet named '${LOGS_SHEET}' not found.`);
    }
    const lastRow = logsSheet.getLastRow();
    const logCount = lastRow - 1;
    if (logCount < LOG_ARCHIVE_THRESHOLD) {
      const message = `Log archival not needed. Current rows: ${logCount}, Threshold: ${LOG_ARCHIVE_THRESHOLD}.`;
      console.log(message);
      if (Session.getEffectiveUser().getEmail()) {
        ui.alert(message);
      }
      return;
    }
    console.log(`Threshold met. Archiving ${logCount} log entries.`);
    const dataToArchive = logsSheet.getRange(2, 1, logCount, logsSheet.getLastColumn()).getValues();
    const headerRow = logsSheet.getRange(1, 1, 1, logsSheet.getLastColumn()).getValues();
    const archiveFolder = DriveApp.getFolderById(LOG_ARCHIVE_FOLDER_ID);
    const timestamp = Utilities.formatDate(new Date(), "Europe/London", "yyyy-MM-dd");
    const archiveFileName = `Gate Pass Logs Archive - ${timestamp}`;
    const newSpreadsheet = SpreadsheetApp.create(archiveFileName);
    const newSheet = newSpreadsheet.getSheets()[0];
    newSheet.setName(LOGS_SHEET);
    console.log(`Copying data to new archive sheet: ${archiveFileName}`);
    newSheet.getRange(1, 1, 1, headerRow[0].length).setValues(headerRow).setFontWeight('bold');
    if (dataToArchive.length > 0) {
      newSheet.getRange(2, 1, dataToArchive.length, dataToArchive[0].length).setValues(dataToArchive);
    }
    SpreadsheetApp.flush();
    const newFile = DriveApp.getFileById(newSpreadsheet.getId());
    archiveFolder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);
    if (newSheet.getLastRow() === lastRow) {
      console.log("Verification successful. Clearing original logs.");
      logsSheet.getRange(2, 1, logCount, logsSheet.getLastColumn()).clearContent();
      const archiveUrl = newSpreadsheet.getUrl();
      const subject = "Gate Pass System Logs Successfully Archived";
      const body = `<p>This is an automated notification.</p>
                      <p>The System Action Logs for the Gate Pass application have been successfully archived to a new file, as the log file was nearing capacity.</p>
                      <p><b>Logs Archived:</b> ${logCount}</p>
                      <p>The new archive file can be accessed here: <a href="${archiveUrl}" style="font-weight:bold;">${archiveFileName}</a></p>
                      <p>The live log file has now been cleared for new entries.</p>`;
      const usersSheet = ss.getSheetByName(USERS_SHEET);
      const allUsers = usersSheet.getDataRange().getValues();
      const userHeaders = allUsers.shift();
      const roleIndex = userHeaders.indexOf('Role');
      const emailIndex = userHeaders.indexOf('Username');
      allUsers.forEach(userRow => {
        if (userRow[roleIndex] === ROLES.SUPER_USER) {
          sendAppEmail({ to: userRow[emailIndex], subject: subject, htmlBody: body });
        }
      });
      logAction('SYSTEM', 'LOGS_ARCHIVED', `Archived ${logCount} log entries to file: ${archiveFileName}`);
      if (Session.getEffectiveUser().getEmail()) {
        ui.alert('Success!', `Successfully archived ${logCount} log entries.`, ui.ButtonSet.OK);
      }
    } else {
      throw new Error("Verification failed. Row count in archive does not match original. Original logs were not cleared to prevent data loss.");
    }
  } catch (e) {
    console.error(`Log archival failed: ${e.stack}`);
    logAction('SYSTEM_ERROR', 'ARCHIVE_FAILURE', `Log archival failed. Error: ${e.message}`);
    sendAppEmail({ to: 'martin.bancroft@uk.g4s.com', subject: 'CRITICAL: Gate Pass Log Archival FAILED', htmlBody: `<p>The automated log archival process failed with the following error:</p><p><i>${e.message}</i></p><p>Please review the system logs immediately. No logs were cleared.</p>` });
    if (Session.getEffectiveUser().getEmail()) {
      ui.alert('Archival Failed', `An error occurred: ${e.message}`, ui.ButtonSet.OK);
    }
  }
}


// --- NEW FUNCTIONS FOR APPROVER MANAGEMENT AND REMINDERS ---

/**
 * NEW: Updates the 'ApproverListLastUpdated' timestamp for a site.
 * @param {string} siteId The ID of the site to update.
 */
function updateApproverTimestamp_(siteId) {
  if (!siteId) return;
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sitesSheet = ss.getSheetByName(SITES_SHEET);
    const data = sitesSheet.getRange("A:E").getValues();
    const headers = data.shift();
    const idColIdx = headers.indexOf('SiteID');
    const timestampColIdx = headers.indexOf('ApproverListLastUpdated');

    const rowIndex = data.findIndex(row => row[idColIdx] === siteId);
    if (rowIndex !== -1) {
      sitesSheet.getRange(rowIndex + 2, timestampColIdx + 1).setValue(new Date());
      CacheService.getScriptCache().remove('SITES_DATA');
    }
    lock.releaseLock();
  } catch (e) {
    logAction('SYSTEM_ERROR', 'TIMESTAMP_UPDATE_FAIL', `Failed to update approver timestamp for site ${siteId}. Error: ${e.message}`);
  }
}

/**
 * NEW: Adds a new approver and updates the timestamp.
 */
function addApprover(token, approverData) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  if (currentUser.role === ROLES.SITE_SUPERVISOR && approverData.siteId !== currentUser.siteId) {
    throw new Error('You can only add approvers to your own site.');
  }

  const approversSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Approvers');
  const newId = 'APP-' + Utilities.getUuid().substring(0, 8);
  approversSheet.appendRow([newId, approverData.fullName, approverData.email, approverData.siteId]);

  updateApproverTimestamp_(approverData.siteId);

  logAction(currentUser.username, 'ADD_APPROVER', `Added: ${approverData.fullName}`);
  CacheService.getScriptCache().remove('ALL_APPROVERS_DATA');
  return { success: true, message: 'Approver added successfully.' };
}

/**
 * NEW: Updates an existing approver and updates the timestamp(s).
 */
function updateApprover(token, approverData) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Approvers');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idIndex = headers.indexOf('ApproverID');

  const rowIndex = data.findIndex(row => row[idIndex] === approverData.approverId);
  if (rowIndex === -1) throw new Error('Approver not found.');

  const originalSiteId = data[rowIndex][headers.indexOf('SiteID')];
  if (currentUser.role === ROLES.SITE_SUPERVISOR && originalSiteId !== currentUser.siteId) {
    throw new Error('You can only edit approvers at your own site.');
  }

  sheet.getRange(rowIndex + 2, 2, 1, 3).setValues([[approverData.fullName, approverData.email, approverData.siteId]]);

  updateApproverTimestamp_(originalSiteId);
  if (originalSiteId !== approverData.siteId) {
    updateApproverTimestamp_(approverData.siteId);
  }

  logAction(currentUser.username, 'UPDATE_APPROVER', `Updated: ${approverData.fullName}`);
  CacheService.getScriptCache().remove('ALL_APPROVERS_DATA');
  return { success: true, message: 'Approver updated successfully.' };
}

/**
 * NEW: Deletes an approver and updates the timestamp.
 */
function deleteApprover(token, approverId) {
  const currentUser = getAuthenticatedUser(token);
  checkAdminPermissions_(currentUser);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Approvers');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idIndex = headers.indexOf('ApproverID');

  const rowIndex = data.findIndex(row => row[idIndex] === approverId);
  if (rowIndex === -1) throw new Error('Approver not found.');

  const approverToDelete = data[rowIndex];
  const siteId = approverToDelete[headers.indexOf('SiteID')];
  const fullName = approverToDelete[headers.indexOf('FullName')];

  if (currentUser.role === ROLES.SITE_SUPERVISOR && siteId !== currentUser.siteId) {
    throw new Error('You can only delete approvers at your own site.');
  }

  sheet.deleteRow(rowIndex + 2);

  updateApproverTimestamp_(siteId);

  logAction(currentUser.username, 'DELETE_APPROVER', `Deleted: ${fullName}`);
  CacheService.getScriptCache().remove('ALL_APPROVERS_DATA');
  return { success: true, message: 'Approver deleted successfully.' };
}

/**
 * NEW: Runs daily to check for stale approver lists and send reminders.
 */
function checkApproverListReminders() {
  const allSites = getSites_internal();
  const allUsersSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USERS_SHEET);
  const usersData = allUsersSheet.getDataRange().getValues();
  const userHeaders = usersData.shift();
  const userRoleCol = userHeaders.indexOf('Role');
  const userSiteCol = userHeaders.indexOf('SiteID');
  const userEmailCol = userHeaders.indexOf('Username');

  const NINETY_DAYS_IN_MS = 90 * 24 * 60 * 60 * 1000;
  const now = new Date().getTime();

  allSites.forEach(site => {
    const lastUpdated = site.approverListLastUpdated;
    if (lastUpdated) {
      const updatedTime = new Date(lastUpdated).getTime();
      if (now - updatedTime > NINETY_DAYS_IN_MS) {
        const supervisor = usersData.find(userRow =>
          userRow[userSiteCol] === site.siteId &&
          userRow[userRoleCol] === ROLES.SITE_SUPERVISOR
        );

        if (supervisor) {
          const supervisorEmail = supervisor[userEmailCol];
          const subject = `Reminder: Please Review Approver List for ${site.siteName}`;
          const body = `
            <p>Hello,</p>
            <p>This is an automated reminder from the Approva Gate Pass system.</p>
            <p>The list of authorized Shell Approvers for your site, <strong>${site.siteName}</strong>, has not been reviewed or updated in over 90 days.</p>
            <p>Please log in to the Admin Management panel to verify that the list is current and accurate.</p>
            <p>Thank you for helping to keep our system secure.</p>
          `;
          sendAppEmail({ to: supervisorEmail, subject: subject, htmlBody: body });
          logAction('SYSTEM', 'SENT_APPROVER_REMINDER', `Sent 90-day reminder for site ${site.siteId} to ${supervisorEmail}`);
        } else {
          logAction('SYSTEM_ERROR', 'APPROVER_REMINDER_FAIL', `Could not find a Site Supervisor for site ${site.siteId} to send reminder.`);
        }
      }
    } else {
      logAction('SYSTEM_INFO', 'APPROVER_REMINDER_SKIP', `Site ${site.siteId} has no 'ApproverListLastUpdated' timestamp. Skipping.`);
    }
  });
}
