// Team Member name to be same as the name in google sheets
const WEBHOOKS = {
  'Team Member 1': '{Webhook 1}',
  'Team Member 2': '{Webhook 2}',
  'Team Member 3': '{Webhook 3}',
  'Team_Head': '{Webhook 4}',
};

function sendUnfilledFormReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('{Sheet_Name}');
  const lastRow = sheet.getLastRow();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Do not send reminders on Sunday
  if (today.getDay() === 0) {
    console.log('No reminders are sent on Sundays.');
    return;
  }

  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 10);
    const data = dataRange.getValues();

    const trainersFilledToday = new Set();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const timestamp = new Date(row[0]);
      timestamp.setHours(0, 0, 0, 0);
      const trainerName = row[2];

      if (today.getTime() === timestamp.getTime()) {
        trainersFilledToday.add(trainerName);
      }
    }

    // Exclude TrainerHead from receiving reminders
    const webhookUrls = { ...WEBHOOKS };
    delete webhookUrls['Team_Head'];

    // trainerName is the team member name in sheet
    for (const [trainerName, webhookUrl] of Object.entries(webhookUrls)) {
      if (!trainersFilledToday.has(trainerName)) {
        const message = `Reminder: Please ensure the EOD training spillage report is filled for today. 
        Here is the link: https://docs.google.com/forms/d/e/1FAIpQLSd3AGm2u2gyO_H-IUFTkX9rTB92JCkugsRTYo8Xz2MlTkhd8A/viewform`;
        sendMessageToChat(webhookUrl, message);
      }
    }
  } else {
    console.log('No data rows to process in sendUnfilledFormReminders.');
  }
}

function sendSpillageRemindersAndReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Crio.Do NHT Daily Training Report');
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 12);
    const data = dataRange.getValues();

    // Calculate the previous working day
    const previousWorkingDay = new Date();
    if (previousWorkingDay.getDay() === 1) { // If today is Monday
      previousWorkingDay.setDate(previousWorkingDay.getDate() - 2); // Set to the previous Saturday
    } else {
      previousWorkingDay.setDate(previousWorkingDay.getDate() - 1);
    }
    previousWorkingDay.setHours(0, 0, 0, 0);

    let report = 'Hi,\nKindly ensure that the following objectives are completed in training today.\n\n';

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const timestamp = new Date(row[0]);
      timestamp.setHours(0, 0, 0, 0);

      const topicSpilled = row[4];
      const spillageReminderSent = row[11];

      if (previousWorkingDay.getTime() === timestamp.getTime() && topicSpilled === 'Yes' && !spillageReminderSent) {
        const trainerName = row[2];
        const topic = row[5];
        const batchNumber = row[3];
        const webhookUrl = WEBHOOKS[trainerName];

        const message = `Reminder: Please ensure you cover the spilled topic: ${topic}`;
        sendMessageToChat(webhookUrl, message);
        
        sheet.getRange(2 + i, 12).setValue('Yes');

        report += `Batch number: ${batchNumber}\nTrainer Name: ${trainerName}\nTopic to be covered: ${topic}\n\n`;
      }
    }

    if (report !== 'Hi,\nKindly ensure that the following objectives are completed in training today.\n\n') {
      const trainerHeadWebhookUrl = WEBHOOKS['TrainerHead'];
      sendMessageToChat(trainerHeadWebhookUrl, report);
    }
  } else {
    console.log('No data rows to process in sendSpillageRemindersAndReport.');
  }
}

function sendMessageToChat(webhookUrl, message) {
  const payload = JSON.stringify({ text: message });
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: payload
  };
  UrlFetchApp.fetch(webhookUrl, options);
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('sendUnfilledFormReminders')
    .timeBased()
    .atHour(19).nearMinute(30)
    .inTimezone('Asia/Kolkata')
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('sendSpillageRemindersAndReport')
    .timeBased()
    .atHour(11).nearMinute(10)
    .inTimezone('Asia/Kolkata')
    .everyDays(1)
    .create();
}

function initialize() {
  setupTriggers();
}
