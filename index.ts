// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// <ProgramSnippet>
import * as readline from "readline-sync";
import { DeviceCodeInfo } from "@azure/identity";
import {
  Message,
  TodoTaskList,
  TodoTask,
  Calendar,
  Event,
} from "@microsoft/microsoft-graph-types";

import settings, { AppSettings } from "./appSettings";
import * as graphHelper from "./graphHelper";

async function main() {
  console.log("TypeScript Graph Tutorial");

  let choice = 0;

  // Initialize Graph
  initializeGraph(settings);

  // Greet the user by name
  await greetUserAsync();

  const choices = [
    "Display access token",
    "List my inbox",
    "Send mail",
    "Make a Graph call",
    "List my task lists",
    "List my tasks",
    "List my task",
    "List my calendars",
    "List my events",
    "Add an event",
  ];

  while (choice != -1) {
    choice = readline.keyInSelect(choices, "Select an option", {
      cancel: "Exit",
    });

    switch (choice) {
      case -1:
        // Exit
        console.log("Goodbye...");
        break;
      case 0:
        // Display access token
        await displayAccessTokenAsync();
        break;
      case 1:
        // List emails from user's inbox
        await listInboxAsync();
        break;
      case 2:
        // Send an email message
        await sendMailAsync();
        break;
      case 3:
        // Run any Graph code
        await makeGraphCallAsync();
        break;
      case 4:
        await listTaskListsAsync();
        break;
      case 5:
        await listTasksAsync(
          "AQMkADAwATM0MDAAMS1hNDAxLTY3NDMtMDACLTAwCgAuAAAD7VyvLT-NU0qZUjzrHbil7AEAvMe3PvoH7EWY6Ys_baveAQAEZrVURwAAAA=="
        );
        break;
      case 6:
        await displayTaskAsync(
          "AQMkADAwATM0MDAAMS1hNDAxLTY3NDMtMDACLTAwCgAuAAAD7VyvLT-NU0qZUjzrHbil7AEAvMe3PvoH7EWY6Ys_baveAQAEZrVURwAAAA==",
          "AQMkADAwATM0MDAAMS1hNDAxLTY3NDMtMDACLTAwCgBGAAAD7VyvLT-NU0qZUjzrHbil7AcAvMe3PvoH7EWY6Ys_baveAQAEZrVURwAAALzHtz76B_xFmOmLPm2r3gEAB4QQwBsAAAA="
        );
        break;
      case 7:
        await listCalendarsAsync();
        break;
      case 8:
        await listEventsAsync(
          "AQMkADAwATM0MDAAMS1hNDAxLTY3NDMtMDACLTAwCgBGAAAD7VyvLT-NU0qZUjzrHbil7AcAvMe3PvoH7EWY6Ys_baveAQAAAgEGAAAAvMe3PvoH7EWY6Ys_baveAQAAAjQ8AAAA"
        );
        break;
      case 9:
        await createEventAsync(
          "AQMkADAwATM0MDAAMS1hNDAxLTY3NDMtMDACLTAwCgBGAAAD7VyvLT-NU0qZUjzrHbil7AcAvMe3PvoH7EWY6Ys_baveAQAAAgEGAAAAvMe3PvoH7EWY6Ys_baveAQAAAjQ8AAAA",
          "TEST",
          "2024-09-13T23:00:00",
          "2024-09-13T24:00:00",
          "normal"
        );
        break;
      default:
        console.log("Invalid choice! Please try again.");
    }
  }
}

main();
// </ProgramSnippet>

// <InitializeGraphSnippet>
function initializeGraph(settings: AppSettings) {
  graphHelper.initializeGraphForUserAuth(settings, (info: DeviceCodeInfo) => {
    // Display the device code message to
    // the user. This tells them
    // where to go to sign in and provides the
    // code to use.
    console.log(info.message);
  });
}
// </InitializeGraphSnippet>

// <GreetUserSnippet>
async function greetUserAsync() {
  try {
    const user = await graphHelper.getUserAsync();
    console.log(`Hello, ${user?.displayName}!`);
    // For Work/school accounts, email is in mail property
    // Personal accounts, email is in userPrincipalName
    console.log(`Email: ${user?.mail ?? user?.userPrincipalName ?? ""}`);
  } catch (err) {
    console.log(`Error getting user: ${err}`);
  }
}
// </GreetUserSnippet>

// <DisplayAccessTokenSnippet>
async function displayAccessTokenAsync() {
  try {
    const userToken = await graphHelper.getUserTokenAsync();
    console.log(`User token: ${userToken}`);
  } catch (err) {
    console.log(`Error getting user access token: ${err}`);
  }
}
// </DisplayAccessTokenSnippet>

// <ListInboxSnippet>
async function listInboxAsync() {
  try {
    const messagePage = await graphHelper.getInboxAsync();
    const messages: Message[] = messagePage.value;

    // Output each message's details
    for (const message of messages) {
      console.log(`Message: ${message.subject ?? "NO SUBJECT"}`);
      console.log(`  From: ${message.from?.emailAddress?.name ?? "UNKNOWN"}`);
      console.log(`  Status: ${message.isRead ? "Read" : "Unread"}`);
      console.log(`  Received: ${message.receivedDateTime}`);
    }

    // If @odata.nextLink is not undefined, there are more messages
    // available on the server
    const moreAvailable = messagePage["@odata.nextLink"] != undefined;
    console.log(`\nMore messages available? ${moreAvailable}`);
  } catch (err) {
    console.log(`Error getting user's inbox: ${err}`);
  }
}
// </ListInboxSnippet>

// <SendMailSnippet>
async function sendMailAsync() {
  try {
    // Send mail to the signed-in user
    // Get the user for their email address
    const user = await graphHelper.getUserAsync();
    const userEmail = user?.mail ?? user?.userPrincipalName;

    if (!userEmail) {
      console.log("Couldn't get your email address, canceling...");
      return;
    }

    await graphHelper.sendMailAsync(
      "Testing Microsoft Graph",
      "Hello world!",
      userEmail
    );
    console.log("Mail sent.");
  } catch (err) {
    console.log(`Error sending mail: ${err}`);
  }
}
// </SendMailSnippet>

// <MakeGraphCallSnippet>
async function makeGraphCallAsync() {
  try {
    await graphHelper.makeGraphCallAsync();
  } catch (err) {
    console.log(`Error making Graph call: ${err}`);
  }
}
// </MakeGraphCallSnippet>

async function listTaskListsAsync() {
  try {
    const taskListsPage = await graphHelper.getTaskListsAsync();
    const taskLists: TodoTaskList[] = taskListsPage.value;

    for (const taskList of taskLists) {
      console.log(`${taskList.id} ${taskList.displayName}`);
    }
  } catch (err) {
    console.log(`Error get task lists: ${err}`);
  }
}

async function listTasksAsync(taskListID: string) {
  try {
    const tasksPage = await graphHelper.getTasksAsync(taskListID);
    const tasks: TodoTask[] = tasksPage.value;

    for (const task of tasks) {
      console.log(
        `${task.id} ${task.title} ${task.status} ${task.categories} ${task.importance} ${task.startDateTime?.dateTime} ${task.startDateTime?.timeZone}`
      );
    }
  } catch (err) {
    console.log(`Error get task lists: ${err}`);
  }
}

async function displayTaskAsync(taskListID: string, taskID: string) {
  try {
    const task = await graphHelper.getTaskAsync(taskListID, taskID);
    if (task.extensions) console.log(Object.keys(task.extensions[0]));
    else console.log("No extensions");
  } catch (err) {
    console.log(`Error get task: ${err}`);
  }
}

async function listCalendarsAsync() {
  try {
    const calendarsPage = await graphHelper.getCalendarsAsync();
    const calendars: Calendar[] = calendarsPage.value;
    for (const calendar of calendars) {
      console.log(`${calendar.id} ${calendar.name}`);
    }
  } catch (err) {
    console.log(`Error get calendars: ${err}`);
  }
}

async function listEventsAsync(calendarID: string) {
  try {
    const eventsPage = await graphHelper.getEventsAsync(calendarID);
    const events: Event[] = eventsPage.value;
    for (const event of events) {
      console.log(`${event.id} ${event.subject}`);
    }
  } catch (err) {
    console.log(`Error get events: ${err}`);
  }
}

async function createEventAsync(
  calendarID: string,
  subject: string,
  start: string,
  end: string,
  importance: string
) {
  try {
    const newEvent = await graphHelper.createEventAsync(
      calendarID,
      subject,
      start,
      end,
      importance
    );
    console.log(`Event created with id ${newEvent.id}`);
  } catch (err) {
    console.log(`Error create event: ${err}`);
  }
}

async function findTodayTasksAsync() {
  try {
    const taskListsPage = await graphHelper.getTaskListsAsync();
    const taskLists: TodoTaskList[] = taskListsPage.value;

    for (const taskList of taskLists) {
      const tasksPage = await graphHelper.getTasksAsync(
        taskList.id ? taskList.id : ""
      );

      const tasks: TodoTask[] = tasksPage.value;
      for (const task of tasks) {
        if (task.status === "notStarted" && task.dueDateTime?.dateTime === "") {
          console.log(
            `${task.id} ${task.title} ${task.status} ${task.categories} ${task.importance} ${task.startDateTime?.dateTime} ${task.startDateTime?.timeZone}`
          );
        }
      }
    }
  } catch (err) {
    console.log(`Error get task lists: ${err}`);
  }
}
