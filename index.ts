// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// <ProgramSnippet>
import * as readline from 'readline-sync';
import { DeviceCodeInfo } from '@azure/identity';
import {
  Message,
  TodoTaskList,
  TodoTask,
  Calendar,
  Event,
  Importance,
  TaskStatus,
} from '@microsoft/microsoft-graph-types';

import settings, { AppSettings } from './appSettings';
import * as graphHelper from './graphHelper';

const localTasks: [string, string][] = [];
let workingIdx: number = -1;

async function main() {
  console.log('TypeScript Graph Tutorial');

  let choice = 0;

  // Initialize Graph
  initializeGraph(settings);

  // Greet the user by name
  await greetUserAsync();

  const choices = [
    'Show my tasks today',
    'Start a task',
    'Stop a task',
    'List my task',
    'Summary today',
    'List my task lists',
    'List my tasks',
    'List my calendars',
    'List my events',
    'Add an event',
  ];

  while (choice != -1) {
    choice = readline.keyInSelect(choices, 'Select an option', {
      cancel: 'Exit',
    });

    switch (choice) {
      case -1:
        // Exit
        console.log('Goodbye...');
        break;
      case 0:
        await findTodayTasksAsync();
        break;
      case 1:
        await startTaskAsync();
        break;
      case 2:
        await stopTaskAsync();
        break;
      case 3:
        await displayTaskAsync();
        break;
      case 4:
        await summaryTheDayAsync();
        break;
      case 8:
        await listTaskListsAsync();
        break;
      case 9:
        await listTasksAsync(
          'AQMkADAwATM0MDAAMS1hNDAxLTY3NDMtMDACLTAwCgAuAAAD7VyvLT-NU0qZUjzrHbil7AEAvMe3PvoH7EWY6Ys_baveAQAEZrVURwAAAA==',
        );
        break;
      case 7:
        await listCalendarsAsync();
        break;
      case 5:
        await createEventAsync(
          'AQMkADAwATM0MDAAMS1hNDAxLTY3NDMtMDACLTAwCgBGAAAD7VyvLT-NU0qZUjzrHbil7AcAvMe3PvoH7EWY6Ys_baveAQAAAgEGAAAAvMe3PvoH7EWY6Ys_baveAQAAAjQ8AAAA',
          'TEST',
          '2024-09-13T23:00:00',
          '2024-09-13T24:00:00',
          'normal',
        );
        break;
      default:
        console.log('Invalid choice! Please try again.');
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
    console.log(`Email: ${user?.mail ?? user?.userPrincipalName ?? ''}`);
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
      console.log(`Message: ${message.subject ?? 'NO SUBJECT'}`);
      console.log(`  From: ${message.from?.emailAddress?.name ?? 'UNKNOWN'}`);
      console.log(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
      console.log(`  Received: ${message.receivedDateTime}`);
    }

    // If @odata.nextLink is not undefined, there are more messages
    // available on the server
    const moreAvailable = messagePage['@odata.nextLink'] != undefined;
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
      'Testing Microsoft Graph',
      'Hello world!',
      userEmail,
    );
    console.log('Mail sent.');
  } catch (err) {
    console.log(`Error sending mail: ${err}`);
  }
}
// </SendMailSnippet>


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
        `${task.id} ${task.title} ${task.status} ${task.categories} ${task.importance} ${task.startDateTime?.dateTime} ${task.startDateTime?.timeZone}`,
      );
    }
  } catch (err) {
    console.log(`Error get task lists: ${err}`);
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
  importance: Importance,
) {
  try {
    const newEvent = await graphHelper.createEventAsync(
      calendarID,
      subject,
      start,
      end,
      importance,
    );
    console.log(`Event created with id ${newEvent.id}`);
  } catch (err) {
    console.log(`Error create event: ${err}`);
  }
}

function getMidnightUTC(): string {
  const today = new Date();

  // Set the time to midnight (00:00:00) in UTC
  const midnight = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate(),
  );

  // Format the date as a string: YYYY-MM-DDTHH:mm:ss.ssssssZ
  return midnight.toISOString();
}

function getCurrentTime(): string {
  const now = new Date();
  return now.toISOString();
}

function calculateTotalDuration(
  text: string,
  isFinished: boolean = false,
  skipFilter: boolean = true,
): number {
  const lines: string[] = text.trim().split('\n');
  let totalMilliseconds = 0;

  for (const line of lines.slice(
    1,
    isFinished ? lines.length - 1 : lines.length,
  )) {
    // Split each line into start and end times
    const [startTimeStr, endTimeStr] = line.split(' ');

    // Parse the start and end times into Date objects
    const startTime = new Date(startTimeStr);
    const endTime = new Date(endTimeStr);

    // Calculate the difference in milliseconds
    if (skipFilter || isToday(startTimeStr)) {
      const durationMs = endTime.getTime() - startTime.getTime();

      // Add the duration to the total
      totalMilliseconds += Math.abs(durationMs);
    }
  }
  // Convert total duration from milliseconds to minutes
  const totalMinutes = totalMilliseconds / (1000 * 60);
  console.log(text, totalMinutes);
  return totalMinutes;
}

function extractEstimateTime(text: string): string {
  const lines = text.trim().split('\n');
  return lines.length ? lines[0] : '';
}

function isToday(utcDateStr: string): boolean {
  const date = new Date(utcDateStr);
  const now = new Date();
  return (
    date.getFullYear === now.getFullYear &&
    date.getMonth === now.getMonth &&
    date.getDay === now.getDay
  );
}

function calculateTodayDuration(text: string, isFinished: boolean): number {
  const lines: string[] = text.split('\n');
  const totalTodayMinutes = lines
    .slice(0, isFinished ? lines.length - 1 : lines.length)
    .reduce((acc, line) => {
      const [startTimeStr, endTimeStr] = line.split(' ');
      const startTime = new Date(startTimeStr);
      const endTime = new Date(endTimeStr);
      const durationMs = isToday(startTimeStr)
        ? endTime.getTime() - startTime.getTime()
        : 0;
      return acc + durationMs;
    }, 0);
  console.log(text, totalTodayMinutes / (1000 * 60));
  return totalTodayMinutes / (1000 * 60);
}

async function findTodayTasksAsync() {
  try {
    const taskListsPage = await graphHelper.getTaskListsAsync();
    const taskLists: TodoTaskList[] = taskListsPage.value;
    const midnight = getMidnightUTC();

    // Prepare an array of promises
    const tasksPromises = taskLists.map(async (taskList) => {
      if (!taskList.id) return { taskList, tasks: [] }; // Skip if taskList has no id
      const tasksPage = await graphHelper.getTasksAsync(
        taskList.id,
        `dueDateTime/dateTime eq '${midnight.slice(0, -1)}'`,
      );
      return { taskList, tasks: tasksPage.value }; // Return the tasks for this taskList
    });

    // Wait for all promises to resolve
    const allTasks = await Promise.all(tasksPromises);
    let idx = 1;
    // Process the tasks
    allTasks.forEach(({ taskList, tasks }) => {
      tasks.forEach((task) => {
        localTasks.push([
          taskList.id ? taskList.id : '',
          task.id ? task.id : '',
        ]);
        console.log(
          `${idx++} | ${taskList.displayName} | ${task.title} | ${task.importance} | ${task.status}`,
        );
      });
    });
  } catch (err) {
    console.log(`Error find today tasks: ${err}`);
  }
}

async function startTaskAsync() {
  if (workingIdx >= 0) {
    console.log(
      `You're working at ${workingIdx + 1}th task, stop it first before start a new task!`,
    );
    return;
  }
  const idx: number =
    Number(readline.question('Which task do you want to start: ')) - 1;
  if (idx >= localTasks.length || idx < 0) {
    console.log('The idx is out of the range, quit');
    return;
  }
  try {
    const [taskListId, taskId] = localTasks[idx];
    const task: TodoTask = await graphHelper.getTaskAsync(taskListId, taskId);
    const now: string = getCurrentTime();
    let estimateTime = '';
    if (task.status === 'inProgress') {
      console.log('The task has been started, continue...');
    } else if (task.status === 'completed') {
      console.log('The task has been completed, skip');
      return;
    } else {
      console.log("A new task. Let's eat it!");
      estimateTime = `${readline.question(
        'Please estimate the total time it need: ',
      )}`;
    }
    const oldBody: string = task.body?.content?.trim() || '';
    const body: string = `${oldBody}${estimateTime}\n${now}`;
    await graphHelper.updateTaskAsync(taskListId, taskId, body, 'inProgress');
    workingIdx = idx;
  } catch (err) {
    console.log(`Error start task: ${err}`);
  }
}

async function stopTaskAsync() {
  if (workingIdx < 0) {
    console.log('There is no working task');
    return;
  }
  try {
    const [taskListId, taskId] = localTasks[workingIdx];
    const task: TodoTask = await graphHelper.getTaskAsync(taskListId, taskId);
    const now: string = getCurrentTime();
    const isFinished: boolean =
      readline.question('Have you finished it? ') === 'y';
    const status: TaskStatus = isFinished ? 'completed' : 'inProgress';
    let body: string = `${task.body?.content?.trim() || ''} ${now}`;
    if (isFinished) {
      body = `${body}\n${calculateTotalDuration(body)}`;
    }
    await graphHelper.updateTaskAsync(taskListId, taskId, body, status);
    workingIdx = -1;
  } catch (err) {
    console.log(`Error stop task: ${err}`);
  }
}

async function displayTaskAsync() {
  try {
    const idx: number =
      Number(readline.question('Which task do you want to display: ')) - 1;
    if (idx >= localTasks.length || idx < 0) {
      console.log('The idx is out of the range, quit');
      return;
    }
    const [taskListId, taskId] = localTasks[idx];
    const task: TodoTask = await graphHelper.getTaskAsync(taskListId, taskId);
    console.log(task.title, task.body?.content, task.status);
  } catch (err) {
    console.log(`Error display task: ${err}`);
  }
}

async function summaryTheDayAsync() {
  if (workingIdx >= 0) {
    console.log(
      `You're working at ${workingIdx + 1}th task, stop it first before summary the day!`,
    );
    return;
  }

  const tasksPromises = localTasks.map(async ([taskListId, taskId]) => {
    const tasks = await graphHelper.getTaskAsync(taskListId, taskId);
    return tasks;
  });

  // Wait for all promises to resolve
  const allTasks: TodoTask[] = await Promise.all(tasksPromises);
  const tasksNum: number = allTasks.length;
  const finishedTasksNum: number = allTasks.filter((task) => {
    return task.status === 'completed';
  }).length;
  const doingTasksNum: number = allTasks.filter((task) => {
    return task.status === 'inProgress';
  }).length;
  const totalMinutes: number = allTasks.reduce((acc, task) => {
    return (
      acc +
      calculateTotalDuration(
        task.body?.content ? task.body?.content : '',
        task.status === 'completed',
        false,
      )
    );
  }, 0);
  const tasksInfo: string[] = allTasks.map((task) => {
    const info: string = `${task.title} | ${task.importance} | ${task.status} | ${extractEstimateTime(task.body?.content ? task.body?.content : '')}`;
  })
  console.log(`Good Job! Today you spent ${totalMinutes} minutes at tasks.`);
  console.log(`Today you have total ${tasksNum} tasks\n
  finished ${finishedTasksNum} tasks\n
  doing ${doingTasksNum} tasks\n
  left ${tasksNum - finishedTasksNum - doingTasksNum} tasks.`);
}
