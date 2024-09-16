// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// <ProgramSnippet>
import * as readline from 'readline-sync';
import { DeviceCodeInfo } from '@azure/identity';
import {
  TodoTaskList,
  TodoTask,
  TaskStatus,
} from '@microsoft/microsoft-graph-types';
import { format, startOfDay } from 'date-fns';

import settings, { AppSettings } from './appSettings';
import * as graphHelper from './graphHelper';

const localTasks: [string, string][] = [];
const localTaskInfos: string[] = [];
let workingIdx: number = -1;
let defaultListId: string = '';

async function main() {
  let choice = 0;

  // Initialize Graph
  initializeGraph(settings);

  const choices = [
    'Show my tasks today',
    'Start a task',
    'Stop a task',
    'List a task',
    'Summary today',
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

function getMidnight(): string {
  const now = new Date();
  const midnightToday = startOfDay(now);
  return format(midnightToday, "yyyy-MM-dd'T'HH:mm:ss.SSS");
}

function getCurrentTime(): string {
  const now = new Date();
  return format(now, "yyyy-MM-dd'T'HH:mm:ss.SSS");
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
  return totalMinutes;
}

function extractEstimateTime(text: string): string {
  const lines = text.trim().split('\n');
  return lines.length ? lines[0] : '0';
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

async function findTodayTasksAsync() {
  try {
    if (workingIdx >= 0) {
      console.log(
        'Idx might change after pull, please finish the current task first',
      );
      return;
    }
    const taskListsPage = await graphHelper.getTaskListsAsync();
    const taskLists: TodoTaskList[] = taskListsPage.value;
    defaultListId =
      taskLists.find((list) => list.wellknownListName === 'defaultList')?.id ??
      '';
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
    // clear the localTasks first
    localTasks.length = 0;
    let idx = 1;
    // Process the tasks
    allTasks.forEach(({ taskList, tasks }) => {
      tasks.forEach((task) => {
        localTasks.push([taskList.id ?? '', task.id ?? '']);
        localTaskInfos.push(task.title ?? '');
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
  const idx: number = readline.keyInSelect(
    localTaskInfos,
    'Which task do you want to start: ',
  );
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
    const oldBody: string = task.body?.content?.trim() ?? '';
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
    const now: string = getCurrentTime();
    const [taskListId, taskId] = localTasks[workingIdx];
    const task: TodoTask = await graphHelper.getTaskAsync(taskListId, taskId);
    const taskList: TodoTaskList =
      await graphHelper.getTaskListAsync(taskListId);
    const listName: string = taskList.displayName ?? '';
    const isFinished: boolean =
      readline.question('Have you finished it? ') === 'y';
    const status: TaskStatus = isFinished ? 'completed' : 'inProgress';
    const oldBody: string = task.body?.content?.trim() ?? '';
    const startTimeStr: string = oldBody.split('\n').at(-1) ?? '';
    await graphHelper.createEventAsync(
      task.title ?? '',
      startTimeStr,
      now,
      task.importance ?? 'normal',
      [listName],
    );
    let body: string = `${oldBody} ${now}`;
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
  try {
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
    const headers: string[] = [
      'title',
      'importance',
      'status',
      'estimate time',
      'total duration (today)',
      'total duration (all)',
    ];
    const paddings: number[] = [40, 10, 10, 13, 22, 20];
    const tasksInfo: string[] = allTasks.map((task) => {
      return [
        task.title?.padEnd(paddings[0], ' '),
        task.importance?.padEnd(paddings[1], ' '),
        task.status?.padEnd(paddings[2], ' '),
        extractEstimateTime(task.body?.content ?? '').padEnd(paddings[3], ' '),
        calculateTotalDuration(
          task.body?.content ?? '',
          task.status === 'completed',
          false,
        )
          .toFixed(1)
          .padEnd(paddings[4], ' '),
        calculateTotalDuration(
          task.body?.content ?? '',
          task.status === 'completed',
          true,
        )
          .toFixed(1)
          .padEnd(paddings[5], ' '),
      ].join(' | ');
    });

    const summaryInfo = [
      `Good Job! Today you spent ${totalMinutes} minutes at tasks.`,
      `Today you have total ${tasksNum} tasks`,
      `  finished ${finishedTasksNum} tasks`,
      `  doing ${doingTasksNum} tasks`,
      `  left ${tasksNum - finishedTasksNum - doingTasksNum} tasks.`,
      headers
        .map((header, index) => `${header.padEnd(paddings[index], ' ')}`)
        .join(' | '),
      paddings.map((padding) => '-'.repeat(padding)).join(' | '),
    ];
    summaryInfo.push(...tasksInfo);
    const summaryText: string = summaryInfo.join('\n');
    console.log(summaryText);
    await graphHelper.createTaskAsync(
      defaultListId,
      'TodaySummary',
      summaryText,
      'inProgress',
      getMidnight(),
    );
  } catch (error) {
    console.error('Error summary today:', error);
  }
}
