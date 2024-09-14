// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// <UserAuthConfigSnippet>
import 'isomorphic-fetch';
import {
  DeviceCodeCredential,
  TokenCachePersistenceOptions,
  DeviceCodePromptCallback,
  useIdentityPlugin,
} from '@azure/identity';
import { cachePersistencePlugin } from '@azure/identity-cache-persistence';
import { Client, PageCollection } from '@microsoft/microsoft-graph-client';
import {
  User,
  Message,
  TodoTask,
  Event,
  Importance,
  TaskStatus,
} from '@microsoft/microsoft-graph-types';
// prettier-ignore
import { TokenCredentialAuthenticationProvider } from
  '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

import { AppSettings } from './appSettings';

let _settings: AppSettings | undefined = undefined;
let _deviceCodeCredential: DeviceCodeCredential | undefined = undefined;
let _userClient: Client | undefined = undefined;

export function initializeGraphForUserAuth(
  settings: AppSettings,
  deviceCodePrompt: DeviceCodePromptCallback,
) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;
  useIdentityPlugin(cachePersistencePlugin);
  const tokenCachePersistenceOptions: TokenCachePersistenceOptions = {
    enabled: true, // Enable persistent token caching
    name: 'msgraph', // Optional, default cache name, can be customized
    unsafeAllowUnencryptedStorage: true,
  };

  _deviceCodeCredential = new DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.tenantId,
    tokenCachePersistenceOptions,
    userPromptCallback: deviceCodePrompt,
  });

  const authProvider = new TokenCredentialAuthenticationProvider(
    _deviceCodeCredential,
    {
      scopes: settings.graphUserScopes,
    },
  );

  _userClient = Client.initWithMiddleware({
    authProvider: authProvider,
  });
}
// </UserAuthConfigSnippet>

// <GetUserTokenSnippet>
export async function getUserTokenAsync(): Promise<string> {
  // Ensure credential isn't undefined
  if (!_deviceCodeCredential) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Ensure scopes isn't undefined
  if (!_settings?.graphUserScopes) {
    throw new Error('Setting "scopes" cannot be undefined');
  }

  // Request token with given scopes
  const response = await _deviceCodeCredential.getToken(
    _settings?.graphUserScopes,
  );
  return response.token;
}
// </GetUserTokenSnippet>

// <GetUserSnippet>
export async function getUserAsync(): Promise<User> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Only request specific properties with .select()
  return _userClient
    .api('/me')
    .select(['displayName', 'mail', 'userPrincipalName'])
    .get();
}
// </GetUserSnippet>

// <GetInboxSnippet>
export async function getInboxAsync(): Promise<PageCollection> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  return _userClient
    .api('/me/mailFolders/inbox/messages')
    .select(['from', 'isRead', 'receivedDateTime', 'subject'])
    .top(25)
    .orderby('receivedDateTime DESC')
    .get();
}
// </GetInboxSnippet>

// <SendMailSnippet>
export async function sendMailAsync(
  subject: string,
  body: string,
  recipient: string,
) {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Create a new message
  const message: Message = {
    subject: subject,
    body: {
      content: body,
      contentType: 'text',
    },
    toRecipients: [
      {
        emailAddress: {
          address: recipient,
        },
      },
    ],
  };

  // Send the message
  return _userClient.api('me/sendMail').post({ message: message });
}
// </SendMailSnippet>

// <MakeGraphCallSnippet>
// This function serves as a playground for testing Graph snippets
// or other code
export async function makeGraphCallAsync() {
  // INSERT YOUR CODE HERE
}
// </MakeGraphCallSnippet>

export async function getTaskListsAsync(): Promise<PageCollection> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  return _userClient.api('me/todo/lists').get();
}

export async function getTasksAsync(
  taskListID: string,
  filter: string = '',
): Promise<PageCollection> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  return _userClient
    .api(`me/todo/lists/${taskListID}/tasks`)
    .filter(filter)
    .get();
}

export async function getTaskAsync(
  taskListID: string,
  taskID: string,
): Promise<TodoTask> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  return _userClient.api(`me/todo/lists/${taskListID}/tasks/${taskID}`).get();
}

export async function createTaskAsync(
  taskListID: string,
  title: string,
  content: string | undefined = undefined,
  status: TaskStatus = 'notStarted',
  dueDateTime: string | undefined = undefined,
): Promise<TodoTask> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }

  const task: TodoTask = {
    title: title,
    body: {
      content: content,
      contentType: 'text',
    },
    dueDateTime: {
      dateTime: dueDateTime,
      timeZone: 'UTZ',
    },
    status: status,
  };
  return _userClient.api(`me/todo/lists/${taskListID}/tasks`).post(task);
}

export async function updateTaskAsync(
  taskListID: string,
  taskID: string,
  body: string,
  status: TaskStatus = 'notStarted',
): Promise<TodoTask> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  const updatedTask: TodoTask = {
    body: { content: body, contentType: 'text' },
    status: status,
  };
  return _userClient
    .api(`me/todo/lists/${taskListID}/tasks/${taskID}`)
    .update(updatedTask);
}

export async function getCalendarsAsync(): Promise<PageCollection> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  return _userClient.api('me/calendars').get();
}

export async function getEventsAsync(
  calendarID: string,
): Promise<PageCollection> {
  // Ensure client isn't undefined
  if (!_userClient) {
    throw new Error('Graph has not been initialized for user auth');
  }
  return _userClient.api(`me/calendars/${calendarID}/events`).get();
}

export async function createEventAsync(
  calendarID: string,
  subject: string,
  start: string,
  end: string,
  importance: Importance,
) {
  const event: Event = {
    subject: subject,
    start: { dateTime: start, timeZone: 'UTC' },
    end: { dateTime: end, timeZone: 'UTC' },
    importance: importance,
  };
  return _userClient?.api(`me/calendars/${calendarID}/events`).post(event);
}
