// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// <SettingsSnippet>
const settings: AppSettings = {
  clientId: '72000ad4-3e42-4653-ac5e-e6bc7d28c773',
  tenantId: 'common',
  graphUserScopes: [
    'user.read',
    'mail.readwrite',
    'mail.send',
    'tasks.readwrite',
    'calendars.readwrite',
    'contacts.readwrite',
  ],
};

export interface AppSettings {
  clientId: string;
  tenantId: string;
  graphUserScopes: string[];
}

export default settings;
// </SettingsSnippet>
