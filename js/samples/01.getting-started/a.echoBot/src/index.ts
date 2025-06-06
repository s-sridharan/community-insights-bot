// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.


// Import required packages
import { config } from 'dotenv';
import * as path from 'path';
import * as restify from 'restify';

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { ActivityTypes, ConfigurationServiceClientCredentialFactory, MemoryStorage, TurnContext } from 'botbuilder';

import { Application, TurnState, TeamsAdapter } from '@microsoft/teams-ai';

/*********************************************************************
 *  ONE-OFF AUTH DIAGNOSTIC  – remove after 401 is solved
 *********************************************************************/
import fs from "fs";
import { fileURLToPath } from "url";

//   1 ) Load whichever .env the process can see
config();   // ← simpler than the custom path

//   2 ) Dump the auth settings we’ll actually use
const dump = {
  PWD: process.cwd(),
  ENV_FOUND: fs.existsSync(".env"),
  BOT_ID: process.env.BOT_ID,
  BOT_ID_LEN: process.env.BOT_ID?.length || 0,
  BOT_PASSWORD_LEN: process.env.BOT_PASSWORD?.length || 0,
  BOT_TENANT_ID: process.env.BOT_TENANT_ID,
  APP_TYPE: "SingleTenant",
};
console.table(dump);

//   3 ) Double-check manifest & endpoint
console.log(
  "Messaging endpoint expected:",
  `https://${process.env.PUBLIC_HOST || "<codespace-hash>-3978.app.github.dev"}/api/messages`
);
/*********************************************************************/


// Read botFilePath and botFileSecret from .env file.
const ENV_FILE = path.join(__dirname, '..', '.env');
config({ path: ENV_FILE });

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
// const adapter = new TeamsAdapter(
//     {},
//     new ConfigurationServiceClientCredentialFactory({
//         MicrosoftAppId: process.env.BOT_ID,
//         MicrosoftAppPassword: process.env.BOT_PASSWORD,
//         MicrosoftAppType: 'MultiTenant'
//     })
// );

const adapter = new TeamsAdapter(
  {},
  new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId:       process.env.BOT_ID,
    MicrosoftAppPassword: process.env.BOT_PASSWORD,
    MicrosoftAppType:     "SingleTenant",              // <- explicit
    MicrosoftAppTenantId: process.env.BOT_TENANT_ID    // <- target tenant
  })
);


// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: any) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error}`);
    console.log(error);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
    console.log('\nTo test your bot in Teams, sideload the app manifest.json within Teams Apps.');
});

interface ConversationState {
    count: number;
}
type ApplicationTurnState = TurnState<ConversationState>;

// Define storage and application
const storage = new MemoryStorage();
const app = new Application<ApplicationTurnState>({
    storage
});

// Listen for user to say '/reset' and then delete conversation state
app.message('/reset', async (context: TurnContext, state: ApplicationTurnState) => {
    state.deleteConversationState();
    await context.sendActivity(`Ok I've deleted the current conversation state.`);
});

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
app.activity(ActivityTypes.Message, async (context: TurnContext, state: ApplicationTurnState) => {
    // Increment count state
    let count = state.conversation.count ?? 0;
    state.conversation.count = ++count;

    // Echo back users request
    await context.sendActivity(`[${count}] you said: ${context.activity.text}`);
});

server.get('/', (_req, res, next) => {
    res.send(200, 'OK');
    next();
});

// Listen for incoming server requests.
server.post('/api/messages', async (req, res) => {
    // Route received a request to adapter for processing
    await adapter.process(req, res as any, async (context) => {
        // Dispatch to application for routing
        await app.run(context);
    });
});
