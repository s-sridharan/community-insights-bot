/**
 * @module teams-ai
 */
/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
import "dotenv/config";
import { spawn } from 'child_process';
import { McpClientPlugin } from '@microsoft/teams.mcpclient';
export * from './augmentations';
export * from './dataSources';
export * from './embeddings';
export * from './external';
export * from './models';
export * from './moderators';
export * from './planners';
export * from './prompts';
export * from './tokenizers';
export * from './authentication';
export * from './validators';
export * from './AdaptiveCards';
export * from './AI';
export * from './Application';
export * from './ApplicationBuilder';
export * from './InputFileDownloader';
export * from './MemoryFork';
export * from './MessageExtensions';
export * from './StreamingResponse';
export * from './TaskModules';
export * from './TeamsAttachmentDownloader';
export * from './TurnState';
export * from './Utilities';
export * from './authentication/TeamsBotSsoPrompt';
export * from './TeamsAdapter';
export { ActionHandler, PredictedDoCommandAndHandler } from './actions';

// start local MCP server for fetching Stack Overflow questions
spawn('tsx', ['src/mcp/so-server.ts'], { stdio: 'inherit' });
export const mcp = new McpClientPlugin({ url: 'http://localhost:3001/mcp' });
export { McpClientPlugin };
