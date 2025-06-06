import { McpServer, ResourceDefinition } from '@microsoft/teams.mcp';
import fetch from 'node-fetch';

const server = new McpServer({ port: 3001 });

const resource: ResourceDefinition = {
    uri: 'so://teams-questions',
    name: 'Latest Teams Q&A',
    mimeType: 'application/json',
    async read() {
        const url =
            'https://api.stackexchange.com/2.3/questions?tagged=microsoft-teams&pagesize=20&order=desc&sort=creation&site=stackoverflow&filter=withbody';
        try {
            const response = await fetch(url);
            if (!response.ok) {
                console.error(
                    `Failed to fetch questions: ${response.status} ${response.statusText}`
                );
                return JSON.stringify([]);
            }
            const json = await response.json();
            return JSON.stringify(json.items);
        } catch (err) {
            console.error('Error fetching questions:', err);
            return JSON.stringify([]);
        }
    }
};

server.add(resource);

server.listen(() => {
    console.log('MCP server listening on 3001');
});
