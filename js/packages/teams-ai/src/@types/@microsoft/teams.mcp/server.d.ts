export interface ServerOptions {
    port: number;
}
export interface ResourceDefinition {
    uri: string;
    name: string;
    mimeType: string;
    read(): Promise<string>;
}
export class McpServer {
    constructor(options: ServerOptions);
    add(resource: ResourceDefinition): void;
    listen(callback: () => void): void;
}
