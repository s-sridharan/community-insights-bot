export interface ResourceDefinition {
    uri: string;
    name: string;
    mimeType: string;
    read(): Promise<string>;
}
