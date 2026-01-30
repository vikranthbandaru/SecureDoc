/**
 * TypeScript interfaces for redaction results and operation status
 */

export interface RedactionResults {
    emailsRedacted: number;
    phonesRedacted: number;
    ssnsRedacted: number;
}

export interface OperationStatus {
    headerInserted: boolean;
    trackChangesStatus: 'enabled' | 'not_supported' | 'error';
    redactionResults: RedactionResults;
    error?: string;
}
