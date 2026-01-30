/**
 * Taskpane main logic - orchestrates redaction operations
 */

import { redactDocument } from '../word/redaction';
import { insertConfidentialHeader } from '../word/header';
import { enableTrackChanges } from '../word/tracking';
import type { OperationStatus } from '../types/results';

// UI Elements
let redactButton: HTMLButtonElement;
let loadingSection: HTMLElement;
let resultsSection: HTMLElement;
let errorSection: HTMLElement;
let emailCount: HTMLElement;
let phoneCount: HTMLElement;
let ssnCount: HTMLElement;
let headerStatus: HTMLElement;
let trackingStatus: HTMLElement;
let errorMessage: HTMLElement;
let dismissError: HTMLButtonElement;

/**
 * Initialize Office.js and set up event handlers
 */
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Get UI elements
        redactButton = document.getElementById('redactButton') as HTMLButtonElement;
        loadingSection = document.getElementById('loadingSection') as HTMLElement;
        resultsSection = document.getElementById('resultsSection') as HTMLElement;
        errorSection = document.getElementById('errorSection') as HTMLElement;
        emailCount = document.getElementById('emailCount') as HTMLElement;
        phoneCount = document.getElementById('phoneCount') as HTMLElement;
        ssnCount = document.getElementById('ssnCount') as HTMLElement;
        headerStatus = document.getElementById('headerStatus') as HTMLElement;
        trackingStatus = document.getElementById('trackingStatus') as HTMLElement;
        errorMessage = document.getElementById('errorMessage') as HTMLElement;
        dismissError = document.getElementById('dismissError') as HTMLButtonElement;

        // Set up event handlers
        redactButton.addEventListener('click', handleRedactClick);
        dismissError.addEventListener('click', hideError);

        console.log('Office Add-in initialized successfully');
    } else {
        showError('This add-in must be run in Microsoft Word.');
    }
});

/**
 * Handle redact button click
 */
async function handleRedactClick(): Promise<void> {
    try {
        showLoading();
        hideError();
        hideResults();

        // Step 1: Enable track changes first (if supported)
        const trackStatus = await enableTrackChanges();

        // Step 2: Insert confidential header
        const headerInserted = await insertConfidentialHeader();

        // Step 3: Perform redactions
        const redactionResults = await redactDocument();

        // Show results
        const status: OperationStatus = {
            headerInserted,
            trackChangesStatus: trackStatus,
            redactionResults,
        };

        showResults(status);
    } catch (error) {
        console.error('Error during redaction operation:', error);
        showError(
            error instanceof Error ? error.message : 'An unexpected error occurred during redaction.'
        );
    } finally {
        hideLoading();
    }
}

/**
 * Show loading state
 */
function showLoading(): void {
    redactButton.disabled = true;
    loadingSection.classList.remove('hidden');
}

/**
 * Hide loading state
 */
function hideLoading(): void {
    redactButton.disabled = false;
    loadingSection.classList.add('hidden');
}

/**
 * Show results
 */
function showResults(status: OperationStatus): void {
    emailCount.textContent = status.redactionResults.emailsRedacted.toString();
    phoneCount.textContent = status.redactionResults.phonesRedacted.toString();
    ssnCount.textContent = status.redactionResults.ssnsRedacted.toString();
    headerStatus.textContent = status.headerInserted ? 'Yes' : 'Already exists';

    // Format track changes status
    let trackText = '';
    switch (status.trackChangesStatus) {
        case 'enabled':
            trackText = 'Enabled ✓';
            break;
        case 'not_supported':
            trackText = 'Not Supported';
            break;
        case 'error':
            trackText = 'Error';
            break;
    }
    trackingStatus.textContent = trackText;

    resultsSection.classList.remove('hidden');
}

/**
 * Hide results
 */
function hideResults(): void {
    resultsSection.classList.add('hidden');
}

/**
 * Show error message
 */
function showError(message: string): void {
    errorMessage.textContent = message;
    errorSection.classList.remove('hidden');
}

/**
 * Hide error message
 */
function hideError(): void {
    errorSection.classList.add('hidden');
}
