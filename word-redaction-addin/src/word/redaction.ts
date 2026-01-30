/**
 * Redaction functionality for sensitive information
 */

import type { RedactionResults } from '../types/results';

const REDACTED_EMAIL = '[REDACTED EMAIL]';
const REDACTED_PHONE = '[REDACTED PHONE]';
const REDACTED_SSN = '[REDACTED SSN]';

// Email pattern
const EMAIL_REGEX = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/gi;

// Phone patterns - various formats
const PHONE_REGEX = /(\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}\b/g;

// SSN patterns
const SSN_REGEX = /\b\d{3}-\d{2}-\d{4}\b/g;

/**
 * Redact patterns using direct text search
 */
async function redactPattern(
    context: Word.RequestContext,
    pattern: RegExp,
    replacement: string
): Promise<number> {
    let count = 0;

    try {
        // Get all paragraphs in the document body
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        // Process each paragraph
        for (const paragraph of paragraphs.items) {
            paragraph.load('text');
            await context.sync();

            const text = paragraph.text;
            const matches = Array.from(text.matchAll(pattern));

            if (matches.length > 0) {
                // For each match, search and replace it
                for (const match of matches) {
                    const matchText = match[0];

                    // Search for this exact text in the paragraph
                    const searchResults = paragraph.search(matchText, {
                        matchCase: false,
                        matchWholeWord: false,
                    });
                    searchResults.load('items');
                    await context.sync();

                    // Replace all occurrences
                    for (const result of searchResults.items) {
                        result.insertText(replacement, Word.InsertLocation.replace);
                        count++;
                    }
                    await context.sync();
                }
            }
        }
    } catch (error) {
        console.error('Error in redactPattern:', error);
    }

    return count;
}

/**
 * Main redaction function - redacts all sensitive information
 */
export async function redactDocument(): Promise<RedactionResults> {
    try {
        return await Word.run(async (context) => {
            // Perform redactions in sequence to avoid conflicts
            // SSNs first as they're most specific
            const ssnsRedacted = await redactPattern(context, SSN_REGEX, REDACTED_SSN);

            // Then emails
            const emailsRedacted = await redactPattern(context, EMAIL_REGEX, REDACTED_EMAIL);

            // Then phones (last to avoid conflicting with SSNs)
            const phonesRedacted = await redactPattern(context, PHONE_REGEX, REDACTED_PHONE);

            return {
                emailsRedacted,
                phonesRedacted,
                ssnsRedacted,
            };
        });
    } catch (error) {
        console.error('Error in redactDocument:', error);
        throw error;
    }
}
