/**
 * Regex patterns for detecting sensitive information
 */

// Email pattern: standard email format
export const EMAIL_PATTERN = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;

// Phone number patterns - supports multiple formats
// Matches: 123-456-7890, (123) 456-7890, +1 123 456 7890, 123.456.7890, 1234567890
export const PHONE_PATTERN = /(\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}\b/g;

// SSN patterns
// Matches: 123-45-6789 and 123456789
export const SSN_PATTERN = /\b\d{3}-\d{2}-\d{4}\b|\b\d{9}\b/g;

/**
 * Validate if a string matches email pattern
 */
export function isEmail(text: string): boolean {
    const emailRegex = /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}$/;
    return emailRegex.test(text);
}

/**
 * Validate if a string matches phone pattern
 */
export function isPhone(text: string): boolean {
    const phoneRegex = /^(\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$/;
    return phoneRegex.test(text);
}

/**
 * Validate if a string matches SSN pattern
 */
export function isSSN(text: string): boolean {
    const ssnRegex = /^\d{3}-\d{2}-\d{4}$|^\d{9}$/;
    return ssnRegex.test(text);
}

/**
 * Normalize phone number for searching
 * Removes all non-digit characters except leading +
 */
export function normalizePhone(phone: string): string {
    if (phone.startsWith('+')) {
        return '+' + phone.slice(1).replace(/\D/g, '');
    }
    return phone.replace(/\D/g, '');
}
