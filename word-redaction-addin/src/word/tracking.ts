/**
 * Track Changes functionality for Word API 1.5+
 */

/**
 * Check if track changes is supported in the current Word environment
 */
export function isTrackChangesSupported(): boolean {
    try {
        return Office.context.requirements.isSetSupported('WordApi', '1.5');
    } catch (error) {
        console.error('Error checking track changes support:', error);
        return false;
    }
}

/**
 * Enable track changes if supported
 * @returns Status: 'enabled', 'not_supported', or 'error'
 */
export async function enableTrackChanges(): Promise<'enabled' | 'not_supported' | 'error'> {
    if (!isTrackChangesSupported()) {
        return 'not_supported';
    }

    try {
        await Word.run(async (context) => {
            // Enable track changes mode
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
        });
        return 'enabled';
    } catch (error) {
        console.error('Error enabling track changes:', error);
        return 'error';
    }
}
