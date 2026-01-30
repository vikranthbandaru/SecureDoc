/**
 * Header insertion functionality
 */

const CONFIDENTIAL_TEXT = 'CONFIDENTIAL DOCUMENT';

/**
 * Check if confidential header already exists
 */
async function headerExists(context: Word.RequestContext): Promise<boolean> {
    try {
        const sections = context.document.sections;
        sections.load('items');
        await context.sync();

        if (sections.items.length === 0) {
            return false;
        }

        // Check the first section's primary header
        const header = sections.items[0].getHeader(Word.HeaderFooterType.primary);
        header.load('text');
        await context.sync();

        const headerText = header.text.trim().toUpperCase();
        return headerText.includes(CONFIDENTIAL_TEXT);
    } catch (error) {
        console.error('Error checking header existence:', error);
        return false;
    }
}

/**
 * Insert confidential header at the top of the document
 * @returns true if header was inserted, false if it already existed
 */
export async function insertConfidentialHeader(): Promise<boolean> {
    try {
        return await Word.run(async (context) => {
            // Check if header already exists
            const exists = await headerExists(context);
            if (exists) {
                console.log('Header already exists, skipping insertion');
                return false;
            }

            // Get the first section
            const sections = context.document.sections;
            sections.load('items');
            await context.sync();

            if (sections.items.length === 0) {
                throw new Error('No sections found in document');
            }

            // Get primary header
            const header = sections.items[0].getHeader(Word.HeaderFooterType.primary);

            // Clear existing content and insert new header
            header.clear();
            const headerParagraph = header.insertParagraph(CONFIDENTIAL_TEXT, Word.InsertLocation.start);

            // Style the header
            headerParagraph.alignment = Word.Alignment.centered;
            headerParagraph.font.size = 14;
            headerParagraph.font.bold = true;
            headerParagraph.font.color = '#D32F2F'; // Red color for confidential
            headerParagraph.spaceAfter = 12;

            await context.sync();
            console.log('Confidential header inserted successfully');
            return true;
        });
    } catch (error) {
        console.error('Error inserting confidential header:', error);
        throw error;
    }
}
