# SecureDoc - Professional Document Redaction Add-in

A powerful Word Add-in that automatically redacts sensitive information, adds confidentiality headers, and enables tracking changes to maintain document security and compliance.

## What it does?

1. **Redact Sensitive Information**
    - Retrieve the document's complete content
    - Locate and identify sensitive information (emails, phone numbers, social security numbers)
    - Replace this information with redaction markers in the document
2. **Add Confidential Header**
    - Insert a header at the top of the document stating "CONFIDENTIAL DOCUMENT"
    - Ensure this header addition is tracked by the Tracking Changes feature
3. **Enable Tracking Changes**
    - Use the Office Tracking Changes API to enable tracking changes
    - Make sure to only use Tracking Changes if the Word API is available
    [Word JavaScript API requirement set 1.5 - Office Add-ins | Microsoft Learn](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-1-5-requirement-set?view=word-js-preview)
## Demo Link
https://drive.google.com/file/d/1z4Ijfcg4Ly16D8BzxiKr5ZjsTAv0MkNq/view?usp=sharing
## Technical Stack
    
- **TypeScript** - Type-safe development
- **Vite** - Fast build tool and modern development server  
- **Custom CSS** - Handcrafted styles for a premium user experience
- **Office.js** - Word JavaScript API for cross-platform compatibility
- Works in **Word on the Web** and **Word Desktop** (Windows/Mac)


## Testing Your Solution
Use the attached Document-To-Be-Redacted.docx file to test your solution. The document contains various instances of sensitive information that should be redacted when your add-in is executed.


## Getting Started

1. Cloning the repository to your local machine to get started is recommended.
2. `npm install`
3. `npm start`
   - Starts local server on port 3000.
   - Compiles TypeScript.
   - Attempts to sideload to Word.

If automatic sideloading fails, please [sideload the manifest manually](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).


## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License
