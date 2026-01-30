# SecureDoc - Word Redaction Add-in

A professional Word Add-in built with **Vite + TypeScript** that redacts sensitive information (emails, phone numbers, SSNs), adds a confidential header, and enables track changes when available.

## Features

✅ **Redact Sensitive Information**
- Emails (e.g., `john@example.com`)
- Phone numbers (multiple formats: `123-456-7890`, `(123) 456-7890`, `+1 123 456 7890`, `1234567890`)
- Social Security Numbers (e.g., `123-45-6789`, `123456789`)

✅ **Add Confidential Header**
- Inserts "CONFIDENTIAL DOCUMENT" at the top of the document
- Prevents duplicate headers
- Professionally styled (centered, bold, red)

✅ **Track Changes Support**
- Automatically enables Track Changes if Word API 1.5+ is available
- Graceful fallback for older Word versions
- All redactions and header insertions are tracked when available

✅ **Works Everywhere**
- Word on the Web
- Word Desktop (Windows, Mac)

## Prerequisites

- **Node.js** 18+ and npm
- **Microsoft Word** (Desktop or Web)
- **Microsoft 365** subscription (for Word on the Web)

## Installation

1. **Clone or download this repository**

2. **Install dependencies**
   ```bash
   cd word-redaction-addin
   npm install
   ```

3. **Start the development server**
   ```bash
   npm run dev
   ```

   The server will start on `https://localhost:3000` with a self-signed SSL certificate.

   > **Note**: You may see a browser warning about the self-signed certificate. This is normal for development. Click "Advanced" and proceed to localhost.

## Sideloading the Add-in

### Option 1: Word Desktop (Windows)

1. **Start the dev server** (if not already running):
   ```bash
   npm run dev
   ```

2. **Register the manifest**:
   - Open File Explorer and navigate to the project's `word-redaction-addin\public\` folder
   - Right-click on `manifest.xml` and select "Copy as path"
   - Open PowerShell as Administrator
   - Run:
     ```powershell
     # Replace <PATH> with the full path to manifest.xml
     New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Force
     New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Name "UseManifest" -Value "<PATH>\manifest.xml" -PropertyType String -Force
     ```

   Or use this automated registration script:

   ```powershell
   # Save as register-addin.ps1 in the project root
   $manifestPath = "$PSScriptRoot\public\manifest.xml"
   $regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
   
   New-Item -Path $regPath -Force | Out-Null
   New-ItemProperty -Path $regPath -Name "UseManifest" -Value $manifestPath -PropertyType String -Force | Out-Null
   
   Write-Host "Add-in registered successfully!" -ForegroundColor Green
   Write-Host "Restart Word to see the add-in."
   ```

   Then run:
   ```powershell
   .\register-addin.ps1
   ```

3. **Restart Word**
   - Close all Word windows
   - Open Word
   - Look for **"SecureDoc"** group on the **Home** ribbon tab
   - Click **"Show Taskpane"** to open the add-in

### Option 2: Word Desktop (Mac)

1. **Start the dev server**:
   ```bash
   npm run dev
   ```

2. **Sideload via Shared Folder**:
   - Create a folder: `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   - Copy `manifest.xml` to this folder
   - Restart Word
   - Go to **Insert** > **Add-ins** > **My Add-ins** > **Shared Folder**
   - Select "Word Redaction Add-in"

### Option 3: Word on the Web

1. **Start the dev server**:
   ```bash
   npm run dev
   ```

2. **Upload the manifest**:
   - Go to [Office.com](https://office.com) and sign in
   - Open Word Online
   - Create or open a document
   - Click **Insert** > **Add-ins** > **Upload My Add-in**
   - Click **Browse** and select `public/manifest.xml`
   - The add-in will appear in the ribbon

   > **Note**: For Word on the Web, you may need to update the manifest URLs to use a publicly accessible URL instead of `localhost` for production use.

## Usage

1. **Open a Word document** with sensitive information (or use the provided `Document-To-Be-Redacted.docx`)

2. **Open the add-in taskpane**:
   - Click **Home** > **SecureDoc** > **Show Taskpane**

3. **Click "Redact & Mark Confidential"**

4. **View results**:
   - Emails redacted count
   - Phone numbers redacted count
   - SSNs redacted count
   - Header insertion status
   - Track changes status

5. **Review the document**:
   - Sensitive information will be replaced with `[REDACTED EMAIL]`, `[REDACTED PHONE]`, or `[REDACTED SSN]`
   - Header section will display "CONFIDENTIAL DOCUMENT"
   - If Track Changes is supported, all changes will be tracked

## Testing with Sample Document

A sample document `Document-To-Be-Redacted.docx` should contain examples like:

```
Contact: john.doe@example.com
Phone: 123-456-7890 or (555) 123-4567
SSN: 123-45-6789
Alternative phone: +1 555 123 4567
Another email: jane.smith@company.org
```

After running the redaction, these will be replaced with appropriate redaction markers.

## Track Changes Availability

**Track Changes is enabled automatically** if your Word version supports **Word API 1.5+**:

- ✅ **Word 2016 or later** (Desktop)
- ✅ **Word Online** (modern versions)

If Track Changes is not supported:
- The add-in will still work perfectly
- Redactions and header insertion will proceed normally
- You'll see "Not Supported" in the results panel

## Project Structure

```
word-redaction-addin/
├── src/
│   ├── taskpane/
│   │   ├── index.html       # Taskpane UI
│   │   ├── taskpane.ts      # Main logic
│   │   └── taskpane.css     # Custom styles
│   ├── word/
│   │   ├── redaction.ts     # Redaction engine
│   │   ├── header.ts        # Header insertion
│   │   └── tracking.ts      # Track changes
│   ├── utils/
│   │   └── regex.ts         # Regex patterns
│   └── types/
│       └── results.ts       # TypeScript types
├── public/
│   ├── manifest.xml         # Office Add-in manifest
│   └── icon-*.png           # Icons
├── vite.config.ts           # Vite configuration
├── tsconfig.json            # TypeScript config
└── package.json             # Dependencies
```

## Building for Production

```bash
npm run build
```

The production build will be in the `dist/` folder. Update the manifest URLs to point to your production server.

## Troubleshooting

### Add-in doesn't appear in Word
- Ensure the dev server is running (`npm run dev`)
- Check that the manifest is registered correctly
- Restart Word completely
- Check the Windows Registry entry (for Windows Desktop)

### Certificate/HTTPS warnings
- This is normal for development with self-signed certificates
- Click "Advanced" and proceed to localhost
- For production, use a valid SSL certificate

### Track Changes not working
- Check if your Word version supports API 1.5+
- Track Changes requires Word 2016+ or modern Word Online
- The add-in will still work without Track Changes

### Redaction not finding all patterns
- The add-in uses Word's search API for robust pattern matching
- Some complex formats may require manual review
- You can run the redaction multiple times safely

## Development

### Run in development mode
```bash
npm run dev
```

### Build TypeScript
```bash
npm run build
```

### Lint and format
The project uses TypeScript strict mode for type safety.

## Technologies Used

- **Vite** - Fast build tool and dev server
- **TypeScript** - Type-safe development
- **Office.js** - Word JavaScript API
- **Custom CSS** - No frameworks, handcrafted styles
- **vite-plugin-mkcert** - Automatic HTTPS certificates

## License

MIT

## Support

For issues or questions, please check:
- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Word JavaScript API Reference](https://docs.microsoft.com/en-us/javascript/api/word)

---

**Built using Vite + TypeScript**
