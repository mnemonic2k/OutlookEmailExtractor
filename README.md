# Outlook Email Extractor

A JavaScript tool for extracting email data from Outlook Web App (OWA) using browser developer console. This script automatically captures email content, subjects, senders, and timestamps as you browse through your emails.

> **⚠️ IMPORTANT**: For corporate/work accounts, obtain explicit permission from your IT department before use. For personal accounts, you can use this tool freely.

## Features

- 🔍 **Automatic Email Detection**: Monitors DOM changes and automatically captures emails when opened
- 📧 **Complete Email Data**: Extracts subject, sender, date, content, and URL
- 📊 **Multiple Export Formats**: Download data as JSON or CSV files
- 🚫 **Duplicate Prevention**: Prevents saving the same email multiple times
- 🐛 **Debug Mode**: Built-in debugging tools for troubleshooting
- 📁 **Multi-Folder Support**: Works in inbox and custom folders
- 🧹 **Content Cleaning**: Removes HTML comments and CSS artifacts from email content

## Installation & Usage

### Step 1: Access Outlook Web App
Navigate to your Outlook Web App (e.g., `outlook.office.com`)

### Step 2: Open Developer Console
- **Chrome/Edge**: Press `F12` or `Ctrl+Shift+I`
- **Firefox**: Press `F12` or `Ctrl+Shift+K`
- **Safari**: Press `Cmd+Option+I`

### Step 3: Run the Script
1. Copy the entire script from `outlook-email-extractor.js`
2. Paste it into the browser console
3. Press `Enter` to execute

### Step 4: Start Browsing Emails
The script will automatically start monitoring and capturing emails as you open them.

## Available Commands

Once the script is running, you can use these commands in the console:

```javascript
// View all captured emails in a table format
exportEmails.show()

// Manually capture the currently open email
exportEmails.capture()

// Enable debug mode for troubleshooting
exportEmails.debug()

// Download emails as JSON file
exportEmails.downloadJSON()

// Download emails as CSV file (Excel-compatible)
exportEmails.downloadCSV()

// Clear all captured emails from memory
exportEmails.clear()

// Stop the automatic monitoring
outlookEmailExtractor.stop()
```

## Data Structure

Each captured email contains the following fields:

```json
{
  "subject": "Email Subject",
  "sender": "Sender Name",
  "date": "2025-07-18T10:30:00Z",
  "content": "Email content text...",
  "timestamp": "2025-07-18T10:30:15.123Z",
  "url": "https://outlook.office.com/mail/..."
}
```

## Troubleshooting

### Email Not Being Captured

1. Enable debug mode:
   ```javascript
   exportEmails.debug()
   ```

2. Try manual capture:
   ```javascript
   exportEmails.capture()
   ```

3. Check console output for error messages and DOM structure

### Working in Different Folders

The script automatically detects emails in:
- Inbox
- Custom folders
- Shared mailboxes

If emails in specific folders aren't being captured, use debug mode to check URL patterns.

### Browser Compatibility

Tested and working with:
- ✅ Chrome 90+
- ✅ Microsoft Edge 90+
- ✅ Firefox 88+
- ✅ Safari 14+

## Security & Privacy

- **Client-Side Only**: All processing happens in your browser
- **No Data Transmission**: Email data stays on your device
- **No External Dependencies**: Pure JavaScript, no third-party libraries
- **Temporary Storage**: Data is stored in browser memory only

## ⚠️ Important Usage Guidelines

**AUTHORIZATION REQUIRED**: Before using this tool, ensure you have proper authorization:

### Corporate/Work Accounts
- ✅ **GET EXPLICIT PERMISSION** from your IT department or supervisor before use
- ✅ **CHECK COMPANY POLICIES** regarding browser automation and data extraction tools
- ✅ **VERIFY COMPLIANCE** with your organization's security and data handling policies
- ❌ **DO NOT USE** on company systems without proper authorization

### Personal Accounts
- ✅ **SAFE TO USE** on your personal Outlook account
- ✅ **YOUR OWN DATA** - you have full rights to extract your personal emails
- ✅ **NO RESTRICTIONS** for personal use cases

### Legal Considerations
- **Data Ownership**: Only extract emails you have legal rights to access
- **Privacy Laws**: Comply with local data protection regulations (GDPR, CCPA, etc.)
- **Terms of Service**: Ensure usage complies with Microsoft's Terms of Service
- **Third-Party Content**: Be mindful of emails containing confidential information from others

### Best Practices
1. **Always ask permission** before using on work computers
2. **Document authorization** if required by your organization
3. **Use responsibly** and only for legitimate purposes
4. **Protect extracted data** according to your organization's data handling policies
5. **Delete extracted data** when no longer needed

## Limitations

- **Session-Based**: Data is lost when you refresh the page or close the tab
- **Outlook Web App Only**: Does not work with desktop Outlook application
- **DOM Dependent**: May need updates if Outlook changes their web interface
- **Company Policies**: Check your organization's policy regarding browser automation tools

## Customization

### Adapting to Different Outlook Versions

If the script doesn't work with your Outlook version, you can modify the CSS selectors:

```javascript
// Example: Custom subject selectors
const subjectSelectors = [
    'your-custom-selector',
    '.allowTextSelection:first-of-type',
    // ... other selectors
];
```

### Adding Custom Data Fields

You can extend the email object to capture additional fields:

```javascript
// In the extractCurrentEmail function
email.customField = document.querySelector('your-selector')?.textContent;
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Development Setup

1. Fork the repository
2. Create a feature branch
3. Test your changes with different Outlook configurations
4. Submit a pull request

## License

This project is licensed under the MIT License - see the file for details.

## Disclaimer

**⚠️ AUTHORIZATION REQUIRED**: This tool should only be used with proper authorization. Users must:

- **Corporate Accounts**: Obtain explicit written permission from IT department/management before use
- **Personal Accounts**: Tool can be used freely on your own personal Outlook account
- **Legal Compliance**: Ensure all usage complies with local laws and Microsoft's Terms of Service
- **Data Responsibility**: Only extract emails you have legal rights to access

This tool is for educational and legitimate business purposes only. Users are responsible for:
- Complying with their organization's IT policies and getting proper authorization
- Ensuring data privacy and security of extracted information
- Using the tool ethically and legally within their rights
- Protecting confidential information contained in extracted emails

**The authors are not responsible for any unauthorized use, policy violations, or consequences arising from improper use of this tool.**

By using this script, you acknowledge that you have:
1. ✅ Obtained necessary permissions (for corporate use)
2. ✅ Verified compliance with applicable policies and laws
3. ✅ Accepted full responsibility for your usage

## Version History

- **v1.0.0** - Initial release with basic email extraction
- **v1.1.0** - Added CSV export and duplicate prevention
- **v1.2.0** - Enhanced folder support and debug mode
- **v1.3.0** - Improved content cleaning and error handling

## Support

If you encounter issues:

1. Check the troubleshooting section
2. Enable debug mode to gather more information
3. Open an issue on GitHub with:
   - Browser version
   - Outlook Web App URL pattern
   - Console error messages
   - Debug output (without sensitive data)

---

**Note**: This script interacts with the Outlook Web App DOM structure. Microsoft may change their interface, which could require updates to the selectors used in this script.
