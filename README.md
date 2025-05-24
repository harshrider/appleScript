# AppleScript Marketing Automation

Automated marketing tools for macOS that send personalized messages through iMessage and email using AppleScript integration with Numbers spreadsheets.

## Features

- **iMessage Automation**: Send bulk personalized iMessages
- **Email Automation**: Automated email campaigns via Apple Mail  
- **Spreadsheet Integration**: Read contacts and templates from Numbers
- **Personalization**: Dynamic message customization with recipient names
- **Batch Processing**: Handle multiple contacts efficiently
- **Safety Mode**: Messages typed but not auto-sent for review

## Prerequisites

- macOS 10.14+
- Numbers app
- Messages app (for iMessage script)
- Apple Mail (for email script)
- Configured iMessage/email accounts

## Setup

### 1. System Permissions

Go to **System Preferences → Security & Privacy → Privacy**:

**Accessibility:**
- Add and enable Script Editor

**Automation:**
- Allow Script Editor to control Numbers, Messages, Mail, and System Events

**Full Disk Access** (if using iCloud files):
- Add Script Editor

**Files and Folders:**
- Grant access to your spreadsheet location (Desktop, Documents, or iCloud Drive)

### 2. Spreadsheet Structure

Create a Numbers file with two tables:

**Addresses Table:**
| Column A | Column B |
|----------|----------|
| Phone/Email | Name |
| +1234567890 | John Smith |
| user@email.com | Jane Doe |

**Content Table:**
| Cell | Content |
|------|---------|
| A2 | Subject line or message template |
| B2 | Email body content (email script only) |

### 3. Script Configuration

1. Download the script you need
2. Open in Script Editor
3. Update the file path:
   ```applescript
   set numbersFilePath to "/path/to/your/spreadsheet.numbers"
   ```
4. Customize the message template in the script

## Usage

### iMessage Script
```applescript
-- Reads from Addresses table (phone numbers) and Content table
-- Sends personalized iMessages through Messages app
-- Messages are typed but not automatically sent
```

### Email Script  
```applescript
-- Reads from Addresses table (email addresses) and Content table
-- Sends emails through Apple Mail
-- Supports immediate sending or draft creation
```

## Customization

### Basic Personalization
```applescript
set personalizedContent to "Dear " & contactName & ", " & theContent
```

### HTML Email (Email Script)
```applescript
set personalizedContent to "<html><body><p>Hello " & recipientName & "</p></body></html>"
```

### Add Delays
```applescript
delay 2 -- Increase delay between messages
```

### Draft Mode (Email Script)
```applescript
-- Replace 'send' with 'save' to create drafts instead
save
```

## File Structure

```
├── imessage-automation.applescript    # iMessage marketing script
├── email-automation.applescript      # Email marketing script  
└── README.md                         # This file
```

## How It Works

1. **Data Reading**: Scripts parse your Numbers spreadsheet for contacts and message templates
2. **Message Creation**: Personalizes content with recipient names  
3. **Automation**: Uses System Events to control Messages/Mail apps
4. **Safety**: iMessage script types messages for manual review; email script can auto-send or save drafts

## Troubleshooting

**Permission denied errors:**
- Check System Preferences privacy settings
- Restart applications after granting permissions

**Can't access iCloud files:**
- Enable Full Disk Access for Script Editor
- Verify iCloud Drive sync status

**Messages not sending:**
- Verify app configurations (Messages/Mail)
- Check internet connectivity
- Ensure valid phone numbers/email formats

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature-name`)
3. Commit changes (`git commit -am 'Add feature'`)
4. Push to branch (`git push origin feature-name`)
5. Create a Pull Request

## License

MIT License - see LICENSE file for details

## Disclaimer

Use responsibly and ensure you have permission to contact recipients. Users are responsible for compliance with applicable messaging laws and regulations.
