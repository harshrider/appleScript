-- =====================================================
-- Email Marketing Automation Script
-- =====================================================
-- This script automates sending personalized marketing emails through Apple Mail by reading contact data and content from a Numbers spreadsheet
-- 
-- REQUIREMENTS:
-- 1. Numbers app installed on macOS
-- 2. Apple Mail app configured with your email account
-- 3. Numbers file with specific table structure (see setup instructions)
-- 
-- SETUP INSTRUCTIONS:
-- 1. Update the file path below to point to your Numbers file
-- 2. Ensure your Numbers file has the correct table structure:
--    - Table named "Addresses" with columns A (email addresses) and B (names)
--    - Table named "Content" with cell A2 (subject) and B2 (email body)
-- 3. Test with a small dataset first
-- 4. Ensure Apple Mail is properly configured with your sending account
-- =====================================================


-- CONFIGURATION SECTION
-- Replace this path with the location of your Numbers file
-- Example formats:
-- Desktop: "/Users/[username]/Desktop/filename.numbers"
-- iCloud: "/Users/[username]/Library/Mobile Documents/com~apple~CloudDocs/Desktop/filename.numbers"

--set numbersFilePath to "/path/to/your/email-spreadsheet.numbers"

-- Convert the file path to an alias for AppleScript
set numbersFile to POSIX file numbersFilePath as alias

-- =====================================================
-- DATA EXTRACTION FROM NUMBERS SPREADSHEET
-- =====================================================
tell application "Numbers"
    -- Open the specified Numbers file
    open numbersFile
    
    tell front document -- Reference the opened document (now frontmost)
        tell active sheet
            -- Extract email addresses from column A (starting from row 2, skipping header)
            -- Assumes table named "Addresses" exists
            set theAddresses to value of cells 2 thru -1 of column "A" of table "Addresses"
            
            -- Extract contact names from column B (starting from row 2, skipping header)
            -- These names can be used for personalization
            set names to value of cells 2 thru -1 of column "B" of table "Addresses"
            
            -- Get the email subject line from the "Content" table, cell A2
            -- This allows you to update your subject line in the spreadsheet
            set theSubject to value of cell "A2" of table "Content"
            
            -- Get the email body content from the "Content" table, cell B2
            -- This is your main email template
            set theContent to value of cell "B2" of table "Content"
        end tell
    end tell
end tell

-- =====================================================
-- MAIN EMAIL SENDING LOOP
-- =====================================================
tell application "Mail"
    -- Process each email address in the spreadsheet
    repeat with i from 1 to count of theAddresses
        set anAddress to item i of theAddresses
        
        -- PERSONALIZATION OPTIONS:
        -- Option 1: Use names for personalization (currently commented out)
        -- Uncomment the line below to get the recipient's name
        -- set recipientName to item i of names
        
        -- Option 2: Basic content (current implementation)
        set personalizedContent to theContent
        
        -- Option 3: Personalized content with name (recommended)
        -- Uncomment and modify the lines below for personalization:
        -- set recipientName to item i of names
        -- set personalizedContent to "Dear " & recipientName & "," & return & return & theContent & return & return & "Best regards," & return & "[Your Name]"
        
        -- =====================================================
        -- EMAIL CREATION AND SENDING
        -- =====================================================
        
        -- Create a new outgoing message with subject and content
        set newMessage to make new outgoing message with properties {subject:theSubject, content:personalizedContent}
        
        -- Small delay to ensure message is created properly
        delay 1
        
        tell newMessage
            -- Add the recipient to the "To" field
            make new to recipient at end of every to recipient with properties {address:anAddress}
            
            -- SENDING OPTIONS:
            -- Option 1: Send immediately (current implementation)
            send
            
            -- Option 2: Save as draft for review (uncomment to use instead)
            -- save
            -- This will create drafts that you can review and send manually
        end tell
        
        -- Delay between emails to avoid overwhelming the mail server
        -- Adjust this value based on your email provider's limits
        delay 1 -- 1 second delay (increase for slower sending)
    end repeat
end tell

-- =====================================================
-- CUSTOMIZATION GUIDE
-- =====================================================

-- TO CUSTOMIZE THE EMAIL TEMPLATE:
-- 1. Open your Numbers file
-- 2. Navigate to the "Content" table
-- 3. Edit cell A2 for the subject line
-- 4. Edit cell B2 for the email body
-- 5. Use placeholders like [Name] if you want to add personalization later

-- TO ADD PERSONALIZATION:
-- 1. Uncomment the personalization lines above (around line 62-64)
-- 2. Modify the personalizedContent format to include:
--    - Greeting: "Dear " & recipientName & ","
--    - Your message content
--    - Sign-off: "Best regards," & return & "[Your Name]"

-- TO ADD HTML FORMATTING:
-- Replace the content line with HTML:
-- set personalizedContent to "<html><body><p>Dear " & recipientName & ",</p><p>" & theContent & "</p></body></html>"

-- TO ADD ATTACHMENTS:
-- Add this line inside the newMessage tell block:
-- make new attachment with properties {file name:"/path/to/your/attachment.pdf"}

-- TO CHANGE EMAIL ACCOUNT:
-- Add account specification to the message properties:
-- set newMessage to make new outgoing message with properties {subject:theSubject, content:personalizedContent, sender:"your-email@domain.com"}

-- TO ADD CC/BCC RECIPIENTS:
-- Add these lines inside the newMessage tell block:
-- make new cc recipient with properties {address:"cc-email@domain.com"}
-- make new bcc recipient with properties {address:"bcc-email@domain.com"}

-- =====================================================
-- SPREADSHEET STRUCTURE REQUIRED
-- =====================================================

-- Addresses Table:
-- | Column A (Email)        | Column B (Name)    |
-- |------------------------|--------------------|
-- | john@example.com       | John Smith         |
-- | jane@example.com       | Jane Doe           |

-- Content Table:
-- | Cell | Content                           |
-- |------|-----------------------------------|
-- | A2   | Your Email Subject Line           |
-- | B2   | Your email body content here...   |
