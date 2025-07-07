# MailLib Testing Guide

This guide explains how to test the MailLib component using the provided VBScript test files.

## Files Included

- `TestMailLib.vbs` - Main test script for MailLib
- `SetupEnvironment.bat` - Batch file to set up environment variables
- `TESTING.md` - This guide

## Testing Options

### Option 1: Private Testing (Hardcoded Credentials)

1. **Edit `TestMailLib.vbs`** and fill in your SMTP credentials:
   ```vbscript
   SMTP_HOST = "smtp.gmail.com"
   SMTP_PORT = 587
   SMTP_USERNAME = "your-email@gmail.com"
   SMTP_PASSWORD = "your-app-password"
   FROM_EMAIL = "sender@yourdomain.com"
   TO_EMAIL = "recipient@example.com"
   ```

2. **Run the test**:
   ```cmd
   cscript TestMailLib.vbs
   ```

### Option 2: Public Testing (Environment Variables)

1. **Edit `SetupEnvironment.bat`** with your credentials:
   ```batch
   set MAILLIB_SMTP_HOST=smtp.gmail.com
   set MAILLIB_SMTP_PORT=587
   set MAILLIB_SMTP_USERNAME=your-email@gmail.com
   set MAILLIB_SMTP_PASSWORD=your-app-password
   set MAILLIB_FROM_EMAIL=sender@yourdomain.com
   set MAILLIB_TO_EMAIL=recipient@example.com
   ```

2. **Run the setup script**:
   ```cmd
   SetupEnvironment.bat
   ```

3. **Edit `TestMailLib.vbs`** and uncomment the environment variable lines:
   ```vbscript
   ' Uncomment these lines to use environment variables
   SMTP_HOST = GetEnvironmentVariable("MAILLIB_SMTP_HOST")
   SMTP_PORT = CInt(GetEnvironmentVariable("MAILLIB_SMTP_PORT"))
   SMTP_USERNAME = GetEnvironmentVariable("MAILLIB_SMTP_USERNAME")
   SMTP_PASSWORD = GetEnvironmentVariable("MAILLIB_SMTP_PASSWORD")
   FROM_EMAIL = GetEnvironmentVariable("MAILLIB_FROM_EMAIL")
   TO_EMAIL = GetEnvironmentVariable("MAILLIB_TO_EMAIL")
   ```

4. **Run the test**:
   ```cmd
   cscript TestMailLib.vbs
   ```

## Prerequisites

1. **Register MailLib.dll**:
   ```cmd
   RegAsm.exe MailLib.dll /codebase /tlb
   ```

2. **Build the project** in Visual Studio or using MSBuild

3. **Ensure .NET Framework 4.8** is installed

## Common SMTP Settings

### Gmail
- Host: `smtp.gmail.com`
- Port: `587` (STARTTLS) or `465` (SSL)
- Username: Your Gmail address
- Password: App Password (not your regular password)

### Outlook/Hotmail
- Host: `smtp-mail.outlook.com`
- Port: `587`
- Username: Your Outlook email address
- Password: Your account password

### Yahoo
- Host: `smtp.mail.yahoo.com`
- Port: `587`
- Username: Your Yahoo email address
- Password: App Password

## What the Test Does

The test script performs the following:

1. **Validates configuration** - Checks that all required settings are provided
2. **Creates EmailSender object** - Tests COM interop
3. **Configures SMTP settings** - Sets host, port, credentials, and security
4. **Sends test email** - Creates both HTML and text versions
5. **Tests CDO configuration** - Demonstrates CDO-style configuration approach
6. **Reports results** - Shows success/failure and any errors

## Error Troubleshooting

### "Failed to create MailLib.EmailSender object"
- Ensure MailLib.dll is registered with RegAsm
- Check that .NET Framework 4.8 is installed
- Verify the DLL is in the correct location

### "Authentication failed"
- Check username and password
- For Gmail, use App Password instead of regular password
- Ensure 2-factor authentication is enabled for App Passwords

### "Connection failed"
- Check SMTP host and port
- Verify firewall settings
- Test with a different SMTP server

## Security Notes

- **Never commit credentials** to version control
- **Use environment variables** for public repositories
- **Use App Passwords** instead of regular passwords for Gmail
- **Keep test files private** when containing real credentials

## Environment Variables for CI/CD

For automated testing, set these environment variables:

```bash
MAILLIB_SMTP_HOST=smtp.gmail.com
MAILLIB_SMTP_PORT=587
MAILLIB_SMTP_USERNAME=test@example.com
MAILLIB_SMTP_PASSWORD=app-password
MAILLIB_FROM_EMAIL=test@example.com
MAILLIB_TO_EMAIL=test@example.com
```

This allows the test script to run in CI/CD environments without exposing credentials in the code. 