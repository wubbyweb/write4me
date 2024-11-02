# Outlook Email Affirmation Assistant

A VBA-based Microsoft Outlook add-in that helps generate AI-powered affirmative responses to emails using OpenAI's GPT API.

## Setup Instructions

1. **Config File Setup**
   - Create a `config.ini` file in the same directory as your Outlook files
   - Add the following content:
     ```ini
     OPENAI_API_KEY=your_openai_api_key_here
     API_ENDPOINT=https://api.openai.com/v1/chat/completions
     ```

2. **VBA Module Import**
   - Import `ConfigManager.vba` and `MainModule.vba` into your Outlook VBA project
   - Enable VBA in Outlook if not already enabled

## Features

- Customizable email response generation
- Multiple tone options (formal, casual, humorous)
- Adjustable response length (short, long)
- Professional HTML formatting of responses
- Error handling and validation

## Usage
### Setting up the Quick Access Button
1. Right-click on the Outlook ribbon
2. Select "Customize the Ribbon"
3. In the right column, select "New Mail Message"
4. Click "New Group" at the bottom
5. Rename the group (e.g., "AI Assistant")
6. In the left column, choose "Macros"
7. Find "GenerateAffirmation" and click "Add"
8. Click "OK" to save changes
9. 
### Using the Tool
1. Create a new email or reply to an existing one
2. Click the GenerateAffirmation button in your custom ribbon group
3. Select your preferred tone and length in the dialog
4. The AI-generated response will be inserted at the top of your email

## Dependencies

- Microsoft Outlook
- OpenAI API access
- Active internet connection

## Error Handling

The system includes comprehensive error handling for:
- Missing configuration
- API communication issues
- Invalid responses
- File system access

## Security Note

Store your OpenAI API key securely and never share it in your code.