# Teams Message Fetcher - Streamlit App

A web application built with Streamlit to fetch, analyze, and export Microsoft Teams channel messages with optional Jira integration.

## Features

- 🔐 **Microsoft Graph API Integration**: Authenticate and fetch messages from Teams channels
- 📊 **Data Analysis**: View summaries, charts, and statistics of your Teams messages
- 🔍 **Advanced Filtering**: Filter by date, sender, message type, and more
- 📥 **Export Options**: Download data as Excel or CSV files
- 🎫 **Jira Integration**: Automatically create Jira tasks from Teams messages
- 📈 **Interactive Visualizations**: Charts showing message distribution and trends
- ⏱️ **Task Duration Tracking**: Calculate estimated hours based on "done" replies

## Installation

1. Clone or download this project
2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the Streamlit app:
   ```bash
   streamlit run streamlit_app.py
   ```

2. Open your browser and go to `http://localhost:8501`

3. Configure the following in the sidebar:

### Authentication
- **Tenant ID**: Your Azure AD tenant ID
- **Client ID**: Your app registration client ID
- **Client Secret**: Your app registration client secret

### Teams Configuration
- **Team ID**: The ID of your Microsoft Teams team
- **Channel ID**: The ID of the specific channel to fetch messages from

### Filters
- **Filter from date**: Only fetch messages from this date onwards
- **Filter by sender**: Optional filter to show messages from specific users

### Jira Integration (Optional)
- **Enable Jira Integration**: Check to enable Jira features
- **Jira URL**: Your Jira instance URL
- **Jira Email**: Your Jira account email
- **Jira API Token**: Your Jira API token
- **Project Key**: The Jira project key where tasks will be created
- **Epic Key**: The epic under which tasks will be created

## How to Get Required IDs

### Azure App Registration
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to "Azure Active Directory" > "App registrations"
3. Create a new app registration or use existing one
4. Note down the **Application (client) ID** and **Directory (tenant) ID**
5. Create a client secret and note it down
6. Add Microsoft Graph API permissions: `ChannelMessage.Read.All`, `Team.ReadBasic.All`

### Teams IDs
1. **Team ID**: In Teams web app, the URL contains the team ID
2. **Channel ID**: You can get this from Teams web app URL or using Graph API

### Jira API Token
1. Go to [Atlassian Account Settings](https://id.atlassian.com/manage-profile/security/api-tokens)
2. Create a new API token
3. Note down the token (keep it secure)

## Features Overview

### Data Analysis
- **Summary Statistics**: Total messages, average hours, unique senders
- **Message Type Classification**: Automatic categorization (Request, Bug, Issue)
- **Time-based Analysis**: Charts showing message trends over time
- **Duration Tracking**: Calculates time between message and "done" reply

### Export Options
- **Excel Export**: Multi-sheet Excel file with messages and summary
- **CSV Export**: Simple CSV format for data analysis
- **Jira Integration**: Create tasks directly in Jira with proper formatting

### Interactive Features
- **Real-time Filtering**: Filter messages by type, sender, date range
- **Progress Tracking**: Visual progress bars during data fetching
- **Error Handling**: Comprehensive error messages and recovery options

## Security Notes

- API tokens and secrets are entered as password fields (masked in UI)
- No credentials are stored permanently
- All API calls use secure HTTPS connections
- Session state manages temporary data only

## Troubleshooting

### Authentication Issues
- Verify your Azure app registration has correct permissions
- Ensure client secret hasn't expired
- Check that the app has been granted admin consent

### Teams Access Issues
- Verify you have access to the specified team and channel
- Ensure the app registration has been added to the team (if required)
- Check that the team and channel IDs are correct

### Jira Integration Issues
- Verify your Jira API token is valid
- Ensure you have permissions to create tasks in the specified project
- Check that the epic key exists and you have access to it

## Technical Details

- Built with Streamlit for the web interface
- Uses Microsoft Graph API for Teams data
- Jira REST API for task creation
- Pandas for data manipulation and analysis
- BeautifulSoup for HTML content cleaning
- Supports timezone-aware datetime handling

## Requirements

- Python 3.7+
- Active Microsoft 365 subscription with Teams
- Azure app registration with appropriate permissions
- Optional: Jira instance with API access

## License

This project is for internal use. Ensure compliance with your organization's data handling policies.