# events-cal-praises-planner

## Summary

A comprehensive SharePoint Framework (SPFx) webpart that creates a unified dashboard displaying:
- **Events & Calendar**: View and navigate your Microsoft 365 calendar events
- **Planner & To Do Tasks**: Track tasks from Microsoft Planner (team tasks) and Microsoft To Do (private tasks)
- **Viva Insights Praises**: Display recognition and praises received through Viva Insights

Built with TypeScript, React patterns, and Microsoft Graph API integration. The webpart features a responsive three-column layout with real-time data from Microsoft 365 services.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

### Required Licenses
- **Microsoft 365 E3/E5** or Business Premium subscription
- **Viva Insights** license (for praises feature)
- SharePoint Online

### Required Permissions
After deploying the solution, a SharePoint or Global Administrator must approve the following Microsoft Graph API permissions in **SharePoint Admin Center → API access**:

- `Calendars.Read` - Read user calendars
- `Tasks.Read` - Read Planner and To Do tasks
- `Tasks.ReadWrite` - Create and update Planner and To Do tasks
- `Group.Read.All` - Read groups (for Planner)
- `Mail.Read` - **Required for Viva Insights praises**
- `Community.Read.All` - Optional for Viva Engage praises
- `User.Read.All` - Read user profile information

### Setting up Viva Insights Praises

The praises feature displays recognition you've received through **Viva Insights → Personal Insights → Send praise to teammates**.

**To enable praises in this webpart:**

1. **Deploy the solution** and add the webpart to your page
2. **Approve API permissions** (SharePoint Admin Center → API access → Approve pending requests)
3. **Send/Receive praises** through Viva Insights:
   - Open Microsoft Teams or Outlook
   - Go to Viva Insights app
   - Navigate to "Personal Insights" or use the "Send praise" feature
   - Send praise to teammates

4. **Praise emails are detected** based on:
   - Subject containing "praise", "recognition", or "badge"
   - Emails from Viva Insights service
   - Keywords in email body

**Note:** Praises sent through Viva Insights are delivered as email notifications, which the webpart reads using the `Mail.Read` permission. If no praises appear, verify:
- You have received praise emails from Viva Insights
- The `Mail.Read` permission is approved
- Check browser console for any error messages

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

This webpart provides a comprehensive dashboard combining three key productivity features:

### 1. Events and Calendar
- Interactive calendar view with navigation
- Displays events from user's Microsoft 365 calendar
- Visual indicators for days with events
- Event list with date ribbons showing upcoming events
- Click on dates to view event details

### 2. Planner & To Do Tasks
- Displays tasks from Microsoft Planner (team tasks) and Microsoft To Do (private tasks)
- **Create Private Tasks**: Select "📋 My Tasks (Private)" to create personal To Do tasks
- **Create Team Tasks**: Select any Planner plan to create shared team tasks
- Search functionality to filter tasks
- Shows task priority, due dates, and completion status
- Progress tracking with slider (0-100% in 25% increments)
- Click on tasks to view and update progress in modal
- Unified task list showing both private and team tasks

### 3. Viva Insights Praises
- Automatically fetches praises received through Viva Insights
- Displays praise badges with custom colors (Thank You, Great Job, Kudos, etc.)
- Shows who sent the praise and the message
- Attempts multiple sources: Viva Insights emails, Viva Engage, Teams messages
- Visual praise cards with badge types and colors

This extension illustrates the following concepts:

- Microsoft Graph API integration (Calendar, Tasks, Mail, Communities)
- SPFx webpart development with TypeScript
- Permission management and API access approval
- Multi-source data aggregation
- Responsive dashboard layout with SCSS modules
- Interactive UI with modals and tooltips

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
