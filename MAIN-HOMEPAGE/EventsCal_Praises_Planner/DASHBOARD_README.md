# Enterprise Intranet Dashboard Web Component

## Overview
A modern, three-column enterprise intranet dashboard for SharePoint that displays Events & Calendar, Planner Tasks, and Praises in a clean, corporate layout.

## Features Implemented

### 🎨 Design Specifications
- **Width**: 97vw (bleeds edge-to-edge, adapts to SharePoint)
- **Font**: Arial, 8-10px sizes
- **Colors**:
  - Primary: Pantone Green (#00A651)
  - Secondary: Orange (#FF6B35), Black (#1C1C1C), Brown (#5C4033)
  - Background: Soft neutral grey (#F5F5F5)
  - Cards: White with 10px border radius
- **Effects**: Subtle shadows, smooth hover elevation
- **Responsive**:
  - Desktop: 3 columns
  - Tablet: 2 columns
  - Mobile: Stacked single column

---

## Column 1: Events & Calendar

### Calendar Component
- **Monthly view** with month/year navigation
- **Day indicators**:
  - Current day highlighted in Pantone Green
  - Days with events show green dot indicator
  - Hover reveals event details tooltip
- **Navigation**: Previous/next month arrows

### Event List (Ribbon Style)
- **Top 5 upcoming events** displayed
- **Each card shows**:
  - Circular date badge (month + day)
  - Event name and time
  - Short description preview
  - Arrow ribbon design edge
- **Data Source**: Microsoft Graph API - Outlook Calendar
- **Auto-refresh**: Updates when navigating calendar months

---

## Column 2: Planner (To-Do)

### Task Display
- **Top 5 tasks** sorted by due date (ascending) or priority
- **Each card shows**:
  - Task name (with ellipsis for long text)
  - Plan name
  - Due date
  - Priority (highlighted orange if High)
  - Progress status (Not Started/In Progress/Completed)
  - Paperclip icon if attachments exist

### Search Functionality
- **Mini search bar** at top of section
- Filters tasks by:
  - Task name
  - Plan name
- Real-time filtering

### Task Details Modal
- **Click any task** to open detailed view
- Shows:
  - Full description
  - Checklist items
  - Attachments indicator
  - Assigned users count
  - Created date
  - Bucket information

### Add Task Feature
- **"Add Task" button** at bottom
- Opens modal form with fields:
  - Task Name (required)
  - Plan selection (required)
  - Due Date
  - Priority (Normal/High/Urgent)
  - Description
- **Auto-refresh** after task creation

### Data Source
- Microsoft Graph API - Microsoft Planner
- Requires Planner API permissions

---

## Column 3: Praises

### Praise Cards (Ribbon Style)
- **Top 5 praises** dedicated to logged-in user
- **Each card shows**:
  - Sender profile photo (circular avatar with initials)
  - Praise title/badge type
  - Sender name
  - Praise message
  - Colored background (20-30% opacity)
- **Background colors**: Dynamic based on praise type (Blue, Orange, Teal, Red)

### Data Source
- Microsoft Graph API - Viva Insights
- Note: May require Viva Insights license and API configuration

---

## Technical Implementation

### Technologies Used
- **SharePoint Framework (SPFx)** 1.20.0
- **TypeScript** 4.7.4
- **SCSS** with modular styling
- **Microsoft Graph API** (MSGraphClientV3)
- **Fluent UI React** 8.106.4

### API Integrations

#### Events (Outlook Calendar)
```typescript
GET /me/events
Filter: Current month events
Select: id, subject, start, end, location, bodyPreview
```

#### Planner Tasks
```typescript
GET /me/planner/tasks
GET /planner/plans/{planId}
GET /planner/tasks/{taskId}/details
POST /planner/tasks (for creating tasks)
```

#### Praises (Viva Insights)
```typescript
GET /me/insights/used
Note: This is a placeholder - actual Viva API may differ
```

### Key Components

#### TypeScript Classes & Interfaces
- `IEvent` - Calendar event structure
- `IPlannerTask` - Planner task structure
- `IPraise` - Praise/recognition structure
- `ICalendarDay` - Calendar day with event data

#### Main Methods
- `_fetchEvents()` - Retrieves Outlook calendar events
- `_fetchPlannerTasks()` - Retrieves Planner tasks and plans
- `_fetchPraises()` - Retrieves Viva Insights praises
- `_generateCalendar()` - Builds calendar grid with 42 days
- `_renderCalendar()` - Renders calendar HTML with event indicators
- `_renderEvents()` - Renders event cards list
- `_renderTasks()` - Renders task cards with search filtering
- `_renderPraises()` - Renders praise cards
- `_showTaskDetails()` - Opens modal with full task information
- `_createTask()` - Creates new Planner task

### Loading States
- **Skeleton loaders** for all three columns during data fetch
- Smooth transitions from loading to content
- Animated shimmer effect

### Error Handling
- Try/catch blocks for all API calls
- Graceful fallback to empty states
- User-friendly error messages
- Console logging for debugging

---

## Permissions Required

Add these permissions to `package-solution.json`:

```json
"webApiPermissionRequests": [
  {
    "resource": "Microsoft Graph",
    "scope": "Calendars.Read"
  },
  {
    "resource": "Microsoft Graph",
    "scope": "Tasks.ReadWrite"
  },
  {
    "resource": "Microsoft Graph",
    "scope": "Group.ReadWrite.All"
  },
  {
    "resource": "Microsoft Graph",
    "scope": "User.Read"
  }
]
```

**Note**: Admin consent required for these permissions in SharePoint Admin Center.

---

## Installation & Deployment

### Build the Solution
```bash
cd MAIN-HOMEPAGE/EventsCal_Praises_Planner
npm install
npx gulp build
npx gulp bundle --ship
npx gulp package-solution --ship
```

### Deploy to SharePoint
1. Upload `.sppkg` file from `sharepoint/solution/` to App Catalog
2. Grant API permissions in SharePoint Admin Center:
   - Go to API Management
   - Approve pending permission requests
3. Add web part to SharePoint page

### Local Development
```bash
npx gulp serve
```
Then add the web part to your SharePoint workbench.

---

## SCSS Architecture

### Modular Structure
- Variables for colors, spacing, typography
- BEM-style naming convention
- Responsive breakpoints:
  - Desktop: > 1024px (3 columns)
  - Tablet: 768px-1024px (2 columns)
  - Mobile: < 768px (1 column)

### Key Style Classes
- `.dashboardContainer` - Main grid container
- `.column` - Individual column wrapper
- `.calendarContainer` - Calendar component
- `.eventCard` - Event card with ribbon
- `.taskCard` - Task card with left border
- `.praiseCard` - Praise card with avatar
- `.modal` - Modal overlay for task details/creation

---

## Accessibility Features

- **ARIA labels** on interactive elements
- **Keyboard navigation** support
- **Focus indicators** with green outline
- **Semantic HTML** structure
- **Screen reader friendly** content
- **Reduced motion** support for animations

---

## Browser Compatibility

- ✅ Microsoft Edge (Chromium)
- ✅ Google Chrome
- ✅ Mozilla Firefox
- ✅ Safari (macOS/iOS)
- ⚠️ Internet Explorer 11 (limited support)

---

## Known Limitations

1. **Viva Insights API**: The praises feature uses a placeholder implementation. Actual Viva Insights API access may require:
   - Viva Insights license
   - Different API endpoint
   - Additional permissions

2. **Task Creation**: Requires at least one bucket to exist in the selected plan

3. **Calendar Events**: Only shows events from the user's primary calendar

4. **Performance**: Large datasets (>100 tasks/events) may impact load time

---

## Customization Options

### Colors
Modify SCSS variables in `EventsCalPraisesPlannerWebPart.module.scss`:
```scss
$primary-green: #00A651;
$secondary-orange: #FF6B35;
$bg-neutral: #F5F5F5;
```

### Card Limits
Adjust in TypeScript:
```typescript
.slice(0, 5) // Change 5 to desired number
```

### Column Layout
Modify grid breakpoints in SCSS:
```scss
@media (max-width: 1024px) {
  grid-template-columns: repeat(2, 1fr);
}
```

---

## Troubleshooting

### "Failed to load dashboard data"
- Check API permissions are granted
- Verify user has access to Planner/Calendar
- Check browser console for specific errors

### Tasks Not Appearing
- Ensure user is assigned to at least one Planner plan
- Verify Planner API permissions are approved

### Modal Not Opening
- Check browser console for JavaScript errors
- Ensure modal event listeners attached correctly

### Styling Issues
- Run `npx gulp build` to recompile SCSS
- Clear browser cache
- Check for CSS conflicts with SharePoint theme

---

## Future Enhancements

- [ ] Drag-and-drop task re-ordering
- [ ] Event creation directly from calendar
- [ ] Task completion toggling from card
- [ ] Advanced filtering options
- [ ] Export to Excel/PDF
- [ ] Dark mode support
- [ ] Multi-language support
- [ ] Notifications for upcoming tasks/events
- [ ] Team calendar view option
- [ ] Integration with Microsoft To Do

---

## Support & Maintenance

**Built with**: SharePoint Framework 1.20.0  
**Node Version**: 18.20.8  
**TypeScript**: 4.7.4  
**Last Updated**: February 2026

For issues or questions, refer to:
- [SharePoint Framework Documentation](https://aka.ms/spfx)
- [Microsoft Graph API Reference](https://docs.microsoft.com/graph/api/overview)
- [Planner API Documentation](https://docs.microsoft.com/graph/api/resources/planner-overview)

---

## License

This web component is built for enterprise use within your organization's SharePoint environment.
