import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import styles from './EventsCalPraisesPlannerWebPart.module.scss';
import * as strings from 'EventsCalPraisesPlannerWebPartStrings';

export interface IEventsCalPraisesPlannerWebPartProps {
  description: string;
}

interface IEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location: { displayName: string };
  bodyPreview: string;
}

interface IPlannerTask {
  id: string;
  title: string;
  dueDateTime: string | null;
  priority: number;
  percentComplete: number;
  hasDescription: boolean;
  assignments: any;
  planId: string;
  bucketId: string;
  createdDateTime: string;
}

interface IPraise {
  id: string;
  displayName: string;
  givenBy: {
    displayName: string;
    id: string;
    photoUrl?: string;
  };
  message: string;
  badgeType: string;
  backgroundColor: string;
}

interface ICalendarDay {
  date: Date;
  day: number;
  isCurrentMonth: boolean;
  isToday: boolean;
  hasEvents: boolean;
  events: IEvent[];
}

export default class EventsCalPraisesPlannerWebPart extends BaseClientSideWebPart<IEventsCalPraisesPlannerWebPartProps> {

  private graphClient: MSGraphClientV3;
  private events: IEvent[] = [];
  private tasks: IPlannerTask[] = [];
  private plans: Map<string, string> = new Map();
  private praises: IPraise[] = [];
  private currentMonth: Date = new Date();
  private calendarDays: ICalendarDay[] = [];
  private searchQuery: string = '';

  public async render(): Promise<void> {
    this.domElement.innerHTML = `
    <section class="${styles.eventsCalPraisesPlanner}">
      <div class="${styles.dashboardContainer}">
        <!-- Column 1: Events & Calendar -->
        <div class="${styles.column}">
          <h3 class="${styles.sectionTitle}">EVENTS AND CALENDAR</h3>
          <div class="${styles.calendarContainer}">
            <div class="${styles.calendarHeader}">
              <button class="${styles.navButton}" id="prevMonth" aria-label="Previous month">‹</button>
              <span class="${styles.monthDisplay}" id="monthDisplay">Loading...</span>
              <button class="${styles.navButton}" id="nextMonth" aria-label="Next month">›</button>
            </div>
            <div class="${styles.calendarGrid}" id="calendarGrid">
              ${this._renderLoadingSkeleton('calendar')}
            </div>
          </div>
          <div class="${styles.eventsList}" id="eventsList">
            ${this._renderLoadingSkeleton('events')}
          </div>
        </div>

        <!-- Column 2: Planner Tasks -->
        <div class="${styles.column}">
          <h3 class="${styles.sectionTitle}">PLANNER (TO-DO)</h3>
          <div class="${styles.searchContainer}">
            <input 
              type="text" 
              class="${styles.searchInput}" 
              id="taskSearch" 
              placeholder="Search tasks or plans..."
              aria-label="Search tasks"
            />
          </div>
          <div class="${styles.tasksList}" id="tasksList">
            ${this._renderLoadingSkeleton('tasks')}
          </div>
          <button class="${styles.addButton}" id="addTaskBtn">+ Add Task</button>
        </div>

        <!-- Column 3: Praises -->
        <div class="${styles.column}">
          <h3 class="${styles.sectionTitle}">PRAISES</h3>
          <div class="${styles.praisesList}" id="praisesList">
            ${this._renderLoadingSkeleton('praises')}
          </div>
        </div>
      </div>

      <!-- Task Modal -->
      <div class="${styles.modal}" id="taskModal" style="display: none;" role="dialog" aria-modal="true">
        <div class="${styles.modalContent}">
          <span class="${styles.closeModal}" id="closeModal" aria-label="Close">&times;</span>
          <div id="modalBody"></div>
        </div>
      </div>
    </section>`;

    this._attachEventListeners();
    await this._loadData();
  }

  protected async onInit(): Promise<void> {
    this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    return Promise.resolve();
  }

  private async _loadData(): Promise<void> {
    try {
      await Promise.all([
        this._fetchEvents(),
        this._fetchPlannerTasks(),
        this._fetchPraises()
      ]);
      
      this._generateCalendar();
      this._renderCalendar();
      this._renderEvents();
      this._renderTasks();
      this._renderPraises();
    } catch (error) {
      console.error('Error loading data:', error);
      this._showError('Failed to load dashboard data. Please refresh the page.');
    }
  }

  private async _fetchEvents(): Promise<void> {
    try {
      const startOfMonth = new Date(this.currentMonth.getFullYear(), this.currentMonth.getMonth(), 1);
      const endOfMonth = new Date(this.currentMonth.getFullYear(), this.currentMonth.getMonth() + 1, 0, 23, 59, 59);
      
      const response = await this.graphClient
        .api('/me/events')
        .filter(`start/dateTime ge '${startOfMonth.toISOString()}' and end/dateTime le '${endOfMonth.toISOString()}'`)
        .select('id,subject,start,end,location,bodyPreview')
        .orderby('start/dateTime')
        .top(50)
        .get();
      
      this.events = response.value || [];
    } catch (error) {
      console.error('Error fetching events:', error);
      this.events = [];
    }
  }

  private async _fetchPlannerTasks(): Promise<void> {
    try {
      console.log('Fetching planner tasks...');
      const tasksResponse = await this.graphClient
        .api('/me/planner/tasks')
        .select('id,title,dueDateTime,priority,percentComplete,hasDescription,assignments,planId,bucketId,createdDateTime')
        .top(100)
        .get();
      
      const allTasks: IPlannerTask[] = tasksResponse.value || [];
      console.log(`Fetched ${allTasks.length} total tasks from Planner`);
      
      // Fetch plan names
      const planIdsSet: { [key: string]: boolean } = {};
      allTasks.forEach(t => { planIdsSet[t.planId] = true; });
      const planIds = Object.keys(planIdsSet);
      for (const planId of planIds) {
        try {
          const plan = await this.graphClient.api(`/planner/plans/${planId}`).select('id,title').get();
          this.plans.set(plan.id, plan.title);
        } catch (error) {
          console.error(`Error fetching plan ${planId}:`, error);
        }
      }
      
      // Fetch Microsoft To Do tasks (private tasks)
      try {
        console.log('Fetching To Do tasks...');
        const listsResponse = await this.graphClient.api('/me/todo/lists').get();
        const taskLists = listsResponse.value || [];
        console.log(`Found ${taskLists.length} To Do lists`);
        // Add "My Tasks (Private)" to plans map
        this.plans.set('TODO_PRIVATE', '📋 My Tasks (Private)');
        for (const list of taskLists) {
          try {
            const todoTasksResponse = await this.graphClient
              .api(`/me/todo/lists/${list.id}/tasks`)
              .top(50)
              .get();
            const todoTasks = todoTasksResponse.value || [];
            console.log(`Fetched ${todoTasks.length} tasks from To Do list: ${list.displayName}`);
            for (const todoTask of todoTasks) {
              if (todoTask.status === 'completed') {
                console.log(`Skipping completed To Do task: ${todoTask.title}`);
                continue;
              }
              const convertedTask: IPlannerTask = {
                id: `TODO_${todoTask.id}`,
                title: todoTask.title || 'Untitled Task',
                dueDateTime: todoTask.dueDateTime?.dateTime || null,
                priority: this._convertToDoPriority(todoTask.importance),
                percentComplete: todoTask.status === 'completed' ? 100 : 0,
                hasDescription: !!todoTask.body?.content,
                assignments: {},
                planId: 'TODO_PRIVATE',
                bucketId: list.id,
                createdDateTime: todoTask.createdDateTime
              };
              console.log(`Added To Do task to list: ${todoTask.title} (ID: ${convertedTask.id})`);
              allTasks.push(convertedTask);
            }
          } catch (listError) {
            console.error(`Error fetching tasks from To Do list ${list.id}:`, listError);
          }
        }
        console.log(`Total tasks after adding To Do: ${allTasks.length}`);
      } catch (todoError) {
        console.error('Error fetching To Do tasks:', todoError);
        // Continue with just Planner tasks if To Do fails
      }
      // Sort by due date and priority, then by creation date (newest first)
      this.tasks = allTasks
        .sort((a, b) => {
          if (a.dueDateTime && b.dueDateTime) {
            return new Date(a.dueDateTime).getTime() - new Date(b.dueDateTime).getTime();
          }
          if (a.dueDateTime) return -1;
          if (b.dueDateTime) return 1;
          return new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime();
        });
      console.log(`Displaying ${this.tasks.length} tasks`);
    } catch (error) {
      console.error('Error fetching planner tasks:', error);
      this.tasks = [];
    }
  }
  private _convertToDoPriority(importance: string): number {
    // Convert To Do importance to Planner priority (0-10, lower is higher priority)
    switch (importance) {
      case 'high':
        return 1;
      case 'normal':
        return 5;
      case 'low':
        return 9;
      default:
        return 5;
    }
  }

  /**
   * Fetches praises from Viva Insights and other Microsoft 365 sources
   * 
   * Requirements:
   * 1. Viva Insights license for direct team praises
   * 2. Microsoft Graph API permissions:
   *    - Mail.Read (for Viva Insights praise emails)
   *    - User.Read.All (for user profile information)
   *    - Community.Read.All (for Viva Engage if enabled)
   * 
   * Note: Viva Insights praises are delivered via email notifications.
   */
  private async _fetchPraises(): Promise<void> {
    try {
      // Attempt 1: Fetch Viva Insights praises from emails (Primary source for "Send praise to teammates")
      try {
        console.info('Fetching Viva Insights praises from email...');
        
        // Search for emails from Viva Insights with praise notifications
        const praiseEmailsResponse = await this.graphClient
          .api('/me/messages')
          .select('id,subject,from,body,bodyPreview,receivedDateTime,toRecipients')
          .filter("(contains(subject, 'praise') or contains(subject, 'recognition') or contains(subject, 'badge') or contains(from/emailAddress/address, 'insights') or contains(from/emailAddress/address, 'viva'))")
          .orderby('receivedDateTime desc')
          .top(30)
          .get();

        const praiseEmails = praiseEmailsResponse.value || [];
        console.info(`Found ${praiseEmails.length} potential praise emails`);
        
        if (praiseEmails.length > 0) {
          const parsedPraises: IPraise[] = [];
          
          for (const email of praiseEmails) {
            // Parse the email to extract praise information
            const subject = email.subject || '';
            const body = email.bodyPreview || email.body?.content || '';
            const from = email.from?.emailAddress?.name || email.from?.emailAddress?.address || '';
            
            // Check if this is a Viva Insights praise notification
            const isVivaPraise = subject.toLowerCase().includes('received praise') ||
                               subject.toLowerCase().includes('sent you praise') ||
                               subject.toLowerCase().includes('recognized you') ||
                               from.toLowerCase().includes('viva') ||
                               from.toLowerCase().includes('insights') ||
                               body.toLowerCase().includes('viva insights');
            
            if (isVivaPraise || subject.toLowerCase().includes('praise') || subject.toLowerCase().includes('recognition')) {
              // Extract the sender name from subject or body
              let senderName = from;
              const subjectMatch = subject.match(/from\s+([^<\n]+)/i) || subject.match(/by\s+([^<\n]+)/i);
              if (subjectMatch && subjectMatch[1]) {
                senderName = subjectMatch[1].trim();
              }
              
              // Extract badge type from subject or body
              const badgeInfo = this._extractBadgeType(subject, body);
              
              parsedPraises.push({
                id: email.id,
                displayName: badgeInfo.badgeType,
                givenBy: {
                  displayName: senderName,
                  id: email.from?.emailAddress?.address || '',
                  photoUrl: this._getUserPhotoUrl(email.from?.emailAddress?.address)
                },
                message: this._extractPraiseMessage(body),
                badgeType: badgeInfo.badgeType,
                backgroundColor: badgeInfo.backgroundColor
              });
            }
          }
          
          if (parsedPraises.length > 0) {
            this.praises = parsedPraises.slice(0, 5);
            console.info(`Successfully loaded ${this.praises.length} Viva Insights praises`);
            return;
          }
        }
      } catch (praiseEmailError) {
        console.warn('Could not fetch Viva Insights praise emails:', praiseEmailError);
        
        // Check if this is a permission error
        if (praiseEmailError && praiseEmailError.message && praiseEmailError.message.includes('Access is denied')) {
          console.error('⚠️ PERMISSION REQUIRED: Mail.Read permission not approved.');
          console.error('📋 To fix: Go to SharePoint Admin Center → Advanced → API access → Approve "Mail.Read" permission');
        }
      }

      // Attempt 2: Try beta endpoint for insights (if available)
      try {
        // Note: This endpoint may require beta API version
        const insightsResponse = await this.graphClient
          .api('/me/insights/shared')
          .version('beta')
          .top(10)
          .get();
        
        console.info('Insights API available:', insightsResponse);
      } catch (insightsError) {
        console.warn('Insights API not accessible:', insightsError);
      }

      // Attempt 3: Fetch from Viva Engage communities (if praises are shared publicly)
      try {
        // First, get user's Viva Engage communities
        const communitiesResponse = await this.graphClient
          .api('/employeeExperience/communities')
          .select('id,displayName,description')
          .top(10)
          .get();

        if (communitiesResponse?.value?.length > 0) {
          console.info('Viva Engage communities found:', communitiesResponse.value.length);
          
          // Try to fetch posts from communities that might contain praises
          const allPraises: IPraise[] = [];
          
          for (const community of communitiesResponse.value.slice(0, 3)) {
            try {
              // Attempt to get posts from the community
              const postsResponse = await this.graphClient
                .api(`/employeeExperience/communities/${community.id}/posts`)
                .select('id,content,author,createdDateTime')
                .orderby('createdDateTime desc')
                .top(20)
                .get();

              if (postsResponse?.value) {
                // Filter posts that look like praises (contain praise keywords or badges)
                const praisePosts = postsResponse.value.filter((post: any) => {
                  const content = (post.content?.toLowerCase() || '');
                  return content.includes('praise') || content.includes('badge') || 
                         content.includes('kudos') || content.includes('recognition') ||
                         content.includes('thank') || content.includes('appreciate');
                });

                praisePosts.forEach((post: any) => {
                  const badgeInfo = this._extractBadgeType('', post.content);
                  allPraises.push({
                    id: post.id,
                    displayName: badgeInfo.badgeType,
                    givenBy: {
                      displayName: post.author?.displayName || 'Colleague',
                      id: post.author?.id || '',
                      photoUrl: this._getUserPhotoUrl(post.author?.mail)
                    },
                    message: (post.content || '').substring(0, 100),
                    badgeType: badgeInfo.badgeType,
                    backgroundColor: badgeInfo.backgroundColor
                  });
                });
              }
            } catch (postsError) {
              console.warn(`Could not fetch posts from community ${community.displayName}:`, postsError);
            }
          }

          if (allPraises.length > 0) {
            this.praises = allPraises.slice(0, 5);
            return;
          }
        }
      } catch (vivaError) {
        console.warn('Viva Engage API not available:', vivaError);
        
        // Check if this is a permission error
        if (vivaError && vivaError.message && vivaError.message.includes('Authorization')) {
          console.warn('⚠️ Community.Read.All permission not approved (optional for Viva Engage praises)');
        }
      }

      // Attempt 4: Fetch recent important messages that might contain praise/recognition
      try {
        const messagesResponse = await this.graphClient
          .api('/me/messages')
          .select('id,subject,from,bodyPreview,receivedDateTime,importance')
          .filter("importance eq 'high'")
          .orderby('receivedDateTime desc')
          .top(20)
          .get();

        const messages = messagesResponse.value || [];
        
        // Filter and transform messages that look like praise/recognition
        const praiseKeywords = ['praise', 'thank', 'recognition', 'kudos', 'appreciate', 'great job', 'well done', 'excellent', 'awesome'];
        const praiseMessages = messages.filter((msg: any) => {
          const subjectLower = (msg.subject || '').toLowerCase();
          const bodyLower = (msg.bodyPreview || '').toLowerCase();
          return praiseKeywords.some(keyword => 
            subjectLower.indexOf(keyword) !== -1 || bodyLower.indexOf(keyword) !== -1
          );
        });

        if (praiseMessages.length > 0) {
          this.praises = praiseMessages.slice(0, 5).map((msg: any) => {
            const badgeInfo = this._extractBadgeType(msg.subject, msg.bodyPreview);
            return {
              id: msg.id,
              displayName: badgeInfo.badgeType,
              givenBy: {
                displayName: msg.from?.emailAddress?.name || 'Colleague',
                id: msg.from?.emailAddress?.address || '',
                photoUrl: this._getUserPhotoUrl(msg.from?.emailAddress?.address)
              },
              message: msg.bodyPreview || 'Thank you for your great work!',
              badgeType: badgeInfo.badgeType,
              backgroundColor: badgeInfo.backgroundColor
            };
          });
          return;
        }
      } catch (messagesError) {
        console.warn('Could not fetch messages for praise detection:', messagesError);
        
        // Check if this is a permission error
        if (messagesError && messagesError.message && messagesError.message.includes('Access is denied')) {
          console.error('⚠️ Mail.Read permission not approved. Go to SharePoint Admin Center → API access to approve permissions.');
        }
      }

      // Fallback: Show encouraging message when no praises are available
      console.warn('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      console.warn('⚠️  NO PRAISES FOUND - ACTION REQUIRED');
      console.warn('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      console.warn('');
      console.warn('📋 TO DISPLAY VIVA INSIGHTS PRAISES:');
      console.warn('');
      console.warn('1. Deploy solution to SharePoint App Catalog');
      console.warn('2. Go to SharePoint Admin Center (https://[tenant]-admin.sharepoint.com)');
      console.warn('3. Navigate to: Advanced → API access');
      console.warn('4. APPROVE these permissions:');
      console.warn('   • Mail.Read (REQUIRED for praises)');
      console.warn('   • Calendars.Read');
      console.warn('   • Tasks.Read');
      console.warn('   • Group.Read.All');
      console.warn('   • Community.Read.All (optional)');
      console.warn('   • User.Read.All');
      console.warn('');
      console.warn('5. Refresh this page after approval');
      console.warn('');
      console.warn('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      this.praises = [];

    } catch (error) {
      console.error('Error fetching praises:', error);
      this.praises = [];
    }
  }

  private _extractBadgeType(subject: string, body: string): { badgeType: string; backgroundColor: string } {
    const text = ((subject || '') + ' ' + (body || '')).toLowerCase();
    
    if (text.indexOf('thank') !== -1 || text.indexOf('thanks') !== -1) {
      return { badgeType: 'Thank You', backgroundColor: '#FFB900' };
    } else if (text.indexOf('awesome') !== -1 || text.indexOf('amazing') !== -1) {
      return { badgeType: 'Awesome', backgroundColor: '#00B294' };
    } else if (text.indexOf('great') !== -1 || text.indexOf('excellent') !== -1 || text.indexOf('well done') !== -1) {
      return { badgeType: 'Great Job', backgroundColor: '#E74856' };
    } else if (text.indexOf('kudos') !== -1) {
      return { badgeType: 'Kudos', backgroundColor: '#8764B8' };
    } else if (text.indexOf('badge') !== -1 || text.indexOf('award') !== -1) {
      return { badgeType: 'Badge Earned', backgroundColor: '#CA5010' };
    } else if (text.indexOf('appreciate') !== -1) {
      return { badgeType: 'Appreciation', backgroundColor: '#0078D4' };
    }
    
    return { badgeType: 'Recognition', backgroundColor: '#00A4EF' };
  }

  private _extractPraiseMessage(body: string): string {
    if (!body) {
      return 'Thank you for your great work!';
    }

    // Remove HTML tags
    let message = body.replace(/<[^>]*>/g, '');
    
    // Try to extract the actual praise message from common patterns
    // Pattern 1: "Message: [text]"
    let messageMatch = message.match(/message:\s*(.+?)(?:\n|$)/i);
    if (messageMatch && messageMatch[1]) {
      return messageMatch[1].trim().substring(0, 150);
    }
    
    // Pattern 2: Text after "sent you praise" or "received praise"
    messageMatch = message.match(/(?:sent you praise|received praise)[:\s]+(.+?)(?:\n|$)/i);
    if (messageMatch && messageMatch[1]) {
      return messageMatch[1].trim().substring(0, 150);
    }
    
    // Pattern 3: Look for quoted text
    messageMatch = message.match(/[""](.+?)[""]|'(.+?)'/);
    if (messageMatch) {
      const extractedMsg = messageMatch[1] || messageMatch[2];
      if (extractedMsg && extractedMsg.length > 10) {
        return extractedMsg.trim().substring(0, 150);
      }
    }
    
    // Fallback: Take first meaningful sentence
    const sentences = message.split(/[.!?]+/).filter(s => s.trim().length > 20);
    if (sentences.length > 0) {
      return sentences[0].trim().substring(0, 150);
    }
    
    // Last resort: trim and return preview
    return message.trim().substring(0, 150) || 'Thank you for your great work!';
  }

  private _getUserPhotoUrl(userEmail: string | undefined): string {
    if (!userEmail) {
      return '';
    }
    // Return empty string - photos will be handled as initials in the UI
    // Actual photo fetching can cause 404 errors if user doesn't have a photo
    return '';
  }

  private _generateCalendar(): void {
    const year = this.currentMonth.getFullYear();
    const month = this.currentMonth.getMonth();
    
    const firstDay = new Date(year, month, 1);
    const startDate = new Date(year, month, 1);
    startDate.setDate(startDate.getDate() - firstDay.getDay());
    
    this.calendarDays = [];
    const currentDate = new Date(startDate.getTime());
    
    for (let i = 0; i < 42; i++) {
      const dayEvents = this.events.filter(event => {
        const eventDate = new Date(event.start.dateTime);
        return eventDate.toDateString() === currentDate.toDateString();
      });
      
      const today = new Date();
      this.calendarDays.push({
        date: new Date(currentDate.getTime()),
        day: currentDate.getDate(),
        isCurrentMonth: currentDate.getMonth() === month,
        isToday: currentDate.toDateString() === today.toDateString(),
        hasEvents: dayEvents.length > 0,
        events: dayEvents
      });
      
      currentDate.setDate(currentDate.getDate() + 1);
    }
  }

  private _renderCalendar(): void {
    const monthDisplay = this.domElement.querySelector('#monthDisplay');
    if (monthDisplay) {
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
      monthDisplay.textContent = `${monthNames[this.currentMonth.getMonth()]} ${this.currentMonth.getFullYear()}`;
    }

    const calendarGrid = this.domElement.querySelector('#calendarGrid');
    if (!calendarGrid) return;

    let html = '<div class="' + styles.calendarWeekdays + '">';
    ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'].forEach(day => {
      html += `<div class="${styles.weekday}">${day}</div>`;
    });
    html += '</div><div class="' + styles.calendarDays + '">';

    this.calendarDays.forEach((calDay, index) => {
      const classes = [styles.calendarDay];
      if (!calDay.isCurrentMonth) classes.push(styles.otherMonth);
      if (calDay.isToday) classes.push(styles.today);
      if (calDay.hasEvents) classes.push(styles.hasEvents);

      html += `
        <div class="${classes.join(' ')}" data-date="${calDay.date.toISOString()}" data-index="${index}">
          <span class="${styles.dayNumber}">${calDay.day}</span>
          ${calDay.hasEvents ? `<span class="${styles.eventDot}"></span>` : ''}
        </div>`;
    });

    html += '</div>';
    calendarGrid.innerHTML = html;

    // Add hover event listeners
    this.domElement.querySelectorAll(`.${styles.calendarDay}.${styles.hasEvents}`).forEach((dayEl: Element) => {
      dayEl.addEventListener('mouseenter', (e) => this._showDayEvents(e));
      dayEl.addEventListener('mouseleave', () => this._hideDayEvents());
    });
  }

  private _renderEvents(): void {
    const eventsList = this.domElement.querySelector('#eventsList');
    if (!eventsList) return;

    const upcomingEvents = this.events
      .filter(e => new Date(e.start.dateTime) >= new Date())
      .slice(0, 3); // Reduced to 3 events for compact view

    if (upcomingEvents.length === 0) {
      eventsList.innerHTML = `<div class="${styles.emptyState}">No upcoming events</div>`;
      return;
    }

    let html = '';
    upcomingEvents.forEach(event => {
      const startDate = new Date(event.start.dateTime);
      const monthShort = startDate.toLocaleString('en-US', { month: 'short' }).toUpperCase();
      const day = startDate.getDate();
      const time = startDate.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true });

      html += `
        <div class="${styles.eventCard}">
          <div class="${styles.eventDate}">
            <div class="${styles.eventMonth}">${monthShort}</div>
            <div class="${styles.eventDay}">${day}</div>
          </div>
          <div class="${styles.eventDetails}">
            <div class="${styles.eventTitle}">${escape(event.subject)} | ${time}</div>
            <div class="${styles.eventMeta}">${escape(event.bodyPreview.substring(0, 50))}${event.bodyPreview.length > 50 ? '...' : ''}</div>
          </div>
          <div class="${styles.eventRibbon}"></div>
        </div>`;
    });

    eventsList.innerHTML = html;
  }

  private _renderTasks(): void {
    const tasksList = this.domElement.querySelector('#tasksList');
    if (!tasksList) return;

    let filteredTasks = this.tasks;
    if (this.searchQuery) {
      const query = this.searchQuery.toLowerCase();
      filteredTasks = this.tasks.filter(task => 
        task.title.toLowerCase().indexOf(query) !== -1 ||
        (this.plans.get(task.planId) || '').toLowerCase().indexOf(query) !== -1
      );
    }

    if (filteredTasks.length === 0) {
      tasksList.innerHTML = `<div class="${styles.emptyState}">No tasks found</div>`;
      return;
    }

    let html = '';
    filteredTasks.forEach(task => {
      const planName = this.plans.get(task.planId) || 'Unknown Plan';
      const dueDate = task.dueDateTime ? new Date(task.dueDateTime).toLocaleDateString() : 'No due date';
      const priorityClass = task.priority <= 3 ? styles.highPriority : '';
      const progressText = task.percentComplete === 100 ? 'Completed' : 
                          task.percentComplete > 0 ? 'In Progress' : 'Not Started';

      html += `
        <div class="${styles.taskCard}" data-task-id="${task.id}">
          <div class="${styles.taskHeader}">
            <div class="${styles.taskTitle}">${escape(task.title)}</div>
          </div>
          <div class="${styles.taskPlan}">${escape(planName)}</div>
          <div class="${styles.taskMeta}">
            <span class="${styles.taskDue}">${dueDate}</span>
            <span class="${styles.taskPriority} ${priorityClass}">${task.priority <= 3 ? 'High' : 'Normal'}</span>
            <span class="${styles.taskProgress}">${progressText}</span>
            ${task.hasDescription ? `<span class="${styles.taskAttachment}">📎</span>` : ''}
          </div>
        </div>`;
    });

    tasksList.innerHTML = html;

    // Add click listeners for task details
    this.domElement.querySelectorAll(`.${styles.taskCard}`).forEach((card: Element) => {
      card.addEventListener('click', (e) => {
        const taskId = (e.currentTarget as HTMLElement).dataset.taskId;
        if (taskId) this._showTaskDetails(taskId);
      });
    });
  }

  private _renderPraises(): void {
    const praisesList = this.domElement.querySelector('#praisesList');
    if (!praisesList) return;

    if (this.praises.length === 0) {
      praisesList.innerHTML = `<div class="${styles.emptyState}">No praises yet</div>`;
      return;
    }

    let html = '';
    this.praises.forEach(praise => {
      const bgColor = praise.backgroundColor;
      const opacity = 0.25;
      const rgbaColor = this._hexToRgba(bgColor, opacity);

      html += `
        <div class="${styles.praiseCard}" style="background-color: ${rgbaColor};">
          <div class="${styles.praiseProfile}">
            <div class="${styles.praiseAvatar}">${praise.givenBy.displayName.charAt(0).toUpperCase()}</div>
          </div>
          <div class="${styles.praiseContent}">
            <div class="${styles.praiseTitle}">${escape(praise.badgeType)}</div>
            <div class="${styles.praiseSender}">from ${escape(praise.givenBy.displayName)}</div>
            <div class="${styles.praiseMessage}">${escape(praise.message)}</div>
          </div>
        </div>`;
    });

    praisesList.innerHTML = html;
  }

  private _attachEventListeners(): void {
    // Calendar navigation
    const prevBtn = this.domElement.querySelector('#prevMonth');
    const nextBtn = this.domElement.querySelector('#nextMonth');
    
    if (prevBtn) {
      prevBtn.addEventListener('click', () => {
        this.currentMonth.setMonth(this.currentMonth.getMonth() - 1);
        this._fetchEvents().then(() => {
          this._generateCalendar();
          this._renderCalendar();
          this._renderEvents();
        });
      });
    }
    
    if (nextBtn) {
      nextBtn.addEventListener('click', () => {
        this.currentMonth.setMonth(this.currentMonth.getMonth() + 1);
        this._fetchEvents().then(() => {
          this._generateCalendar();
          this._renderCalendar();
          this._renderEvents();
        });
      });
    }

    // Task search
    const searchInput = this.domElement.querySelector('#taskSearch') as HTMLInputElement;
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        this.searchQuery = (e.target as HTMLInputElement).value;
        this._renderTasks();
      });
    }

    // Add task button
    const addTaskBtn = this.domElement.querySelector('#addTaskBtn');
    if (addTaskBtn) {
      addTaskBtn.addEventListener('click', () => this._showAddTaskModal());
    }

    // Close modal
    const closeModal = this.domElement.querySelector('#closeModal');
    const modal = this.domElement.querySelector('#taskModal');
    if (closeModal && modal) {
      closeModal.addEventListener('click', () => {
        (modal as HTMLElement).style.display = 'none';
      });
      
      modal.addEventListener('click', (e) => {
        if (e.target === modal) {
          (modal as HTMLElement).style.display = 'none';
        }
      });
    }
  }

  private _showDayEvents(event: Event): void {
    const target = event.currentTarget as HTMLElement;
    const index = parseInt(target.dataset.index || '0');
    const calDay = this.calendarDays[index];

    if (!calDay || !calDay.hasEvents) return;

    let tooltip = target.querySelector(`.${styles.dayTooltip}`) as HTMLElement;
    if (!tooltip) {
      tooltip = document.createElement('div');
      tooltip.className = styles.dayTooltip;
      target.appendChild(tooltip);
    }

    let html = '';
    calDay.events.forEach(event => {
      const time = new Date(event.start.dateTime).toLocaleTimeString('en-US', { 
        hour: 'numeric', 
        minute: '2-digit', 
        hour12: true 
      });
      html += `
        <div class="${styles.tooltipEvent}">
          <strong>${escape(event.subject)}</strong><br/>
          ${time}${event.location.displayName ? ` - ${escape(event.location.displayName)}` : ''}<br/>
          <small>${escape(event.bodyPreview.substring(0, 60))}...</small>
        </div>`;
    });

    tooltip.innerHTML = html;
  }

  private _hideDayEvents(): void {
    this.domElement.querySelectorAll(`.${styles.dayTooltip}`).forEach(tooltip => {
      tooltip.remove();
    });
  }

  private async _showTaskDetails(taskId: string): Promise<void> {
    const modal = this.domElement.querySelector('#taskModal') as HTMLElement;
    const modalBody = this.domElement.querySelector('#modalBody');
    if (!modal || !modalBody) return;

    modalBody.innerHTML = '<div class="' + styles.loading + '">Loading task details...</div>';
    modal.style.display = 'block';

    try {
      let task: any;
      let taskDetails: any = {};
      let planName: string = '';
      let dueDate: string = '';
      let isToDoTask = false;
      
      // Check if this is a To Do task
      if (taskId.indexOf('TODO_') === 0) {
        isToDoTask = true;
        const actualTaskId = taskId.replace('TODO_', '');
        
        // Find the task in To Do lists
        const listsResponse = await this.graphClient.api('/me/todo/lists').get();
        const taskLists = listsResponse.value || [];
        
        for (const list of taskLists) {
          try {
            task = await this.graphClient
              .api(`/me/todo/lists/${list.id}/tasks/${actualTaskId}`)
              .get();
            
            planName = '📋 My Tasks (Private)';
            dueDate = task.dueDateTime?.dateTime ? new Date(task.dueDateTime.dateTime).toLocaleDateString() : 'No due date';
            taskDetails.description = task.body?.content || '';
            break;
          } catch (e) {
            continue;
          }
        }
        
        if (!task) {
          throw new Error('Task not found in any To Do list');
        }
      } else {
        // Planner task - use existing logic
        task = await this.graphClient
          .api(`/planner/tasks/${taskId}`)
          .get();
        
        taskDetails = await this.graphClient
          .api(`/planner/tasks/${taskId}/details`)
          .get();

        planName = this.plans.get(task.planId) || 'Unknown Plan';
        dueDate = task.dueDateTime ? new Date(task.dueDateTime).toLocaleDateString() : 'No due date';
      }

      const assignedUsers = Object.keys(task.assignments || {});
      const priority = isToDoTask 
        ? (task.importance === 'high' ? 'High' : task.importance === 'low' ? 'Low' : 'Normal')
        : (task.priority <= 3 ? 'High' : 'Normal');
      const percentComplete = isToDoTask ? (task.status === 'completed' ? 100 : 0) : task.percentComplete;
      const etag = task['@odata.etag'] || '';

      let html = `
        <h2 class="${styles.modalTitle}">${escape(task.title)}</h2>
        <div class="${styles.modalSection}">
          <strong>Plan:</strong> ${escape(planName)}<br/>
          <strong>Due Date:</strong> ${dueDate}<br/>
          <strong>Priority:</strong> ${priority}<br/>
          <strong>Created:</strong> ${new Date(task.createdDateTime).toLocaleDateString()}
        </div>
        
        <div class="${styles.modalSection}">
          <strong>Update Progress</strong>
          <div class="${styles.progressEditor}">
            <label>Current Progress: <span id="currentProgress">${percentComplete}</span>%</label>
            <input type="range" id="updateProgress" min="0" max="100" value="${percentComplete}" step="25" />
            <div class="${styles.progressLabels}">
              <span>0%</span>
              <span>25%</span>
              <span>50%</span>
              <span>75%</span>
              <span>100%</span>
            </div>
            <button id="saveProgress" class="${styles.primaryButton}" data-task-id="${taskId}" data-etag="${etag}">Save Progress</button>
          </div>
        </div>`;

      if (taskDetails.description) {
        html += `
          <div class="${styles.modalSection}">
            <strong>Description:</strong><br/>
            <p>${escape(taskDetails.description)}</p>
          </div>`;
      }

      if (!isToDoTask && taskDetails.checklist) {
        html += `<div class="${styles.modalSection}"><strong>Checklist:</strong><ul>`;
        const checklistKeys = Object.keys(taskDetails.checklist);
        for (let i = 0; i < checklistKeys.length; i++) {
          const item = taskDetails.checklist[checklistKeys[i]];
          html += `<li>${item.isChecked ? '✓' : '○'} ${escape(item.title)}</li>`;
        }
        html += `</ul></div>`;
      }

      if (assignedUsers.length > 0) {
        html += `<div class="${styles.modalSection}"><strong>Assigned to:</strong> ${assignedUsers.length} user(s)</div>`;
      }

      modalBody.innerHTML = html;
      
      // Add progress slider update listener
      const progressSlider = this.domElement.querySelector('#updateProgress') as HTMLInputElement;
      const progressDisplay = this.domElement.querySelector('#currentProgress');
      if (progressSlider && progressDisplay) {
        progressSlider.addEventListener('input', () => {
          progressDisplay.textContent = progressSlider.value;
        });
      }
      
      // Add save progress button listener
      const saveBtn = this.domElement.querySelector('#saveProgress');
      if (saveBtn) {
        saveBtn.addEventListener('click', () => this._updateTaskProgress(taskId, etag));
      }
    } catch (error) {
      console.error('Error fetching task details:', error);
      modalBody.innerHTML = '<div class="' + styles.error + '">Failed to load task details</div>';
    }
  }

  private _showAddTaskModal(): void {
    const modal = this.domElement.querySelector('#taskModal') as HTMLElement;
    const modalBody = this.domElement.querySelector('#modalBody');
    if (!modal || !modalBody) return;

    let planOptions = '<option value="">Select a plan</option>';
    planOptions += '<option value="TODO_PRIVATE">📋 My Tasks (Private)</option>';
    this.plans.forEach((title, id) => {
      if (id.indexOf('TODO_') !== 0) { // Don't show auto-added To Do plans in dropdown
        planOptions += `<option value="${id}">${escape(title)}</option>`;
      }
    });

    modalBody.innerHTML = `
      <h2 class="${styles.modalTitle}">Add New Task</h2>
      <form id="addTaskForm" class="${styles.taskForm}">
        <div class="${styles.formGroup}">
          <label>Task Name *</label>
          <input type="text" id="taskName" required />
        </div>
        <div class="${styles.formGroup}">
          <label>Plan *</label>
          <select id="taskPlan" required>${planOptions}</select>
          <div id="taskLocationHint" style="margin-top: 8px; padding: 8px; background: #f3f2f1; border-radius: 4px; font-size: 11px; color: #605e5c; display: none;">
            <strong>💡 Where will this task appear?</strong><br/>
            <span id="taskLocationText"></span>
          </div>
        </div>
        <div class="${styles.formGroup}">
          <label>Due Date</label>
          <input type="date" id="taskDueDate" />
        </div>
        <div class="${styles.formGroup}">
          <label>Priority</label>
          <select id="taskPriority">
            <option value="5">Normal</option>
            <option value="3">High</option>
            <option value="1">Urgent</option>
          </select>
        </div>
        <div class="${styles.formGroup}">
          <label>Progress: <span id="progressValue">0</span>%</label>
          <input type="range" id="taskProgress" min="0" max="100" value="0" step="25" />
          <div class="${styles.progressLabels}">
            <span>Not Started</span>
            <span>In Progress</span>
            <span>Completed</span>
          </div>
        </div>
        <div class="${styles.formGroup}">
          <label>Description</label>
          <textarea id="taskDescription" rows="3" placeholder="Add task description..."></textarea>
        </div>
        <div class="${styles.formActions}">
          <button type="submit" class="${styles.primaryButton}">Create Task</button>
          <button type="button" class="${styles.secondaryButton}" id="cancelTask">Cancel</button>
        </div>
      </form>`;

    modal.style.display = 'block';

    // Progress slider update
    const progressSlider = this.domElement.querySelector('#taskProgress') as HTMLInputElement;
    const progressValue = this.domElement.querySelector('#progressValue');
    if (progressSlider && progressValue) {
      progressSlider.addEventListener('input', () => {
        progressValue.textContent = progressSlider.value;
      });
    }
    
    // Plan selector hint
    const planSelector = this.domElement.querySelector('#taskPlan') as HTMLSelectElement;
    const locationHint = this.domElement.querySelector('#taskLocationHint') as HTMLElement;
    const locationText = this.domElement.querySelector('#taskLocationText') as HTMLElement;
    if (planSelector && locationHint && locationText) {
      planSelector.addEventListener('change', () => {
        const selectedPlan = planSelector.value;
        if (selectedPlan === 'TODO_PRIVATE') {
          locationHint.style.display = 'block';
          locationText.innerHTML = '✅ <strong>Microsoft To Do app</strong> (private task, only you can see)<br/>❌ This will NOT appear in Microsoft Planner';
        } else if (selectedPlan) {
          locationHint.style.display = 'block';
          const planName = this.plans.get(selectedPlan) || 'selected plan';
          locationText.innerHTML = `✅ <strong>Microsoft Planner app</strong> → ${escape(planName)}<br/>❌ This will NOT appear in Microsoft To Do`;
        } else {
          locationHint.style.display = 'none';
        }
      });
    }

    const form = this.domElement.querySelector('#addTaskForm');
    const cancelBtn = this.domElement.querySelector('#cancelTask');

    if (form) {
      form.addEventListener('submit', (e) => {
        e.preventDefault();
        this._createTask();
      });
    }

    if (cancelBtn) {
      cancelBtn.addEventListener('click', () => {
        modal.style.display = 'none';
      });
    }
  }

  private async _createTask(): Promise<void> {
    const taskName = (this.domElement.querySelector('#taskName') as HTMLInputElement)?.value;
    const planId = (this.domElement.querySelector('#taskPlan') as HTMLSelectElement)?.value;
    const dueDate = (this.domElement.querySelector('#taskDueDate') as HTMLInputElement)?.value;
    const priority = parseInt((this.domElement.querySelector('#taskPriority') as HTMLSelectElement)?.value || '5');
    const progress = parseInt((this.domElement.querySelector('#taskProgress') as HTMLInputElement)?.value || '0');
    const description = (this.domElement.querySelector('#taskDescription') as HTMLTextAreaElement)?.value;

    if (!taskName || !planId) {
      alert('Please fill in required fields');
      return;
    }

    try {
      // Check if creating a private To Do task
      if (planId === 'TODO_PRIVATE') {
        console.log('Creating private To Do task...');
        
        // Get default task list
        const listsResponse = await this.graphClient.api('/me/todo/lists').get();
        const defaultList = listsResponse.value.find((l: any) => l.wellknownListName === 'defaultList') || listsResponse.value[0];
        
        if (!defaultList) {
          alert('No To Do list found. Please create one in Microsoft To Do first.');
          return;
        }
        
        const todoTaskData: any = {
          title: taskName,
          importance: priority <= 3 ? 'high' : 'normal'
        };
        
        if (dueDate) {
          todoTaskData.dueDateTime = {
            dateTime: new Date(dueDate).toISOString(),
            timeZone: 'UTC'
          };
        }
        
        if (description) {
          todoTaskData.body = {
            content: description,
            contentType: 'text'
          };
        }
        
        console.log('Creating To Do task:', todoTaskData);
        const createdTodoTask = await this.graphClient
          .api(`/me/todo/lists/${defaultList.id}/tasks`)
          .post(todoTaskData);
        
        console.log('To Do task created successfully:', createdTodoTask);
        console.log('✅ Task will appear in Microsoft To Do app (not Planner)');
      } else {
        // Create Planner task
        console.log('Creating Planner task...');
        const buckets = await this.graphClient.api(`/planner/plans/${planId}/buckets`).get();
        const bucketId = buckets.value[0]?.id;

        if (!bucketId) {
          alert('No buckets found in selected plan');
          return;
        }

        const taskData: any = {
          planId: planId,
          bucketId: bucketId,
          title: taskName,
          priority: priority
        };

        if (dueDate) {
          taskData.dueDateTime = new Date(dueDate).toISOString();
        }

        console.log('Creating task with data:', taskData);

        const newTask = await this.graphClient
          .api('/planner/tasks')
          .post(taskData);

        console.log('Planner task created successfully:', newTask);
        console.log('✅ Task will appear in Microsoft Planner app');
        console.log(`Plan: ${this.plans.get(planId) || planId}`);

        // Update progress if not 0%
        if (progress > 0 && newTask.id) {
          try {
            await this.graphClient
              .api(`/planner/tasks/${newTask.id}`)
              .header('If-Match', newTask['@odata.etag'])
              .patch({
                percentComplete: progress
              });
            console.log('Task progress updated to:', progress);
          } catch (progressError) {
            console.warn('Could not set initial progress:', progressError);
          }
        }

        if (description && newTask.id) {
          try {
            // First, get the task details to retrieve the ETag
            const taskDetails = await this.graphClient
              .api(`/planner/tasks/${newTask.id}/details`)
              .get();

            // Then update with the correct ETag
            await this.graphClient
              .api(`/planner/tasks/${newTask.id}/details`)
              .header('If-Match', taskDetails['@odata.etag'])
              .patch({
                description: description
              });
          } catch (detailsError) {
            console.warn('Could not add task description:', detailsError);
            // Don't fail the whole operation if description update fails
          }
        }
      }

      // Close modal and refresh tasks
      const modal = this.domElement.querySelector('#taskModal') as HTMLElement;
      if (modal) modal.style.display = 'none';
      
      console.log('Refreshing task list...');
      await this._fetchPlannerTasks();
      console.log('Task list refreshed. Total tasks in memory:', this.tasks.length);
      this._renderTasks();
      console.log('Task list re-rendered in UI');
      
      const taskLocation = planId === 'TODO_PRIVATE' 
        ? 'Microsoft To Do app (private task)' 
        : `Microsoft Planner app (${this.plans.get(planId) || 'team task'})`;
      
      alert(`Task created successfully!\n\nCheck: ${taskLocation}\n\nNote: It may take a few seconds for the task to sync.`);
    } catch (error) {
      console.error('Error creating task:', error);
      console.error('Error details:', JSON.stringify(error, null, 2));
      const errorMessage = error && error.message ? error.message : 'Unknown error';
      alert(`Failed to create task: ${errorMessage}\n\nCheck browser console for details.`);
    }
  }

  private async _updateTaskProgress(taskId: string, etag: string): Promise<void> {
    const progressSlider = this.domElement.querySelector('#updateProgress') as HTMLInputElement;
    if (!progressSlider) return;

    const newProgress = parseInt(progressSlider.value);

    try {
      // Check if this is a To Do task
      if (taskId.indexOf('TODO_') === 0) {
        const actualTaskId = taskId.replace('TODO_', '');
        
        // For To Do tasks, we can only mark as completed (100%) or not completed
        const status = newProgress === 100 ? 'completed' : 'notStarted';
        
        // We need to find the list ID for this task
        const listsResponse = await this.graphClient.api('/me/todo/lists').get();
        const taskLists = listsResponse.value || [];
        
        // Find the task in one of the lists
        let found = false;
        for (const list of taskLists) {
          try {
            await this.graphClient
              .api(`/me/todo/lists/${list.id}/tasks/${actualTaskId}`)
              .patch({
                status: status
              });
            found = true;
            break;
          } catch (e) {
            // Task not in this list, continue searching
            continue;
          }
        }
        
        if (!found) {
          throw new Error('Task not found in any To Do list');
        }
      } else {
        // Planner task - use existing logic
          try {
            await this.graphClient
              .api(`/planner/tasks/${taskId}`)
              .header('If-Match', etag)
              .patch({
                percentComplete: newProgress
              });
          } catch (patchError) {
            console.error('Planner PATCH failed:', patchError);
            alert('Failed to update Planner task progress. Please check your permissions or try again.');
            return;
          }
      }

      // Close modal and refresh tasks
      const modal = this.domElement.querySelector('#taskModal') as HTMLElement;
      if (modal) modal.style.display = 'none';
      
      await this._fetchPlannerTasks();
      this._renderTasks();
      
      alert('Task progress updated successfully!');
    } catch (error) {
      console.error('Error updating task progress:', error);
      alert('Failed to update task progress. Please try again.');
    }
  }

  private _renderLoadingSkeleton(type: string): string {
    const skeletonClass = styles.skeleton;
    if (type === 'calendar') {
      return `<div class="${skeletonClass}" style="height: 250px;"></div>`;
    } else if (type === 'events') {
      let html = '';
      for (let i = 0; i < 3; i++) {
        html += `<div class="${skeletonClass}" style="height: 80px; margin-bottom: 10px;"></div>`;
      }
      return html;
    } else if (type === 'tasks') {
      let html = '';
      for (let i = 0; i < 5; i++) {
        html += `<div class="${skeletonClass}" style="height: 100px; margin-bottom: 10px;"></div>`;
      }
      return html;
    } else if (type === 'praises') {
      let html = '';
      for (let i = 0; i < 5; i++) {
        html += `<div class="${skeletonClass}" style="height: 90px; margin-bottom: 10px;"></div>`;
      }
      return html;
    }
    return '';
  }

  private _showError(message: string): void {
    this.domElement.innerHTML = `
      <div class="${styles.errorContainer}">
        <div class="${styles.errorMessage}">${escape(message)}</div>
      </div>`;
  }

  private _hexToRgba(hex: string, alpha: number): string {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
