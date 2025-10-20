/**
 * Google Calendar Multi-View Application v2.3
 * メインバックエンドスクリプト（完全版）
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('カレンダー統合ビュー')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getAvailableCalendars() {
  try {
    var calendars = [];
    
    var ownCalendars = CalendarApp.getAllCalendars();
    ownCalendars.forEach(function(cal) {
      calendars.push({
        id: cal.getId(),
        name: cal.getName(),
        color: cal.getColor(),
        isOwned: true,
        isEditable: true
      });
    });
    
    var subscribedCalendars = CalendarApp.getAllOwnedCalendars();
    subscribedCalendars.forEach(function(cal) {
      var calId = cal.getId();
      if (!calendars.some(c => c.id === calId)) {
        calendars.push({
          id: calId,
          name: cal.getName(),
          color: cal.getColor(),
          isOwned: false,
          isEditable: cal.isOwnedByMe()
        });
      }
    });
    
    return calendars;
  } catch (error) {
    Logger.log('Error getting calendars: ' + error.toString());
    return [];
  }
}

function getCalendarEvents(calendarIds, startDate, endDate) {
  try {
    var start = new Date(startDate);
    var end = new Date(endDate);
    var allEvents = {};
    
    calendarIds.forEach(function(calId) {
      try {
        var calendar = CalendarApp.getCalendarById(calId);
        if (!calendar) {
          Logger.log('Calendar not found: ' + calId);
          return;
        }
        
        var events = calendar.getEvents(start, end);
        
        allEvents[calId] = events.map(function(event) {
          return {
            id: event.getId(),
            title: event.getTitle(),
            start: event.getStartTime().toISOString(),
            end: event.getEndTime().toISOString(),
            isAllDay: event.isAllDayEvent(),
            color: event.getColor(),
            description: event.getDescription() || '',
            location: event.getLocation() || '',
            calendarId: calId
          };
        });
      } catch (e) {
        Logger.log('Error getting events for calendar ' + calId + ': ' + e.toString());
        allEvents[calId] = [];
      }
    });
    
    return allEvents;
  } catch (error) {
    Logger.log('Error in getCalendarEvents: ' + error.toString());
    return {};
  }
}

function getEventDetails(calendarId, eventId) {
  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      return { success: false, error: 'カレンダーが見つかりません' };
    }
    
    var event = calendar.getEventById(eventId);
    if (!event) {
      return { success: false, error: 'イベントが見つかりません' };
    }
    
    return {
      success: true,
      event: {
        id: event.getId(),
        title: event.getTitle(),
        start: event.getStartTime().toISOString(),
        end: event.getEndTime().toISOString(),
        isAllDay: event.isAllDayEvent(),
        color: event.getColor(),
        description: event.getDescription() || '',
        location: event.getLocation() || '',
        calendarId: calendarId
      }
    };
  } catch (error) {
    Logger.log('Error getting event details: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function createEvent(calendarId, eventData) {
  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      return { success: false, error: 'カレンダーが見つかりません' };
    }
    
    var startDate = new Date(eventData.start);
    var endDate = new Date(eventData.end);
    
    var event;
    if (eventData.isAllDay) {
      event = calendar.createAllDayEvent(
        eventData.title,
        startDate,
        endDate
      );
    } else {
      event = calendar.createEvent(
        eventData.title,
        startDate,
        endDate
      );
    }
    
    if (eventData.description) {
      event.setDescription(eventData.description);
    }
    
    if (eventData.location) {
      event.setLocation(eventData.location);
    }
    
    if (eventData.color) {
      event.setColor(eventData.color);
    }
    
    return {
      success: true,
      eventId: event.getId()
    };
  } catch (error) {
    Logger.log('Error creating event: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function updateEvent(calendarId, eventId, eventData) {
  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      return { success: false, error: 'カレンダーが見つかりません' };
    }
    
    var event = calendar.getEventById(eventId);
    if (!event) {
      return { success: false, error: 'イベントが見つかりません' };
    }
    
    if (eventData.title) {
      event.setTitle(eventData.title);
    }
    
    if (eventData.start && eventData.end) {
      var startDate = new Date(eventData.start);
      var endDate = new Date(eventData.end);
      event.setTime(startDate, endDate);
    }
    
    if (eventData.description !== undefined) {
      event.setDescription(eventData.description);
    }
    
    if (eventData.location !== undefined) {
      event.setLocation(eventData.location);
    }
    
    if (eventData.color) {
      event.setColor(eventData.color);
    }
    
    return { success: true };
  } catch (error) {
    Logger.log('Error updating event: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function deleteEvent(calendarId, eventId) {
  try {
    var calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      return { success: false, error: 'カレンダーが見つかりません' };
    }
    
    var event = calendar.getEventById(eventId);
    if (!event) {
      return { success: false, error: 'イベントが見つかりません' };
    }
    
    event.deleteEvent();
    
    return { success: true };
  } catch (error) {
    Logger.log('Error deleting event: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function saveUserSettings(settings) {
  try {
    var userEmail = getCurrentUserEmail();
    var userProperties = PropertiesService.getUserProperties();
    
    userProperties.setProperty('settings_' + userEmail, JSON.stringify(settings));
    return { success: true };
  } catch (error) {
    Logger.log('Error saving settings: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function loadUserSettings() {
  try {
    var userEmail = getCurrentUserEmail();
    var userProperties = PropertiesService.getUserProperties();
    
    var settingsJson = userProperties.getProperty('settings_' + userEmail);
    
    if (settingsJson) {
      return JSON.parse(settingsJson);
    } else {
      return {
        selectedCalendars: [],
        calendarOrder: [],
        viewMode: 'week',
        startDay: 1,
        darkMode: false,
        customColors: {},
        columnWidth: 150,
        calendarColumnWidth: 180
      };
    }
  } catch (error) {
    Logger.log('Error loading settings: ' + error.toString());
    return {
      selectedCalendars: [],
      calendarOrder: [],
      viewMode: 'week',
      startDay: 1,
      darkMode: false,
      customColors: {},
      columnWidth: 150,
      calendarColumnWidth: 180
    };
  }
}

function updateCalendarOrder(calendarOrder) {
  try {
    var settings = loadUserSettings();
    settings.calendarOrder = calendarOrder;
    return saveUserSettings(settings);
  } catch (error) {
    Logger.log('Error updating calendar order: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function updateSelectedCalendars(selectedCalendars) {
  try {
    var settings = loadUserSettings();
    settings.selectedCalendars = selectedCalendars;
    
    selectedCalendars.forEach(function(calId) {
      if (settings.calendarOrder.indexOf(calId) === -1) {
        settings.calendarOrder.push(calId);
      }
    });
    
    settings.calendarOrder = settings.calendarOrder.filter(function(calId) {
      return selectedCalendars.indexOf(calId) !== -1;
    });
    
    return saveUserSettings(settings);
  } catch (error) {
    Logger.log('Error updating selected calendars: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function saveCalendarCustomColor(calendarId, color) {
  try {
    var settings = loadUserSettings();
    if (!settings.customColors) {
      settings.customColors = {};
    }
    settings.customColors[calendarId] = color;
    return saveUserSettings(settings);
  } catch (error) {
    Logger.log('Error saving custom color: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}