/**
 * scriptRunner.js - Handles execution of VBScript files and processes their output
 */

import { exec } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';

// Get the directory name of the current module
const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Constants
const SCRIPTS_DIR = path.resolve(__dirname, '../scripts');
const SUCCESS_PREFIX = 'SUCCESS:';
const ERROR_PREFIX = 'ERROR:';
// Default path to cscript.exe, can be overridden with CSCRIPT_PATH environment variable
const CSCRIPT_PATH = process.env.CSCRIPT_PATH || 'C:\\Windows\\System32\\cscript.exe';

/**
 * Executes a VBScript file with the given parameters
 * @param {string} scriptName - Name of the script file (without .vbs extension)
 * @param {Object} params - Parameters to pass to the script
 * @param {string} locale - Language/locale code (e.g., "en", "es") - defaults to "en"
 * @returns {Promise<Object>} - Promise that resolves with the script output
 */
export async function executeScript(scriptName, params = {}, locale = "en") {
  return new Promise((resolve, reject) => {
    // Build the command
    const scriptPath = path.join(SCRIPTS_DIR, `${scriptName}.vbs`);
    // Use absolute path to cscript.exe to avoid PATH resolution issues
    let command = `"${CSCRIPT_PATH}" //NoLogo "${scriptPath}"`;
    
    // Add parameters
    for (const [key, value] of Object.entries(params)) {
      if (value !== undefined && value !== null && value !== '') {
        // Handle special characters in values
        const escapedValue = value.toString().replace(/"/g, '\\"');
        command += ` /${key}:"${escapedValue}"`;
      }
    }
    
    // Add locale parameter if specified
    if (locale && locale !== '') {
      command += ` /locale:"${locale}"`;
    }
    
    // Execute the command
    exec(command, (error, stdout, stderr) => {
      // Check for execution errors
      if (error && !stdout.includes(SUCCESS_PREFIX)) {
        return reject(new Error(`Script execution failed: ${error.message}`));
      }
      
      // Check for script errors
      if (stdout.includes(ERROR_PREFIX)) {
        const errorMessage = stdout.substring(stdout.indexOf(ERROR_PREFIX) + ERROR_PREFIX.length).trim();
        return reject(new Error(`Script error: ${errorMessage}`));
      }
      
      // Process successful output
      if (stdout.includes(SUCCESS_PREFIX)) {
        try {
          const jsonStr = stdout.substring(stdout.indexOf(SUCCESS_PREFIX) + SUCCESS_PREFIX.length).trim();
          const result = JSON.parse(jsonStr);
          return resolve(result);
        } catch (parseError) {
          return reject(new Error(`Failed to parse script output: ${parseError.message}`));
        }
      }
      
      // If we get here, something unexpected happened
      reject(new Error(`Unexpected script output: ${stdout}`));
    });
  });
}

/**
 * Lists calendar events within a specified date range
 * @param {string} startDate - Start date in locale appropriate format (MM/DD/YYYY for en, DD/MM/YYYY for es)
 * @param {string} endDate - End date in locale appropriate format (optional)
 * @param {string} calendar - Calendar name (optional)
 * @param {string} locale - Language/locale code (e.g., "en", "es") - defaults to "en"
 * @returns {Promise<Array>} - Promise that resolves with an array of events
 */
export async function listEvents(startDate, endDate, calendar, locale = "en") {
  return executeScript('listEvents', { startDate, endDate, calendar }, locale);
}

/**
 * Creates a new calendar event
 * @param {Object} eventDetails - Event details
 * @param {string} locale - Language/locale code (e.g., "en", "es") - defaults to "en"
 * @returns {Promise<Object>} - Promise that resolves with the created event ID
 */
export async function createEvent(eventDetails, locale = "en") {
  return executeScript('createEvent', eventDetails, locale);
}

/**
 * Finds free time slots in the calendar
 * @param {string} startDate - Start date in locale appropriate format (MM/DD/YYYY for en, DD/MM/YYYY for es)
 * @param {string} endDate - End date in locale appropriate format (optional)
 * @param {number} duration - Duration in minutes (optional)
 * @param {number} workDayStart - Work day start hour (0-23) (optional)
 * @param {number} workDayEnd - Work day end hour (0-23) (optional)
 * @param {string} calendar - Calendar name (optional)
 * @param {string} locale - Language/locale code (e.g., "en", "es") - defaults to "en"
 * @returns {Promise<Array>} - Promise that resolves with an array of free time slots
 */
export async function findFreeSlots(startDate, endDate, duration, workDayStart, workDayEnd, calendar, locale = "en") {
  return executeScript('findFreeSlots', {
    startDate,
    endDate,
    duration,
    workDayStart,
    workDayEnd,
    calendar
  }, locale);
}

/**
 * Gets the response status of meeting attendees
 * @param {string} eventId - Event ID
 * @param {string} calendar - Calendar name (optional)
 * @param {string} locale - Language/locale code (e.g., "en", "es") - defaults to "en"
 * @returns {Promise<Object>} - Promise that resolves with meeting details and attendee status
 */
export async function getAttendeeStatus(eventId, calendar, locale = "en") {
  return executeScript('getAttendeeStatus', { eventId, calendar }, locale);
}

/**
 * Deletes a calendar event by its ID
 * @param {string} eventId - Event ID
 * @param {string} calendar - Calendar name (optional)
 * @param {string} locale - Language/locale code (e.g., "en", "es") - defaults to "en"
 * @returns {Promise<Object>} - Promise that resolves with the deletion result
 */
export async function deleteEvent(eventId, calendar, locale = "en") {
  return executeScript('deleteEvent', { eventId, calendar }, locale);
}

/**
 * Updates an existing calendar event
 * @param {string} eventId - Event ID to update
 * @param {string} subject - New subject (optional)
 * @param {string} startDate - New start date in locale appropriate format (MM/DD/YYYY for en, DD/MM/YYYY for es) (optional)
 * @param {string} startTime - New start time in HH:MM AM/PM format (optional)
 * @param {string} endDate - New end date in locale appropriate format (MM/DD/YYYY for en, DD/MM/YYYY for es) (optional)
 * @param {string} endTime - New end time in HH:MM AM/PM format (optional)
 * @param {string} location - New location (optional)
 * @param {string} body - New body/description (optional)
 * @param {string} calendar - Calendar name (optional)
 * @param {string} locale - Language/locale code (e.g., "en", "es") - defaults to "en"
 * @returns {Promise<Object>} - Promise that resolves with the update result
 */
export async function updateEvent(eventId, subject, startDate, startTime, endDate, endTime, location, body, calendar, locale = "en") {
  return executeScript('updateEvent', {
    eventId,
    subject,
    startDate,
    startTime,
    endDate,
    endTime,
    location,
    body,
    calendar
  }, locale);
}

/**
 * Lists available calendars
 * @param {string} locale - Language/locale code (e.g., "en", "es") - defaults to "en"
 * @returns {Promise<Array>} - Promise that resolves with an array of calendars
 */
export async function getCalendars(locale = "en") {
  return executeScript('getCalendars', {}, locale);
}
