// Content script for extracting schedule data from the schedule page

(function() {
  'use strict';

  // Cache of column indexes by header name
  let columnIndexes = null;

  // Listen for messages from popup
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'getSelectedRows') {
      const selectedRows = getSelectedScheduleRows();
      sendResponse({ success: true, data: selectedRows });
      return true; // Keep channel open for async response
    }
    
    if (request.action === 'highlightRows') {
      highlightSelectedRows();
      sendResponse({ success: true });
      return true;
    }
  });

  /**
   * Determine column indexes based on table header text
   */
  function getColumnIndexes() {
    if (columnIndexes) return columnIndexes;

    const headers = document.querySelectorAll('table thead tr th, .table thead tr th, [role="columnheader"]');
    const map = {
      // Original mySchedule columns
      type: -1,
      refNumber: -1,
      name: -1,
      description: -1,
      startDate: -1,
      endDate: -1,
      office: -1,
      projectNumber: -1,
      // laborSchedule token-page columns
      status: -1,
      confirm: -1,
      deny: -1,
      talent: -1,
      task: -1,
      jobNumber: -1,
      jobName: -1,
      client: -1,
      venueName: -1,
      venueRoom: -1,
      address: -1,
      salesperson: -1,
      orderStatus: -1,
      laborCustom: -1
    };

    headers.forEach((th, idx) => {
      const text = (th.textContent || '').trim().toLowerCase().replace(/\s+/g, ' ');

      if (text.includes('type')) {
        map.type = idx;
      } else if (text.includes('ref')) {
        map.refNumber = idx;
      } else if (text === 'name') {
        map.name = idx;
      } else if (text.includes('description')) {
        map.description = idx;
      } else if (text.includes('start date')) {
        map.startDate = idx;
      } else if (text.includes('end date')) {
        map.endDate = idx;
      } else if (text.includes('office')) {
        map.office = idx;
      } else if (text.includes('project')) {
        map.projectNumber = idx;
      } else if (text === 'status') {
        map.status = idx;
      } else if (text === 'confirm') {
        map.confirm = idx;
      } else if (text === 'deny') {
        map.deny = idx;
      } else if (text === 'talent') {
        map.talent = idx;
      } else if (text === 'task') {
        map.task = idx;
      } else if (text === 'job #') {
        map.jobNumber = idx;
      } else if (text === 'job name') {
        map.jobName = idx;
      } else if (text === 'client') {
        map.client = idx;
      } else if (text === 'venue name') {
        map.venueName = idx;
      } else if (text === 'venue room') {
        map.venueRoom = idx;
      } else if (text === 'address') {
        map.address = idx;
      } else if (text === 'salesperson') {
        map.salesperson = idx;
      } else if (text === 'order status') {
        map.orderStatus = idx;
      } else if (text === 'labor custom') {
        map.laborCustom = idx;
      }
    });

    columnIndexes = map;
    return map;
  }

  /**
   * Extract schedule data from selected table rows
   */
  function getSelectedScheduleRows() {
    const rows = [];
    
    // Find the schedule table - adjust selector based on actual page structure
    // Looking for table rows with checkboxes
    const tableRows = document.querySelectorAll('table tbody tr, .table tbody tr, [role="row"]');
    
    tableRows.forEach((row, index) => {
      // Check if row has a checked checkbox
      const checkbox = row.querySelector('input[type="checkbox"]');
      if (checkbox && checkbox.checked) {
        const rowData = extractRowData(row, index);
        if (rowData) {
          rows.push(rowData);
        }
      }
    });

    // If no checkboxes are checked, try to get all visible rows
    // This is a fallback for initial testing
    if (rows.length === 0) {
      console.log('No selected rows found, extracting all visible rows for testing');
      tableRows.forEach((row, index) => {
        const rowData = extractRowData(row, index);
        if (rowData && !isHeaderRow(rowData)) {
          rows.push(rowData);
        }
      });
    }

    return rows;
  }

  /**
   * Extract data from a single table row
   */
  function extractRowData(row, index) {
    try {
      const cells = row.querySelectorAll('td, [role="gridcell"]');
      
      if (cells.length < 6) {
        return null; // Not a data row
      }

      // Get dynamic column indexes based on header row
      const cols = getColumnIndexes();

      // Helper to get text by logical column
      const getByIndex = (colIdx) => {
        if (colIdx == null || colIdx < 0 || colIdx >= cells.length) return '';
        return getCellText(cells[colIdx]);
      };

      // Extract data based on detected column positions
      const data = {
        index: index,
        type: getByIndex(cols.type),
        refNumber: getByIndex(cols.refNumber),
        name: getByIndex(cols.name),
        description: getByIndex(cols.description),
        startDate: parseDate(getByIndex(cols.startDate)),
        endDate: parseDate(getByIndex(cols.endDate)),
        office: getByIndex(cols.office),
        projectNumber: getByIndex(cols.projectNumber),
        // Extended fields for laborSchedule token page
        status: getByIndex(cols.status),
        confirm: getByIndex(cols.confirm),
        deny: getByIndex(cols.deny),
        talent: getByIndex(cols.talent),
        task: getByIndex(cols.task),
        jobNumber: getByIndex(cols.jobNumber),
        jobName: getByIndex(cols.jobName),
        client: getByIndex(cols.client),
        venueName: getByIndex(cols.venueName),
        venueRoom: getByIndex(cols.venueRoom),
        address: getByIndex(cols.address),
        salesperson: getByIndex(cols.salesperson),
        orderStatus: getByIndex(cols.orderStatus),
        laborCustom: getByIndex(cols.laborCustom)
      };

      // Clean up ref number (remove external link icon if present)
      if (data.refNumber) {
        data.refNumber = data.refNumber.replace(/[\u2197\u2192]/g, '').trim();
      }

      // Clean up project number
      if (data.projectNumber) {
        data.projectNumber = data.projectNumber.replace(/[\u2197\u2192]/g, '').trim();
      }

      // If we couldn't parse valid start/end dates, skip this row (likely group or header)
      if (!data.startDate || !data.endDate || isNaN(data.startDate.getTime()) || isNaN(data.endDate.getTime())) {
        return null;
      }

      return data;
    } catch (error) {
      console.error('Error extracting row data:', error);
      return null;
    }
  }

  /**
   * Get text content from a cell, handling links
   */
  function getCellText(cell) {
    if (!cell) return '';
    
    // If cell contains a link, prefer link text
    const link = cell.querySelector('a');
    if (link) {
      return link.textContent.trim();
    }
    
    return cell.textContent.trim();
  }

  /**
   * Parse date string from schedule format to Date object
   * Expected format: "MM/DD/YYYY HH:MM AM/PM" or "M/D/YYYY H:MM AM/PM"
   */
  function parseDate(dateString) {
    if (!dateString) return null;

    // Remove extra spaces and normalize
    dateString = dateString.trim();
    
    // Pattern: "1/12/2026 8:00 AM" or "01/12/2026 08:00 AM"
    const pattern = /(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})\s+(AM|PM)/i;
    const match = dateString.match(pattern);
    
    if (!match) {
      console.warn('Could not parse date:', dateString);
      return null;
    }

    const month = parseInt(match[1]) - 1; // JS months are 0-indexed
    const day = parseInt(match[2]);
    const year = parseInt(match[3]);
    let hour = parseInt(match[4]);
    const minute = parseInt(match[5]);
    const ampm = match[6].toUpperCase();

    // Convert to 24-hour format
    if (ampm === 'PM' && hour !== 12) {
      hour += 12;
    } else if (ampm === 'AM' && hour === 12) {
      hour = 0;
    }

    const date = new Date(year, month, day, hour, minute);
    return date;
  }

  /**
   * Check if row is a header row
   */
  function isHeaderRow(rowData) {
    // Header rows typically have empty or special values
    return !rowData.refNumber || 
           rowData.type === 'Type' || 
           rowData.name === 'Name' ||
           rowData.refNumber === 'Ref #';
  }

  /**
   * Highlight selected rows for visual feedback
   */
  function highlightSelectedRows() {
    const tableRows = document.querySelectorAll('table tbody tr, .table tbody tr, [role="row"]');
    
    tableRows.forEach(row => {
      const checkbox = row.querySelector('input[type="checkbox"]');
      if (checkbox && checkbox.checked) {
        row.style.backgroundColor = '#e3f2fd';
        row.style.transition = 'background-color 0.3s';
      }
    });
  }

  // Inject a visual indicator when extension is active
  if (document.body) {
    const indicator = document.createElement('div');
    indicator.id = 'lightning-outlook-indicator';
    indicator.textContent = '✓ Extension Active';
    indicator.style.cssText = `
      position: fixed;
      top: 10px;
      right: 10px;
      background: #4caf50;
      color: white;
      padding: 8px 12px;
      border-radius: 4px;
      font-size: 12px;
      z-index: 10000;
      font-family: Arial, sans-serif;
      box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    `;
    document.body.appendChild(indicator);
    
    // Remove after 3 seconds
    setTimeout(() => {
      if (indicator.parentNode) {
        indicator.style.opacity = '0';
        indicator.style.transition = 'opacity 0.5s';
        setTimeout(() => indicator.remove(), 500);
      }
    }, 3000);
  }
})();

