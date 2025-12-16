// Content script for extracting schedule data from the schedule page

(function() {
  'use strict';

  // Cache of column indexes by header name
  let columnIndexes = null;

  // Listen for messages from popup
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'getSelectedRows') {
      getSelectedScheduleRows().then(selectedRows => {
        sendResponse({ success: true, data: selectedRows });
      }).catch(error => {
        console.error('Error getting selected rows:', error);
        sendResponse({ success: false, error: error.message });
      });
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
  async function getColumnIndexes() {
    if (columnIndexes) return columnIndexes;

    // Try multiple selectors to get all headers - support both HTML tables and ARIA grids
    // For MUI DataGrid, headers might be in column headers or in the first row
    let headers = document.querySelectorAll('table thead tr th, .table thead tr th, [role="columnheader"]');
    if (headers.length === 0) {
      // Try ARIA grid headers - MUI DataGrid uses columnheader role
      headers = document.querySelectorAll('[role="grid"] [role="columnheader"]');
    }
    if (headers.length === 0) {
      // Try finding first row as header
      const firstRow = document.querySelector('[role="grid"] [role="row"]:first-child, table tr:first-child');
      if (firstRow) {
        headers = firstRow.querySelectorAll('th, td, div, [role="columnheader"], [role="gridcell"]');
      }
    }
    
    // For MUI DataGrid, we might need to get column definitions from the grid itself
    // Check if this is a MUI DataGrid by looking for data-field attributes
    const grid = document.querySelector('[role="grid"]');
    if (grid && headers.length < 10) {
      // Try to get column info from data-field attributes in any visible cells
      const sampleCells = document.querySelectorAll('[role="gridcell"][data-field]');
      const fieldMap = new Map();
      sampleCells.forEach(cell => {
        const field = cell.getAttribute('data-field');
        const colIndex = cell.getAttribute('data-colindex') || cell.getAttribute('aria-colindex');
        if (field && !fieldMap.has(field)) {
          fieldMap.set(field, {
            field: field,
            colIndex: colIndex ? parseInt(colIndex) : null,
            headerText: cell.closest('[role="row"]')?.querySelector(`[data-field="${field}"]`)?.textContent || ''
          });
        }
      });
      
      console.log('Lightning Extension: Found column fields:', Array.from(fieldMap.keys()));
    }
    
    // If we didn't get enough headers, try getting from the first data row or all th elements
    if (headers.length < 10) {
      headers = document.querySelectorAll('table th, .table th, thead th');
    }
    
    // Also try getting from tbody first row if it's a header row
    const firstRow = document.querySelector('table tbody tr:first-child, .table tbody tr:first-child');
    if (firstRow && firstRow.querySelectorAll('th').length > headers.length) {
      headers = firstRow.querySelectorAll('th');
    }
    
    console.log(`Lightning Extension: Found ${headers.length} header columns`);
    
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

    // Also check for MUI DataGrid column fields by looking at data-field attributes
    const fieldToColumnMap = {};
    const gridCellsWithFields = document.querySelectorAll('[role="gridcell"][data-field]');
    gridCellsWithFields.forEach(cell => {
      const field = cell.getAttribute('data-field');
      const colIndex = cell.getAttribute('data-colindex') || cell.getAttribute('aria-colindex');
      if (field && colIndex && !fieldToColumnMap[field]) {
        fieldToColumnMap[field] = parseInt(colIndex);
        console.log(`Lightning Extension: Found field "${field}" at column index ${colIndex}`);
      }
    });
    
    // First, try to detect columns by data-field attributes from any visible cells
    // This works better for MUI DataGrid which uses field-based columns
    const sampleCells = document.querySelectorAll('[role="gridcell"][data-field]');
    sampleCells.forEach(cell => {
      const field = cell.getAttribute('data-field');
      const colIndex = cell.getAttribute('data-colindex') || cell.getAttribute('aria-colindex');
      
      if (field) {
        const fieldLower = field.toLowerCase();
        const idx = colIndex ? parseInt(colIndex) : -1;
        
        // Map field names to our column map
        if (fieldLower === 'jobnumber' || fieldLower === 'job#' || fieldLower.includes('jobnumber')) {
          map.jobNumber = idx >= 0 ? idx : map.jobNumber;
        } else if (fieldLower === 'jobname' || fieldLower.includes('jobname')) {
          map.jobName = idx >= 0 ? idx : map.jobName;
        } else if (fieldLower === 'client') {
          map.client = idx >= 0 ? idx : map.client;
        } else if (fieldLower === 'venuename' || fieldLower.includes('venuename')) {
          map.venueName = idx >= 0 ? idx : map.venueName;
        } else if (fieldLower === 'venueroom' || fieldLower.includes('venueroom')) {
          map.venueRoom = idx >= 0 ? idx : map.venueRoom;
        } else if (fieldLower === 'address') {
          map.address = idx >= 0 ? idx : map.address;
        } else if (fieldLower === 'salesperson' || fieldLower.includes('salesperson')) {
          map.salesperson = idx >= 0 ? idx : map.salesperson;
        } else if (fieldLower === 'orderstatus' || fieldLower.includes('orderstatus')) {
          map.orderStatus = idx >= 0 ? idx : map.orderStatus;
        } else if (fieldLower === 'status') {
          map.status = idx >= 0 ? idx : map.status;
        } else if (fieldLower === 'startdate' || fieldLower.includes('startdate')) {
          map.startDate = idx >= 0 ? idx : map.startDate;
        } else if (fieldLower === 'enddate' || fieldLower.includes('enddate')) {
          map.endDate = idx >= 0 ? idx : map.endDate;
        } else if (fieldLower === 'talent') {
          map.talent = idx >= 0 ? idx : map.talent;
        } else if (fieldLower === 'task') {
          map.task = idx >= 0 ? idx : map.task;
        }
      }
    });
    
    headers.forEach((th, idx) => {
      const text = (th.textContent || '').trim().toLowerCase().replace(/\s+/g, ' ');
      const dataField = th.getAttribute('data-field');
      const colIndex = th.getAttribute('data-colindex') || th.getAttribute('aria-colindex');
      
      // Debug: log all headers to help identify column detection issues
      if (idx === 0) {
        console.log('Lightning Extension: Detecting columns from headers...');
      }
      console.log(`  Column ${idx}: "${text}" (raw: "${th.textContent}", field: ${dataField || 'none'}, colIndex: ${colIndex || 'none'})`);

      // Use data-field if available (MUI DataGrid) - only if not already set
      if (dataField) {
        const fieldLower = dataField.toLowerCase();
        const idxNum = colIndex ? parseInt(colIndex) : idx;
        if (fieldLower === 'jobnumber' && map.jobNumber === -1) {
          map.jobNumber = idxNum;
        } else if (fieldLower === 'jobname' && map.jobName === -1) {
          map.jobName = idxNum;
        } else if (fieldLower === 'client' && map.client === -1) {
          map.client = idxNum;
        } else if (fieldLower === 'venuename' && map.venueName === -1) {
          map.venueName = idxNum;
        } else if (fieldLower === 'venueroom' && map.venueRoom === -1) {
          map.venueRoom = idxNum;
        } else if (fieldLower === 'address' && map.address === -1) {
          map.address = idxNum;
        } else if (fieldLower === 'salesperson' && map.salesperson === -1) {
          map.salesperson = idxNum;
        } else if (fieldLower === 'orderstatus' && map.orderStatus === -1) {
          map.orderStatus = idxNum;
        }
      }

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
      } else if (text.includes('client')) {
        map.client = idx;
      } else if (text.includes('venue name') || (text.includes('venue') && text.includes('name'))) {
        map.venueName = idx;
      } else if (text.includes('venue room') || (text.includes('venue') && text.includes('room'))) {
        map.venueRoom = idx;
      } else if (text.includes('address')) {
        map.address = idx;
      } else if (text.includes('salesperson') || text.includes('sales person')) {
        map.salesperson = idx;
      } else if (text.includes('order status') || (text.includes('order') && text.includes('status'))) {
        map.orderStatus = idx;
      } else if (text.includes('labor custom') || (text.includes('labor') && text.includes('custom'))) {
        map.laborCustom = idx;
      }
    });

    // Check data rows to see how many columns they actually have
    // This helps us detect columns that don't have headers
    // Try multiple selectors to find data rows (support both tables and ARIA grids)
    let sampleRows = document.querySelectorAll('table tbody tr');
    if (sampleRows.length === 0) {
      sampleRows = document.querySelectorAll('[role="grid"] [role="row"]:not(:first-child)');
    }
    if (sampleRows.length === 0) {
      sampleRows = document.querySelectorAll('table tr:not(:first-child)');
    }
    if (sampleRows.length === 0) {
      sampleRows = document.querySelectorAll('[role="row"]:not(:first-child)');
    }
    if (sampleRows.length === 0) {
      const allRows = document.querySelectorAll('table tr, [role="row"]');
      sampleRows = Array.from(allRows).filter((row, idx) => idx > 0);
    }
    let actualColumnCount = 0;
    
    if (sampleRows.length > 0) {
      const firstDataRow = sampleRows[0];
      // Support both table cells and ARIA grid cells
      // For ARIA grids, gridcells are the actual data cells
      let cells = firstDataRow.querySelectorAll('[role="gridcell"]');
      if (cells.length === 0) {
        cells = firstDataRow.querySelectorAll('td');
      }
      if (cells.length === 0) {
        // Fallback: try direct children
        cells = Array.from(firstDataRow.children).filter(child => {
          const role = child.getAttribute('role');
          return role === 'gridcell' || !role || child.tagName === 'TD';
        });
      }
      actualColumnCount = cells.length;
      console.log(`Lightning Extension: Data rows have ${actualColumnCount} columns, but only ${headers.length} headers found`);
      
      // For MUI DataGrid, we need to scroll horizontally to reveal virtualized columns
      // Try to find and scroll the grid container to reveal all columns
      const grid = document.querySelector('[role="grid"]');
      if (grid && actualColumnCount < 16) {
        console.log('Lightning Extension: Grid has fewer columns than expected, attempting to reveal virtualized columns...');
        
        // Try multiple selectors to find the scrollable container
        // MUI DataGrid uses different structures, so we need to try various approaches
        let scrollContainer = null;
        const possibleContainers = [
          grid.closest('.MuiDataGrid-virtualScroller'),
          grid.closest('.MuiDataGrid-root'),
          grid.querySelector('.MuiDataGrid-virtualScroller'),
          grid.querySelector('[class*="virtualScroller"]'),
          grid.querySelector('[class*="VirtualScroller"]'),
          document.querySelector('.MuiDataGrid-virtualScroller'),
          document.querySelector('[class*="virtualScroller"]'),
          grid.parentElement,
          grid
        ];
        
        for (const container of possibleContainers) {
          if (container && container.scrollWidth > container.clientWidth) {
            scrollContainer = container;
            console.log(`Lightning Extension: Found scrollable container: ${container.className || container.tagName}`);
            break;
          }
        }
        
        if (scrollContainer) {
          // Scroll to the right to reveal more columns
          const originalScrollLeft = scrollContainer.scrollLeft || 0;
          const maxScroll = scrollContainer.scrollWidth - scrollContainer.clientWidth;
          
          console.log(`Lightning Extension: Scroll container dimensions - scrollWidth: ${scrollContainer.scrollWidth}, clientWidth: ${scrollContainer.clientWidth}, maxScroll: ${maxScroll}`);
          
          if (maxScroll > 0) {
            console.log(`Lightning Extension: Scrolling grid horizontally (max scroll: ${maxScroll})`);
            scrollContainer.scrollLeft = maxScroll;
            
            // Wait for columns to render
            await new Promise(resolve => setTimeout(resolve, 1000));
            
            // Re-query the first data row after scrolling (it might have changed)
            const newSampleRows = document.querySelectorAll('[role="grid"] [role="row"]:not(:first-child)');
            const newFirstDataRow = newSampleRows.length > 0 ? newSampleRows[0] : firstDataRow;
            
            // Check again for gridcells after scrolling
            const gridcellsAfterScroll = newFirstDataRow.querySelectorAll('[role="gridcell"]');
            console.log(`Lightning Extension: After scrolling, found ${gridcellsAfterScroll.length} gridcells (was ${actualColumnCount})`);
            
            // Log all fields found after scrolling for debugging
            const fieldsAfterScroll = [];
            gridcellsAfterScroll.forEach(cell => {
              const field = cell.getAttribute('data-field');
              if (field) fieldsAfterScroll.push(field);
            });
            console.log(`Lightning Extension: All fields found after scrolling: ${fieldsAfterScroll.join(', ')}`);
            
            if (gridcellsAfterScroll.length > actualColumnCount) {
              actualColumnCount = gridcellsAfterScroll.length;
              
              // Rebuild field map with newly visible cells
              gridcellsAfterScroll.forEach(cell => {
                const field = cell.getAttribute('data-field');
                const colIndex = cell.getAttribute('data-colindex') || cell.getAttribute('aria-colindex');
                if (field && colIndex) {
                  const idx = parseInt(colIndex);
                  const fieldLower = field.toLowerCase();
                  
                  if (fieldLower === 'jobnumber' || fieldLower.includes('jobnumber')) {
                    map.jobNumber = idx;
                    console.log(`Lightning Extension: Found jobNumber at column ${idx}`);
                  } else if (fieldLower === 'jobname' || fieldLower.includes('jobname')) {
                    map.jobName = idx;
                    console.log(`Lightning Extension: Found jobName at column ${idx}`);
                  } else if (fieldLower === 'clientname' || fieldLower === 'client') {
                    map.client = idx;
                    console.log(`Lightning Extension: Found client at column ${idx}`);
                  } else if (fieldLower === 'venuename' || fieldLower.includes('venuename')) {
                    map.venueName = idx;
                    console.log(`Lightning Extension: Found venueName at column ${idx}`);
                  } else if (fieldLower === 'venueroom' || fieldLower.includes('venueroom')) {
                    map.venueRoom = idx;
                    console.log(`Lightning Extension: Found venueRoom at column ${idx}`);
                  } else if (fieldLower === 'address') {
                    map.address = idx;
                    console.log(`Lightning Extension: Found address at column ${idx}`);
                  } else if (fieldLower === 'salesperson' || fieldLower.includes('salesperson')) {
                    map.salesperson = idx;
                    console.log(`Lightning Extension: Found salesperson at column ${idx}`);
                  } else if (fieldLower === 'orderstatus' || fieldLower.includes('orderstatus') || fieldLower === 'order status') {
                    map.orderStatus = idx;
                    console.log(`Lightning Extension: Found orderStatus at column ${idx}`);
                  } else if (fieldLower === 'laborstatus' && map.status === -1) {
                    // laborStatus might be used for order status
                    map.status = idx;
                    console.log(`Lightning Extension: Found status (laborStatus) at column ${idx}`);
                  }
                  
                  // Log all fields found for debugging
                  console.log(`Lightning Extension: Found field "${field}" at column ${idx}`);
                }
              });
            }
            
            // Scroll back to original position
            scrollContainer.scrollLeft = originalScrollLeft;
            await new Promise(resolve => setTimeout(resolve, 300));
          } else {
            console.log('Lightning Extension: Grid does not appear to be horizontally scrollable (maxScroll <= 0)');
          }
        } else {
          console.log('Lightning Extension: Could not find scrollable container');
        }
      }
    }
    
    columnIndexes = map;
    
    // Debug: log detected column mappings
    console.log('Lightning Extension: Detected column mappings:', map);
    console.log('Lightning Extension: Address column index:', map.address);
    console.log('Lightning Extension: Salesperson column index:', map.salesperson);
    console.log('Lightning Extension: Order Status column index:', map.orderStatus);
    
    return map;
  }

  /**
   * Extract schedule data from selected table rows
   */
  async function getSelectedScheduleRows() {
    const rows = [];
    
    // Find the schedule rows - support both HTML tables and ARIA grids
    // Looking for rows with checkboxes
    let tableRows = document.querySelectorAll('table tbody tr');
    if (tableRows.length === 0) {
      // Try ARIA grid rows (skip first row which is usually header)
      tableRows = document.querySelectorAll('[role="grid"] [role="row"]:not(:first-child)');
    }
    if (tableRows.length === 0) {
      tableRows = document.querySelectorAll('table tr:not(:first-child)');
    }
    if (tableRows.length === 0) {
      // Try all rows and filter
      const allRows = document.querySelectorAll('[role="row"], table tr');
      tableRows = Array.from(allRows).filter((row, idx) => {
        // Skip first row (likely header) and rows without checkboxes
        const hasCheckbox = row.querySelector('input[type="checkbox"]');
        return idx > 0 && hasCheckbox;
      });
    }
    
    console.log(`Lightning Extension: Found ${tableRows.length} potential data rows`);
    
    // Find scrollable container to save original position
    const grid = document.querySelector('[role="grid"]');
    let scrollContainer = null;
    let originalScrollLeft = 0;
    
    if (grid) {
      const possibleContainers = [
        grid.closest('.MuiDataGrid-virtualScroller'),
        grid.closest('.MuiDataGrid-root'),
        grid.querySelector('.MuiDataGrid-virtualScroller'),
        document.querySelector('.MuiDataGrid-virtualScroller'),
        grid.parentElement,
        grid
      ];
      
      for (const container of possibleContainers) {
        if (container && container.scrollWidth > container.clientWidth) {
          scrollContainer = container;
          originalScrollLeft = container.scrollLeft || 0;
          break;
        }
      }
    }
    
    // Process rows asynchronously - only extract rows with checked checkboxes
    for (let index = 0; index < tableRows.length; index++) {
      const row = tableRows[index];
      // Check if row has a checked checkbox
      // Try multiple selectors for checkboxes (MUI DataGrid might use different structures)
      let checkbox = row.querySelector('input[type="checkbox"]');
      if (!checkbox) {
        // Try finding checkbox in parent or sibling elements
        checkbox = row.closest('[role="row"]')?.querySelector('input[type="checkbox"]');
      }
      if (!checkbox) {
        // Try aria-checked attribute
        const checkboxElement = row.querySelector('[role="checkbox"]');
        if (checkboxElement && checkboxElement.getAttribute('aria-checked') === 'true') {
          checkbox = checkboxElement;
        }
      }
      
      // Check if checkbox is checked
      const isChecked = checkbox && (
        (checkbox.checked === true) || 
        (checkbox.getAttribute('aria-checked') === 'true') ||
        (row.getAttribute('aria-selected') === 'true')
      );
      
      if (isChecked) {
        console.log(`Lightning Extension: Row ${index + 1} is selected, extracting data...`);
        // Each row will scroll independently to collect all its columns
        // Pass false so we restore scroll position after all rows are processed
        const rowData = await extractRowData(row, index, false);
        if (rowData) {
          console.log(`Lightning Extension: Successfully extracted data from row ${index + 1}:`, {
            jobNumber: rowData.jobNumber,
            jobName: rowData.jobName,
            startDate: rowData.startDate,
            address: rowData.address
          });
          rows.push(rowData);
        } else {
          console.log(`Lightning Extension: Failed to extract data from row ${index + 1} - extractRowData returned null`);
        }
      }
    }

    // Only process selected rows - no fallback to all rows
    if (rows.length === 0) {
      console.log('Lightning Extension: No rows selected. Please check the boxes next to the rows you want to sync.');
    } else {
      console.log(`Lightning Extension: Found ${rows.length} selected row(s) to process`);
    }
    
    // Restore original scroll position after all rows are processed
    if (scrollContainer) {
      scrollContainer.scrollLeft = originalScrollLeft;
      await new Promise(resolve => setTimeout(resolve, 300));
    }

    return rows;
  }

  /**
   * Extract data from a single table row
   * Updated to use the same successful logic as the test script
   */
  async function extractRowData(row, index, shouldScroll = true) {
    try {
      // Support both table cells and ARIA grid cells
      // For MUI DataGrid, cells are identified by data-field attribute
      const cellsByField = {};
      const cellsByIndex = [];
      let originalScrollLeft = 0;
      let scrollContainer = null;
      
      // Find scrollable container
      const grid = document.querySelector('[role="grid"]');
      if (grid) {
        const possibleContainers = [
          grid.closest('.MuiDataGrid-virtualScroller'),
          grid.closest('.MuiDataGrid-root'),
          grid.querySelector('.MuiDataGrid-virtualScroller'),
          document.querySelector('.MuiDataGrid-virtualScroller'),
          grid.parentElement,
          grid
        ];
        
        for (const container of possibleContainers) {
          if (container && container.scrollWidth > container.clientWidth) {
            scrollContainer = container;
            break;
          }
        }
      }
      
      // Always scroll through all positions to collect all columns for this row
      // The shouldScroll parameter only controls whether we restore the scroll position
      if (scrollContainer) {
        originalScrollLeft = scrollContainer.scrollLeft || 0;
        const maxScroll = scrollContainer.scrollWidth - scrollContainer.clientWidth;
        
        // Collect from left position
        scrollContainer.scrollLeft = 0;
        await new Promise(resolve => setTimeout(resolve, 600));
        
        const leftCells = row.querySelectorAll('[role="gridcell"]');
        leftCells.forEach(cell => {
          const field = cell.getAttribute('data-field');
          const colIndex = cell.getAttribute('data-colindex') || cell.getAttribute('aria-colindex');
          const text = cell.textContent.trim();
          
          if (field) {
            cellsByField[field] = text;
          }
          if (colIndex !== null) {
            cellsByIndex[parseInt(colIndex)] = { text, field };
          }
        });
        
        // Collect from right position
        if (maxScroll > 0) {
          scrollContainer.scrollLeft = maxScroll;
          await new Promise(resolve => setTimeout(resolve, 600));
          
          const rightCells = row.querySelectorAll('[role="gridcell"]');
          rightCells.forEach(cell => {
            const field = cell.getAttribute('data-field');
            const colIndex = cell.getAttribute('data-colindex') || cell.getAttribute('aria-colindex');
            const text = cell.textContent.trim();
            
            if (field) {
              cellsByField[field] = text; // Overwrite with right position data
            }
            if (colIndex !== null) {
              cellsByIndex[parseInt(colIndex)] = { text, field };
            }
          });
        }
        
        // Also try scrolling to multiple positions to catch all columns
        // Some columns might be in the middle or require specific scroll positions
        if (maxScroll > 0) {
          // Try a few intermediate positions
          const positions = [maxScroll * 0.25, maxScroll * 0.5, maxScroll * 0.75, maxScroll];
          for (const pos of positions) {
            scrollContainer.scrollLeft = pos;
            await new Promise(resolve => setTimeout(resolve, 400));
            
            const midCells = row.querySelectorAll('[role="gridcell"]');
            midCells.forEach(cell => {
              const field = cell.getAttribute('data-field');
              const colIndex = cell.getAttribute('data-colindex') || cell.getAttribute('aria-colindex');
              const text = cell.textContent.trim();
              
              if (field && (!cellsByField[field] || cellsByField[field] === '')) {
                // Only update if we don't have this field yet or it's empty
                if (text) {
                  cellsByField[field] = text;
                }
              }
              if (colIndex !== null) {
                const idx = parseInt(colIndex);
                if (!cellsByIndex[idx] || !cellsByIndex[idx].text) {
                  cellsByIndex[idx] = { text, field };
                }
              }
            });
          }
        }
        
        // Restore scroll position only if shouldScroll is true
        // If false, we'll restore it in the parent function after all rows are processed
        if (shouldScroll) {
          scrollContainer.scrollLeft = originalScrollLeft;
          await new Promise(resolve => setTimeout(resolve, 300));
        }
      } else {
        // No scrolling needed - collect from current position
        const cells = row.querySelectorAll('[role="gridcell"], td');
        cells.forEach((cell, idx) => {
          const field = cell.getAttribute('data-field');
          const text = cell.textContent.trim();
          if (field) {
            cellsByField[field] = text;
          }
          cellsByIndex[idx] = { text, field };
        });
      }
      
      console.log(`Lightning Extension: Row ${index} extracted ${Object.keys(cellsByField).length} fields: ${Object.keys(cellsByField).sort().join(', ')}`);
      
      // Helper function to get text by field name
      const getByField = (fieldName) => {
        return cellsByField[fieldName] || '';
      };

      // Extract data using field-based lookup (same approach as test script)
      const rawStartDate = getByField('startDate');
      const rawEndDate = getByField('endDate');
      
      const data = {
        index: index,
        type: getByField('type') || '',
        refNumber: getByField('refNumber') || '',
        name: getByField('name') || '',
        description: getByField('description') || '',
        startDate: parseDate(rawStartDate),
        endDate: parseDate(rawEndDate),
        office: getByField('office') || '',
        projectNumber: getByField('projectNumber') || '',
        // Extended fields for laborSchedule token page
        status: getByField('laborStatus') || getByField('status') || '',
        confirm: getByField('confirm') || '',
        deny: getByField('deny') || '',
        talent: getByField('talent') || '',
        task: getByField('task') || '',
        jobNumber: getByField('jobNumber') || '',
        jobName: getByField('jobName') || '',
        client: getByField('clientName') || getByField('client') || '',
        venueName: getByField('venueName') || '',
        venueRoom: getByField('venueRoom') || '',
        address: getByField('address') || '',
        salesperson: getByField('salesperson') || '',
        orderStatus: getByField('orderStatus') || '',
        laborCustom: getByField('laborCustom') || getByField('laborcustom') || getByField('labor_custom') || getByField('LaborCustom') || ''
      };
      
      // Debug: log extracted data for first few rows
      if (index < 3) {
        console.log(`Lightning Extension: Row ${index} extracted data:`, {
          jobNumber: data.jobNumber,
          jobName: data.jobName,
          startDate: rawStartDate,
          endDate: rawEndDate,
          address: data.address,
          salesperson: data.salesperson,
          orderStatus: data.orderStatus,
          laborCustom: data.laborCustom
        });
      }

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
        console.log(`Lightning Extension: Row ${index} has invalid dates - startDate: ${data.startDate}, endDate: ${data.endDate}`);
        console.log(`Lightning Extension: Raw startDate value: "${rawStartDate}", Raw endDate value: "${rawEndDate}"`);
        return null;
      }

      // Validate that we have at least a start date before returning
      if (!data.startDate) {
        console.log(`Lightning Extension: Row ${index} missing startDate, returning null`);
        return null;
      }
      
      console.log(`Lightning Extension: Successfully extracted row ${index} data with ${Object.keys(data).length} fields`);
      return data;
    } catch (error) {
      console.error(`Lightning Extension: Error extracting row ${index} data:`, error);
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

