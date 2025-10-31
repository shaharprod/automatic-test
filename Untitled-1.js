
function _createItem(form, row) {
  const type = (row.Type || '').toString().trim();
  const title = row.Title || '';
  const desc = row.Description || '';
  const optionsRaw = row.Options || '';
  const required = String(row.Required || '').toUpperCase() === 'TRUE';
  const points = row.Points ? Number(row.Points) : null;

  let item = null;
  switch(type) {
    case 'MULTIPLE_CHOICE':
      const mc = form.addMultipleChoiceItem();
      mc.setTitle(title).setHelpText(desc).setRequired(required);
      if (optionsRaw) mc.setChoiceValues(optionsRaw.split(',').map(s=>s.trim()));
      item = mc;
      break;
    case 'CHECKBOX':
      const cb = form.addCheckboxItem();
      cb.setTitle(title).setHelpText(desc).setRequired(required);
      if (optionsRaw) cb.setChoiceValues(optionsRaw.split(',').map(s=>s.trim()));
      item = cb;
      break;
    case 'LIST':
      const list = form.addListItem();
      list.setTitle(title).setHelpText(desc).setRequired(required);
      if (optionsRaw) list.setChoiceValues(optionsRaw.split(',').map(s=>s.trim()));
      item = list;
      break;
    case 'SHORT_ANSWER':
      item = form.addTextItem().setTitle(title).setHelpText(desc).setRequired(required);
      break;
    case 'PARAGRAPH':
      item = form.addParagraphTextItem().setTitle(title).setHelpText(desc).setRequired(required);
      break;
    case 'SCALE':
      // Options expected like "1-5" or "1-10"
      const parts = optionsRaw.split('-').map(p=>p.trim());
      const low = parseInt(parts[0]) || 1;
      const high = parseInt(parts[1]) || 5;
      item = form.addScaleItem().setTitle(title).setHelpText(desc).setBounds(low, high).setRequired(required);
      break;
    case 'DATE':
      item = form.addDateItem().setTitle(title).setHelpText(desc).setRequired(required);
      break;
    case 'TIME':
      item = form.addTimeItem().setTitle(title).setHelpText(desc).setRequired(required);
      break;
    default:
      // Unknown type: add as paragraph
      item = form.addParagraphTextItem().setTitle(title + " (UNKNOWN TYPE: " + type + ")").setHelpText(desc).setRequired(required);
  }

  // If form is a quiz and points provided, set grading (if supported)
  if (points && item && item.setPoints) {
    try { item.setPoints(points); } catch(e) { /* ignore if not supported */ }
  }
  return item;
}

/**
 * Read the active sheet into array of objects keyed by header
 */
function _readSheetAsObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => h.toString().trim());
  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const row = {};
    for (let c = 0; c < headers.length; c++) {
      row[headers[c]] = data[r][c];
    }
    // Stop if blank row (all empty)
    if (Object.values(row).every(v => v === '' || v === null)) continue;
    rows.push(row);
  }
  return rows;
}

/**
 * Main: create a new form from the sheet, or overwrite an existing one
 * It will create the Form in the same Drive folder as the Sheet.
 */
function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const titleForForm = ss.getName() + " - Generated Form";
  const rows = _readSheetAsObjects(sheet);

  if (!rows || rows.length === 0) {
    SpreadsheetApp.getUi().alert("No questions found in the sheet. Fill rows starting from row 2.");
    return;
  }

  // Create form in same folder
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  let folder = DriveApp.getRootFolder();
  if (parents.hasNext()) folder = parents.next();

  const form = FormApp.create(titleForForm);
  // Move the form file into the same folder:
  const formFile = DriveApp.getFileById(form.getId());
  folder.addFile(formFile);
  DriveApp.getRootFolder().removeFile(formFile);

  // Optional: Make it a quiz? Uncomment if wanted:
  // form.setIsQuiz(true);

  // Add questions
  for (let i = 0; i < rows.length; i++) {
    try {
      _createItem(form, rows[i]);
    } catch(e) {
      Logger.log("Error creating item at row " + (i+2) + ": " + e);
    }
  }

  // Show link to form
  const url = form.getPublishedUrl();
  const editUrl = form.getEditUrl();
  SpreadsheetApp.getUi().alert("Form created: " + titleForForm + "\nPublished URL: " + url + "\nEdit URL: " + editUrl);
