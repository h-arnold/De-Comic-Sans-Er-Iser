function onOpen() {
  createMenu()
}

function createMenu() {
  const ui = SlidesApp.getUi();

  ui.createMenu("De-Comic-Sans-er-iser")
  .addItem("Make everything Comic Sans", "installTriggers")
  .addItem("Remove all Comic Sans", "enforceArial")
  .addItem("Remmove all Comic Sans forever!", "enforceArial")
  .addToUi()
}

/**
 * Run this function manually from the Apps Script editor once.
 * It will prompt for authorization and schedule 12 executions,
 * spaced 5 minutes apart.
 */
function installTriggers() {
  // First, clear any existing triggers to avoid duplicates if run multiple times
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  const now = new Date();

  // Create 12 one-off time-based triggers, starting now and every 5 mins after
  for (let i = 0; i < 12; i++) {
    const triggerTime = new Date(now.getTime() + (i * 5 * 60 * 1000));
    ScriptApp.newTrigger('enforceArial')
      .timeBased()
      .at(triggerTime)
      .create();
  }
}

/**
 * Main function to scan the presentation and replace Comic Sans MS with Arial.
 * You can set this function up on a Time-Driven trigger to run every minute.
 */
function enforceArial() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  // Iterate through all slides in the presentation
  slides.forEach(slide => {
    const pageElements = slide.getPageElements();
    processPageElements(pageElements);
  });
  installTriggers();
}

/**
 * Recursively processes page elements to find text and check fonts.
 * This handles standard shapes, tables, and grouped elements.
 * * @param {GoogleAppsScript.Slides.PageElement[]} pageElements 
 */
function processPageElements(pageElements) {
  pageElements.forEach(element => {
    const type = element.getPageElementType();

    if (type === SlidesApp.PageElementType.SHAPE) {
      const shape = element.asShape();
      // Only process shapes that actually support text
      if (shape.getText) {
        checkAndReplaceFont(shape.getText());
      }
    }
    else if (type === SlidesApp.PageElementType.TABLE) {
      const table = element.asTable();
      // Iterate through every cell in the table
      for (let r = 0; r < table.getNumRows(); r++) {
        for (let c = 0; c < table.getNumColumns(); c++) {
          checkAndReplaceFont(table.getCell(r, c).getText());
        }
      }
    }
    else if (type === SlidesApp.PageElementType.GROUP) {
      // If it's a group, recursively process its children
      processPageElements(element.asGroup().getChildren());
    }
  });
}

/**
 * Checks a specific text range and replaces Comic Sans MS with Arial.
 * * @param {GoogleAppsScript.Slides.TextRange} textRange 
 */
function checkAndReplaceFont(textRange) {
  if (!textRange) return;

  // Text ranges are broken up into "runs" based on formatting changes.
  // We check each run individually.
  try {
    const runs = textRange.getRuns();
    runs.forEach(run => {
      const textStyle = run.getTextStyle();
      // Check if the current font is Comic Sans MS
      if (textStyle.getFontFamily() === 'Comic Sans MS') {
        textStyle.setFontFamily('Arial');
      }

    });
  } catch (e) {
    console.error(`Improve your early returns! Error details: ${e.message}`)
  }
}