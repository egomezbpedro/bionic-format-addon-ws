// This function creates a user interface for the add-on
function onOpen(e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem("Format selected text", "showSidebar")
    .addToUi();
}

// This function shows a sidebar with the formatting options
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("Bionic Reader Add-on")
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

// This function formats the selected text with bold first letters of each word
function formatSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  var elements = selection.getRangeElements();
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    if (element.getElement().editAsText) {
      var text = element.getElement().editAsText();
      var startOffset = element.getStartOffset();
      var endOffset = element.getEndOffsetInclusive();
      var selectedText = text.getText().substring(startOffset, endOffset + 1);
      var words = selectedText.split(" ");
      for (var j = 0; j < words.length; j++) {
        var word = words[j];
        if (word.length > 0) {
          var firstLetter = word.charAt(0);
          var restOfWord = word.substring(1);
          text.setBold(startOffset + j, startOffset + j + 1, true);
        }
      }
    }
  }
}

// This function is called when the user clicks the "Format" button in the sidebar
function format() {
  google.script.run.formatSelectedText();
  document.getElementById("status").innerHTML = "Selected text formatted!";
}

// This is the HTML template for the sidebar
function doGet() {
  var template = HtmlService.createTemplateFromFile("sidebar");
  return template.evaluate();
}

// This is the CSS stylesheet for the sidebar
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}