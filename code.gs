// This function creates a user interface for the add-on
function onOpen(e) {
  // Creates a menu item in the UI that shows when the add-on is opened
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem("Format selected text", "showSidebar")
    .addToUi();
}

// This function shows a sidebar with the formatting options
function showSidebar() {
  // Creates an HTML sidebar using the sidebar.html file
  var html = HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("Bionic Reader Add-on")
    .setWidth(300);
  // Displays the sidebar in the UI
  DocumentApp.getUi().showSidebar(html);
}

// This function formats the selected text with bold first letters of each word
function formatSelectedText() {
  // Gets the current selection in the document
  var selection = DocumentApp.getActiveDocument().getSelection();
  
  // Gets the elements within the selection
  var elements = selection.getRangeElements();
  
  // Iterates over each element in the selection
  for (var i = 0; i < elements.length; i++) {
  
    // Gets the current element
    var element = elements[i];
    
    // Checks if the element is editable text
    if (element.getElement().editAsText) {
    
      // Gets the text element of the current element
      var text = element.getElement().editAsText();
      
      // Gets the start and end offsets of the selection within the text element
      var startOffset = element.getStartOffset();
      var endOffset = element.getEndOffsetInclusive();
      
      // Gets the selected text within the text element
      var selectedText = text.getText().substring(startOffset, endOffset + 1);
      
      // Splits the selected text into an array of words
      var words = selectedText.split(" ");
      
      // Iterates over each word in the array
      for (var j = 0; j < words.length; j++) {
      
        // Gets the current word
        var word = words[j];
        
        // Checks if the word is not empty
        if (word.length > 0) {
        
          // Gets the first letter of the word and the rest of the word
          var firstLetter = word.charAt(0);
          var restOfWord = word.substring(1);
          
          // Sets the bold formatting for the first letter of the word
          text.setBold(startOffset + selectedText.indexOf(word), startOffset + selectedText.indexOf(word) + firstLetter.length, true);
        }
      }
    }
  }
}

// This function is called when the user clicks the "Format" button in the sidebar
function format() {
  // Calls the formatSelectedText function in the server-side script
  google.script.run.formatSelectedText();
  // Sets the status message to indicate that the text has been formatted
  document.getElementById("status").innerHTML = "Selected text formatted!";
}

// This is the HTML template for the sidebar
function doGet() {
  // Creates an HTML template from the sidebar.html file
  var template = HtmlService.createTemplateFromFile("sidebar");
  // Evaluates the HTML template
  return template.evaluate();
}

