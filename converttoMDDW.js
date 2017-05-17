//function myFunction() {
// 
//}

// ----------------------------------------------------------------------------
// The following script based on gdocs2md which are available at the follows
//   Original: https://github.com/mangini/gdocs2md
//   Modified: https://github.com/sgregson/gdocs2md
// ----------------------------------------------------------------------------
// Open handler to add Menu
function onOpen(e) {
  var ui = DocumentApp.getUi();
  
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    // This script doesn't work in AuthMode.NONE
    return;
    //ui.createMenu('ExportDokuwiki')
    //  .addToUi();
  } else {
    ui.createMenu('ExportDoc')
      .addItem('Email with Dokuwiki format', 'ConvertToDokuwikiEmail')
      .addItem('Email with Markdown format', 'ConvertToMarkdownEmail')
      .addToUi();
  }
}

function onInstall(e) {
  onOpen(e);
}

// Convert current document to Markdown format and email it 
function ConvertToMarkdownEmail() {
  // Convert to markdown
  var convertedDoc = convertMDDW(0); // 0 -> Convert to Markdown

  // Add markdown document to attachments
  convertedDoc.attachments.push({"fileName":DocumentApp.getActiveDocument().getName()+".md",
                                 "mimeType": "text/plain", "content": convertedDoc.text});

  // In some cases user email is not accessible
  var mail = Session.getActiveUser().getEmail();
  if(mail === '') {
    DocumentApp.getUi().alert("Could not read your email address");
    return;
  }

  // Send email with markdown document
  MailApp.sendEmail(mail,
                    "[MARKDOWN_MAKER] "+DocumentApp.getActiveDocument().getName(),
                    "Your converted markdown document is attached (converted from "+DocumentApp.getActiveDocument().getUrl()+")"+
                    "\n\nDon't know how to use the format options? See http://github.com/mangini/gdocs2md\n",
                    { "attachments": convertedDoc.attachments });
}

// Convert current document to Dokuwiki format and email it 
function ConvertToDokuwikiEmail() {
  // Convert to dokuwiki
  var convertedDoc = convertMDDW(1); // 1 -> Convert to Dokuwiki
  
  // Add dokuwiki formatted text file to attachments
  convertedDoc.attachments.push({"fileName":DocumentApp.getActiveDocument().getName()+".txt", 
                                 "mimeType": "text/plain", "content": convertedDoc.text});

  // In some cases user email is not accessible  
  var mail = Session.getActiveUser().getEmail(); 
  if(mail === '') {
    DocumentApp.getUi().alert("Could not read your email address"); 
    return;
  }
  
  // Send email with markdown document
  MailApp.sendEmail(mail,
					"[DOKUWIKI_MAKER] "+DocumentApp.getActiveDocument().getName(),
					"Your converted Dokuwiki document is attached ( converted from "+DocumentApp.getActiveDocument().getUrl()+" )"+
					"\n\nPlase copy and paste the document to Dokuwiki.\n",
					{ "attachments": convertedDoc.attachments });
}

// CMode: 0 -> MD, 1-> DW
function processSection(section, cmode) {
  var state = {
    'inSource' : false, // Document read pointer is within a fenced code block
    'images' : [], // Image data found in document
    'imageCounter' : 0, // Image counter 
    'prevDoc' : [], // Pointer to the previous element on aparsing tree level
    'nextDoc' : [], // Pointer to the next element on a parsing tree level
    'size' : [], // Number of elements on a parsing tree level
    'listCounters' : [], // List counter
  };
  
  // Process element tree outgoing from the root element
  var textElements = [];
  
  if(cmode == 0){
    textElements = processElementMD(section, state, 0);
  }else{
    textElements = processElementDW(section, state, 0);
  }
  return {
    'textElements' : textElements,
    'state' : state,
  }; 
}

function convertMDDW(cmode) {
  // Text elements
  var textElements = []; 
  
  // Process body only
  var doc = DocumentApp.getActiveDocument().getBody();
  doc = processSection(doc, cmode);
  textElements = textElements.concat(doc.textElements); 
  
  // Build final output string
  var text = textElements.join('');
  
  // Replace critical chars
  text = text.replace('\u201d', '"').replace('\u201c', '"');
  
  // Debug logging
  Logger.log("Result: " + text);
  Logger.log("Images: " + doc.state.imageCounter);
  
  // Build attachment and file lists
  var attachments = [];
  var files = [];
  for(var i in doc.state.images) {
    var image = doc.state.images[i];
    attachments.push( {
      "fileName": image.name,
      "mimeType": image.type,
      "content": image.bytes
    } );
    
    files.push( {
      "name" : image.name,
      "blob" : image.blob
    });
  }
  
  // Results
  return {
    'files' : files,
    'attachments' : attachments,
    'text' : text,
  };
}

function escapeHTML(text) {
  return text.replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Add repeat function to strings
String.prototype.repeat = function( num ) {
  return new Array( num + 1 ).join( this );
}

// FIX This:
function handleTable(element, state, depth, cmode) {
  var textElements = [];
  
  textElements.push("\n");
  
  function buildTable(size) {
    var stack = []
    var maxSize = 0; 
    
    for(var ir=0; ir<element.getNumRows(); ir++) {
      var row = element.getRow(ir);
      
      // Add header seperator
      if(ir == 1) {
        for(var ic=0; ic<row.getNumCells(); ic++) {
          stack.push("|-" + "-".repeat(size));
        }
        stack.push("-|\n");
      }
      
      // Add table data
      for(var ic=0; ic<row.getNumCells(); ic++) {
        var cell = row.getCell(ic);
        
        // Recursively build cell content
        var text = processChilds(cell, state, depth+1, cmode).join('');
        
        text = text.replace(/(\r\n|\n|\r)/gm,"");
        maxSize = Math.max(text.length, maxSize); 
        
        if(size > text.length) {
          text += " ".repeat(size - text.length)
        }
        
        stack.push("| " + text);
      }
      
      stack.push(" |\n");
    }
    
    stack.push("\n");
    return {
      maxSize : maxSize,
      stack : stack,
    };
  }
  
  var table = buildTable(100); 
  table = buildTable(Math.max(10, table.maxSize + 1)); 
  textElements = textElements.concat(table.stack);
  
  textElements.push('\n');
  return textElements;
}

function formatDw(text, indexLeft, formatLeft, indexRight, formatRight) {
  var leftPad = '' + formatLeft; 
  if(indexLeft > 0) {
    if(text[indexLeft - 1] != ' ')
      leftPad = ' ' + formatLeft; 
  }
  
  var rightPad = formatRight + '';
  if(indexRight < text.length) {
    if(text[indexRight] != ' ') {
      rightPad = formatRight + ' ';
    }
  }
  
  var formatted = text.substring(0, indexLeft) + leftPad + text.substring(indexLeft, indexRight) + rightPad + text.substring(indexRight);
  return formatted;
}

function formatMd(text, indexLeft, formatLeft, indexRight, formatRight) {
  var leftPad = '' + formatLeft;
  if(indexLeft > 0) {
    if(text[indexLeft - 1] != ' ')
      leftPad = ' ' + formatLeft;
  }

  var rightPad = formatRight + '';
  if(indexRight < text.length) {
    if(text[indexRight] != ' ') {
      rightPad = formatRight + ' ';
    }
  }

  var formatted = text.substring(0, indexLeft) + leftPad + text.substring(indexLeft, indexRight) + rightPad + text.substring(indexRight);
  return formatted;
}

function handleTextDW(doc, state) {
  var formatted = doc.getText(); 
  var lastIndex = formatted.length; 
  var attrs = doc.getTextAttributeIndices();
  
  // Iterate backwards through all attributes
  for(var i=attrs.length-1; i >= 0; i--) {
    // Current position in text
    var index = attrs[i];
    var dleft, right;
        
    // Handle links
    if(doc.getLinkUrl(index)) {
      var url = doc.getLinkUrl(index);
      if (i > 0 && attrs[i-1] == index - 1 && doc.getLinkUrl(attrs[i-1]) === url) {
        i -= 1;
        index = attrs[i];
        url = txt.getLinkUrl(off);
      }
      formatted = formatted.substring(0, index) + '[[' + url + '|' + formatted.substring(index, lastIndex) + ']]' + formatted.substring(lastIndex);
      
      // Do not handle additional formattings for links
      continue; 
    } 
    
    // Handle font family
    if(doc.getFontFamily(index)) {
      var font = doc.getFontFamily(index); 
      var sourceFont = "Courier New"; 
      
      if (!state.inSource && font === sourceFont) {
        // Scan left until text without source font is found
        while (i > 0 && doc.getFontFamily(attrs[i-1]) && doc.getFontFamily(attrs[i-1]) === sourceFont) {
          i -= 1;
          off = attrs[i];
        }
        
        formatted = formatDw(formatted, index, '`', lastIndex, '`');
        
        // Do not handle additional formattings for code
        continue; 
      }
    }
    
    // Handle bold and bold italic
    dleft = dright = '';
    if(doc.isBold(index)) {
      dleft  += '**';
      dright = '**' + dright;
    }
    if (doc.isUnderline(index)) {
      dleft  += '__'; 
      dright = '__' + dright;
    }
    if (doc.isItalic(index)) {
      dleft  += '//'; 
      dright = '//' + dright;
    }
      
    if (dleft.length > 0) {
      formatted = formatDw(formatted, index, dleft, lastIndex, dright); 
    }
    
    // Keep track of last position in text
    lastIndex = index; 
  }
  
  var textElements = [formatted]; 
  return textElements; 
}

function handleTextMD(doc, state) {
  var formatted = doc.getText();
  var lastIndex = formatted.length;
  var attrs = doc.getTextAttributeIndices();

  // Iterate backwards through all attributes
  for(var i=attrs.length-1; i >= 0; i--) {
    // Current position in text
    var index = attrs[i];

    // Handle links
    if(doc.getLinkUrl(index)) {
      var url = doc.getLinkUrl(index);
      if (i > 0 && attrs[i-1] == index - 1 && doc.getLinkUrl(attrs[i-1]) === url) {
        i -= 1;
        index = attrs[i];
        url = txt.getLinkUrl(off);
      }
      formatted = formatted.substring(0, index) + '[' + formatted.substring(index, lastIndex) + '](' + url + ')' + formatted.substring(lastIndex);

      // Do not handle additional formattings for links
      continue;
    }

    // Handle font family
    if(doc.getFontFamily(index)) {
      var font = doc.getFontFamily(index);
      var sourceFont = "Courier New";

      if (!state.inSource && font === sourceFont) {
        // Scan left until text without source font is found
        while (i > 0 && doc.getFontFamily(attrs[i-1]) && doc.getFontFamily(attrs[i-1]) === sourceFont) {
          i -= 1;
          off = attrs[i];
        }

        formatted = formatMd(formatted, index, '`', lastIndex, '`');

        // Do not handle additional formattings for code
        continue;
      }
    }

    // Handle bold and bold italic
    if(doc.isBold(index)) {
      var dleft, right;
      dleft = dright = "**";
      if (doc.isItalic(index))
      {
        // edbacher: changed this to handle bold italic properly.
        dleft = "**_";
        dright  = "_**";
      }

      formatted = formatMd(formatted, index, dleft, lastIndex, dright);
    }
    // Handle italic
    else if(doc.isItalic(index)) {
      formatted = formatMd(formatted, index, '*', lastIndex, '*');
    }

    // Keep track of last position in text
    lastIndex = index;
  }

  var textElements = [formatted];
  return textElements;
}


function handleListItem(item, state, depth, cmode) {
  var textElements = [];
  
  // Prefix
  var prefix = '  '; // For Dokuwiki, at least one space required 
 
  // Add nesting level
  for (var i=0; i<item.getNestingLevel(); i++) {
    prefix += '  ';
  }
  
  // Add marker based on glyph type
  var glyph = item.getGlyphType();
  Logger.log("Glyph: " + glyph);
  if(cmode == 0){
    switch(glyph) {
      case DocumentApp.GlyphType.BULLET:
      case DocumentApp.GlyphType.HOLLOW_BULLET:
      case DocumentApp.GlyphType.SQUARE_BULLET:
        prefix += '- ';
        break;
      case DocumentApp.GlyphType.NUMBER:
        prefix += '1. ';
        break;
      default:
        prefix += '- ';
        break;
    }    
  }else{
    switch(glyph) {
      case DocumentApp.GlyphType.BULLET:
      case DocumentApp.GlyphType.HOLLOW_BULLET:
      case DocumentApp.GlyphType.SQUARE_BULLET: 
        prefix += '* ';
        break;
      case DocumentApp.GlyphType.NUMBER:
        prefix += '- ';
        break;
      default:
        prefix += '* ';
        break;
    }
  }
  
  // Add prefix
  textElements.push(prefix);
  
  // Handle all childs
  textElements = textElements.concat(processChilds(item, state, depth, cmode));
  
  return textElements;
}

// TODO: Need to support Dokuwiki format
function handleImage(image, state) {
  // Determine file extension based on content type
  var contentType = image.getBlob().getContentType();
  var fileExtension = '';
  if (/\/png$/.test(contentType)) {
    fileExtension = ".png";
  } else if (/\/gif$/.test(contentType)) {
    fileExtension = ".gif";
  } else if (/\/jpe?g$/.test(contentType)) {
    fileExtension = ".jpg";
  } else {
    throw "Unsupported image type: " + contentType;
  }

  // Create filename
  var filename = 'img_' + state.imageCounter + fileExtension;
  state.imageCounter++;
  
  // Add image
  var textElements = []
  textElements.push('![image alt text](' + filename + ')');
  state.images.push( {
    "bytes": image.getBlob().getBytes(), 
    "blob": image.getBlob(), 
    "type": contentType, 
    "name": filename,
  });
  
  return textElements;
}


function processChilds(doc, state, depth, cmode) {
  // Text element buffer
  var textElements = []
  
  // Keep track of child count on this depth
  state.size[depth] = doc.getNumChildren(); 
  
  // Iterates over all childs
  for(var i=0; i < doc.getNumChildren(); i++)  {
    var child = doc.getChild(i); 
    
    // Update pointer on next document
    var nextDoc = (i+1 < doc.getNumChildren())?doc.getChild(i+1) : child;
    state.nextDoc[depth] = nextDoc; 
    
    // Update pointer on prev element 
    var prevDoc = (i-1 >= 0)?doc.getChild(i-1) : child;
    state.prevDoc[depth] = prevDoc; 
    
    if(cmode == 0){
      textElements = textElements.concat(processElementMD(child, state, depth+1)); 
    }else{
      textElements = textElements.concat(processElementDW(child, state, depth+1)); 
    }
  }
  return textElements;
}

function processElementDW(element, state, depth) {
  // Result
  var textElements = [];
    
  switch(element.getType()) {
    case DocumentApp.ElementType.DOCUMENT:
      Logger.log("this is a document"); 
      break; 
      
    case DocumentApp.ElementType.BODY_SECTION: 
      textElements = textElements.concat(processChilds(element, state, depth, 1));
      break; 
      
    case DocumentApp.ElementType.PARAGRAPH:
      // Determine header prefix
      var prefix = ''; 
      var postfix = '';
      switch (element.getHeading()) {
      // Add a = for each heading level. No break, so we accumulate the right number.
      case DocumentApp.ParagraphHeading.HEADING1:
	  prefix += '=';
	  postfix += '=';
      case DocumentApp.ParagraphHeading.HEADING2:
	  prefix += '=';
	  postfix += '=';
      case DocumentApp.ParagraphHeading.HEADING3:
	  prefix += '=';
	  postfix += '=';
      case DocumentApp.ParagraphHeading.HEADING4:
	  prefix += '=';
	  postfix += '=';
      case DocumentApp.ParagraphHeading.HEADING5:
	  prefix += '=';
	  postfix += '=';
      case DocumentApp.ParagraphHeading.HEADING6:
	  prefix += '=';
	  postfix += '=';
      }
      
      // Add space
      if(prefix.length > 0)
        prefix += ' ';
      
      if(postfix.length > 0)
        postfix = ' ' + postfix;

      // Push prefix
      textElements.push(prefix);
      
      // Process childs
      textElements = textElements.concat(processChilds(element, state, depth, 1));
      
      // Add paragraph break only if its not the last element on this layer
      if(state.nextDoc[depth-1] == element)
        break; 
      
      // Push postfix
      textElements.push(postfix);
      
      if(state.inSource)
        textElements.push('\n');
      else
        textElements.push('\n\n');
      
      break; 
      
    case DocumentApp.ElementType.LIST_ITEM:
      textElements = textElements.concat(handleListItem(element, state, depth, 1)); 
      textElements.push('\n');
      
      if(state.nextDoc[depth-1].getType() != element.getType()) {
        textElements.push('\n');
      }
      
      break;
      
    case DocumentApp.ElementType.FOOTNOTE:
      textElements.push(' (NOTE: ');
      textElements = textElements.concat(processChilds(element.getFootnoteContents(), state, depth, 1));
      textElements.push(')');
      break; 
      
    case DocumentApp.ElementType.HORIZONTAL_RULE:
      textElements.push('---\n');
      break; 
     
    case DocumentApp.ElementType.INLINE_DRAWING:
      // Cannot handle this type - there is no export function for rasterized or SVG images...
      break; 
      
    case DocumentApp.ElementType.TABLE:
      textElements = textElements.concat(handleTable(element, state, depth));
      break;
      
    case DocumentApp.ElementType.TABLE_OF_CONTENTS:
      textElements.push('[[TOC]]');
      break;
      
    case DocumentApp.ElementType.TEXT:
      var text = handleTextDW(element, state);
      
      // Check for source code delimiter
      if(/^```.+$/.test(text.join(''))) {
        state.inSource = true; 
      }
      
      if(text.join('') === '```') {
        state.inSource = false; 
      }
      
      textElements = textElements.concat(text);
      break;

    case DocumentApp.ElementType.INLINE_IMAGE: 
      textElements = textElements.concat(handleImage(element, state));
      break; 
      
    case DocumentApp.ElementType.PAGE_BREAK:
      // Ignore page breaks
      break; 
      
    case DocumentApp.ElementType.EQUATION: 
      var latexEquation = handleEquationFunction(element, state); 

      // If equation is the only one in a paragraph - center it 
      var wrap = '$'
      if(state.size[depth-1] == 1) {
        wrap = '$$'
      }
      
      latexEquation = wrap + latexEquation.trim() + wrap; 
      textElements.push(latexEquation);
      break; 
    default:
      throw("Unknown element type: " + element.getType());
  }
  
  return textElements; 
}


function processElementMD(element, state, depth) {
  // Result
  var textElements = [];

  switch(element.getType()) {
    case DocumentApp.ElementType.DOCUMENT:
      Logger.log("this is a document");
      break;

    case DocumentApp.ElementType.BODY_SECTION:
      textElements = textElements.concat(processChilds(element, state, depth, 0));
      break;

    case DocumentApp.ElementType.PARAGRAPH:
      // Determine header prefix
      var prefix = '';
      switch (element.getHeading()) {
        // Add a # for each heading level. No break, so we accumulate the right number.
        case DocumentApp.ParagraphHeading.HEADING6: prefix += '#';
        case DocumentApp.ParagraphHeading.HEADING5: prefix += '#';
        case DocumentApp.ParagraphHeading.HEADING4: prefix += '#';
        case DocumentApp.ParagraphHeading.HEADING3: prefix += '#';
        case DocumentApp.ParagraphHeading.HEADING2: prefix += '#';
        case DocumentApp.ParagraphHeading.HEADING1: prefix += '#';
      }

      // Add space
      if(prefix.length > 0)
        prefix += ' ';

      // Push prefix
      textElements.push(prefix);

      // Process childs
      textElements = textElements.concat(processChilds(element, state, depth, 0));

      // Add paragraph break only if its not the last element on this layer
      if(state.nextDoc[depth-1] == element)
        break;

      if(state.inSource)
        textElements.push('\n');
      else
        textElements.push('\n\n');

      break;

    case DocumentApp.ElementType.LIST_ITEM:
      textElements = textElements.concat(handleListItem(element, state, depth, 0));
      textElements.push('\n');

      if(state.nextDoc[depth-1].getType() != element.getType()) {
        textElements.push('\n');
      }

      break;

    case DocumentApp.ElementType.HEADER_SECTION:
      textElements = textElements.concat(processChilds(element, state, depth, 0));
      break;

    case DocumentApp.ElementType.FOOTER_SECTION:
      textElements = textElements.concat(processChilds(element, state, depth, 0));
      break;

    case DocumentApp.ElementType.FOOTNOTE:
      textElements.push(' (NOTE: ');
      textElements = textElements.concat(processChilds(element.getFootnoteContents(), state, depth, 0));
      textElements.push(')');
      break;

    case DocumentApp.ElementType.HORIZONTAL_RULE:
      textElements.push('---\n');
      break;

    case DocumentApp.ElementType.INLINE_DRAWING:
      // Cannot handle this type - there is no export function for rasterized or SVG images...
      break;

    case DocumentApp.ElementType.TABLE:
      textElements = textElements.concat(handleTable(element, state, depth));
      break;

    case DocumentApp.ElementType.TABLE_OF_CONTENTS:
      textElements.push('[[TOC]]');
      break;

    case DocumentApp.ElementType.TEXT:
      var text = handleTextMD(element, state);

      // Check for source code delimiter
      if(/^```.+$/.test(text.join(''))) {
        state.inSource = true;
      }

      if(text.join('') === '```') {
        state.inSource = false;
      }

      textElements = textElements.concat(text);
      break;

    case DocumentApp.ElementType.INLINE_IMAGE:
      textElements = textElements.concat(handleImage(element, state));
      break;

    case DocumentApp.ElementType.PAGE_BREAK:
      // Ignore page breaks
      break;

    case DocumentApp.ElementType.EQUATION:
      var latexEquation = handleEquationFunction(element, state);

      // If equation is the only one in a paragraph - center it
      var wrap = '$'
      if(state.size[depth-1] == 1) {
        wrap = '$$'
      }

      latexEquation = wrap + latexEquation.trim() + wrap;
      textElements.push(latexEquation);
      break;
    default:
      throw("Unknown element type: " + element.getType());
  }

  return textElements;
}

