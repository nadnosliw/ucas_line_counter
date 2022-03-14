/* ======================= VERSION INFO ======================= //
 Version 1.00.  Manually select text and choose option to count
 number of lines or format the selected text in lines according
 to UCAS methodology.
// ===================== END VERSION INFO ===================== */


// ======================= CREATE MENU ======================= //
function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  DocumentApp.getUi()
  .createMenu('UCAS Line Counter') // Menu name
  .addItem('Count number of lines', 'count_lines_in_selection')
  .addItem('Format selection in UCAS lines','output_selection_formatted_to_lines')
  .addToUi();
}


// ========================= GLOBALS ========================= //
var doc = DocumentApp.getActiveDocument();
var ui = DocumentApp.getUi();
var error_str = 'ERROR';

var max_char_per_line = 94;
var add_line_numbers = true;
var list_of_lines = [];
// ======================= END GLOBALS ======================= //

function add_lines_to_list(string){
  if (string == ''){ // empty string for an extra line break
    list_of_lines.push(string);
    return 1
  }
  var pos_of_all_spaces = get_pos_of_all_spaces(string);
  
  var iteration_counter = 0;
  var current_max = 0;
  var last_end = 0;
  for (var n = 0; n < pos_of_all_spaces.length; n++){
    if (pos_of_all_spaces[n] <= max_char_per_line + last_end){
      var current_max = pos_of_all_spaces[n];
    } else {
      list_of_lines.push(string.substring(last_end,current_max));
      last_end = current_max + 1; // plus one to remove leading space
    }
    
    // logic to handle the final one or two lines
    if (n == pos_of_all_spaces.length - 1){
      // additional logic to catch instances of sentence being one word too long
      if (string.length - last_end > max_char_per_line){
        list_of_lines.push(string.substring(last_end,current_max));
        list_of_lines.push(string.substring(current_max+1,string.length));
      } else {
        list_of_lines.push(string.substring(last_end,string.length));
      }
      break
    }
  }
  return 1
}

function get_pos_of_all_spaces(string){
  var pos_of_spaces = [];
  for (var n = 0; n < string.length; n++){
    if (string[n] == ' '){
      pos_of_spaces.push(n); // Add position number for spaces to list
    }
  }
  return pos_of_spaces
}


function compile_text_in_lines(add_line_numbers){
  var output_str = '';
  var last_item = list_of_lines.pop();
  var line_number = 1;
  // Make all line numbers 2 digits long
  var leading_zero = '';
  for (var n = 0; n < list_of_lines.length; n++){
    if (line_number < 10){
      leading_zero = '0';
    } else {
      leading_zero = '';
    }
    
    // Determine whether to prepend line numbers to each line for output
    if (add_line_numbers){
      output_str = output_str + leading_zero + line_number.toString() + ' ' + list_of_lines[n] + '\n';
      line_number ++;
    } else {
      output_str = output_str + list_of_lines[n] + '\n';
    }
  }
  // Final line is empty due to final line break being added
  // This logic appends the final line which was removed at start of function
  if (add_line_numbers){
    output_str = output_str + leading_zero + line_number.toString() + ' ' + last_item;
  } else {
    output_str = output_str + last_item;
  }
  console.log('output_str generated');
  return output_str
}

function get_selected_text(){
  var selected_text = doc.getSelection();
  
  // Catch errors associated with no text selected
  if (selected_text == null){
    var ui = DocumentApp.getUi();
    var response = ui.alert('Select text to check first', 'This function counts the number of lines in selected text when run.  Please select some text and then retry.', ui.ButtonSet.OK);
    return error_str
  } else {
    return selected_text;
  }
}


function get_text_range_elements(selected_text){
  // Append strings from each range element into this array
  var output_array = [];
  // Catch errors associated with no text selected
  if (selected_text !== error_str){
    var range_elements = selected_text.getRangeElements();
    for (var i = 0; i < range_elements.length; i++){
      var element_text = range_elements[i].getElement().asText().getText();
      // New lines will appear as empty strings
      output_array.push(element_text);
    }
  } else {
    return error_str
  }
  return output_array;
}


// ======================== EXECUTION ======================== //
function count_lines_in_selection() {
  var selected_text = get_selected_text();
  if (selected_text == error_str){
    return 'There was an error in getting the selected text'
  }
  
  // Get all ranges (paragraphs) in the selected text
  var str_array = get_text_range_elements(selected_text);
  
  // Process blocks of text
  for(var n = 0; n<str_array.length; n++){
    add_lines_to_list(str_array[n]);
  }
  
  var number_of_lines = list_of_lines.length;
  var number_of_lines_response = ui.alert('Number of lines', 'This text selection has ' + number_of_lines.toString() + ' lines.', ui.ButtonSet.OK);
  
  console.log('\n\n*** Number of lines returned ***');

  return number_of_lines;
}

function output_selection_formatted_to_lines() {
  var selected_text = get_selected_text();
  if (selected_text == error_str){
    return 'There was an error in getting the selected text'
  }
  
  // Ask user if they want to prepend line numbers to each line
  var add_line_numbers = ui.alert('Add line numbers', 'Do you want to prepend line numbers to each line?', ui.ButtonSet.YES_NO);
  
  if(add_line_numbers == "YES"){
    add_line_numbers = true;
  } else {
    add_line_numbers = false;
  }
  
  // Get all ranges (paragraphs) in the selected text
  var str_array = get_text_range_elements(selected_text);
  
  // Process blocks of text
  for(var n = 0; n<str_array.length; n++){
    add_lines_to_list(str_array[n]);
  }
  
  var number_of_lines = list_of_lines.length;
  
  var output_str = compile_text_in_lines(add_line_numbers);
  var line_formatted_output_response = ui.alert('UCAS formatted output (' + number_of_lines.toString() + ' lines)', 'The text selection conformed to line requirements is shown below:\n\n' + output_str, ui.ButtonSet.OK);
  
  console.log('\n*** Output conformed to line requirements... ***');
}
