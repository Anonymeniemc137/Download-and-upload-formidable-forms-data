jQuery(document).ready(function () {
    jQuery('.upload-btn #formidable-excel-upload-button').change(function () {
        var elementTypeRepeater;
        var inputTypeRepeater;
        let emptyFieldsRemoval = false;
        let hasRunOnce = false;
        var allQuestions = []; 
        var allRepeaterQuestions = [];
        var dataCheckboxQuestionProviders = [];
        var repeaterFields = []; 
        var currentForm = jQuery(this).parents('.frm_grid_container').next().find('form');
        var checkboxesContainer = jQuery(currentForm).find('[data-rfp-heading="Questions For The Provider"] .frm_opt_container');
        var currentFormId = jQuery(currentForm).find('[name="form_id"]').val();
        var fileInput = jQuery(this)[0];

        if (fileInput.files.length === 0) {
            alert('Please select a file!');
            return false;
        }

        // log all the questions of the form. 
        currentForm.find('[data-rfp-field]').each(function() { 
            jQuery(this).find('.su-tooltip-button').remove();
            jQuery(this).find('.frm_required').remove();
            let formLabelText = jQuery(this).find('.form-label').text().trim();
            allQuestions.push([formLabelText]);
        });
        

        // log all the checkboxes of the questions for the provider. 
        checkboxesContainer.find('input[type="checkbox"]').each(function() {
            var checkboxValue = jQuery(this).val();
            dataCheckboxQuestionProviders.push(checkboxValue);
        });
        
        var file = fileInput.files[0];
        var reader = new FileReader();

        reader.onload = function (e) {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, { type: 'array' });

            // first sheet of the workbook
            var firstSheetName = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[firstSheetName];
            var jsonSheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            var currentQuestion = null;

            // If form id is not same return. 
            if (currentFormId != jsonSheetData[0][25]) {
                alert('Please select a valid RFP!');
                return false;
            }
            
            if (jQuery('div.business-always').length) {
                jQuery(currentForm).find('.frm_fields_container .business-always .form-label').next('.frm_opt_container').find('input[type="radio"][value="Yes"]').prop('checked', true).trigger('change');
            }
           
            // Process the JSON data and populate the form fields
            jsonSheetData.forEach(function (row, index) { 
                // Check if the 25th index contains "isRepeater"
                if (row[25] === "isRepeater") {
                    var valueRepeater = row[0];
                    repeaterFields.push(row);
                    var providerQuestionsIndex = index;
                
                    // Collect all questions marked as "isQuestion"
                    if (row[24] === "isQuestion") {
                        allRepeaterQuestions.push(row[0]);
                    }
                
                    // Iterate over the checkboxes in "Questions For The Provider" and store their values
                    checkboxesContainer.find('input[type="checkbox"]').each(function() {
                        var checkboxValue = jQuery(this).val();
                        if (checkboxValue && checkboxValue.length > 0) { 
                            if (checkboxValue === jsonSheetData[providerQuestionsIndex][0]) {
                                jQuery(this).prop('checked', true).trigger('change');
                                providerQuestionsIndex++;
                            }
                        }
                    }); 
                    if (allRepeaterQuestions.length) {
                        allRepeaterQuestions.forEach(function (currentQuestion, questionIndex) {
                            // Only execute if the current question matches the repeater value
                            if (currentQuestion === valueRepeater) { 
                                
                                var questionRepeaterField = jQuery(currentForm).find('[data-rfp-field] .form-label:contains("' + currentQuestion + '")');
                                        
                                if (questionRepeaterField.length === 0) {
                                    questionRepeaterField = jQuery(currentForm).find('[data-rfp-heading="' + currentQuestion + '"]');
                                }
                                if(String(currentQuestion).indexOf('Full Company or person Name - Location of the lawsuit') > -1 ){
                                    for (let i = index + 1; i < jsonSheetData.length; i++) { 
                                        if (jsonSheetData[i][25] && !jsonSheetData[i][24]) {
                                            if (jsonSheetData[i][0] != undefined && jsonSheetData[i][0] != '' && jsonSheetData[i][0].length) {
                                                // Split the string at the last occurrence of the hyphen
                                                let fullString = jsonSheetData[i][0];
                                                let lastHyphenIndex = fullString.lastIndexOf(" - ");
                                
                                                // Extract personName and location
                                                let personName = fullString.substring(0, lastHyphenIndex).trim();
                                                let location = fullString.substring(lastHyphenIndex + 3).trim();
                                
                                                // Only clone once for each entry
                                                let repeatSectionConflict = jQuery(currentForm).find('.legal-service-conflict-person-name');
                                                let insertAfterConflict = repeatSectionConflict.parents('.frm_form_field').find('.frm_repeat_sec:last');
                                                
                                                let clonedSection = insertAfterConflict.clone().insertAfter(insertAfterConflict);
                                
                                                // Insert values into the cloned fields
                                                clonedSection.find('input').val(personName).trigger('change');
                                                clonedSection.find('select').val(location).trigger('change');
                                                clonedSection.find('input').attr("value", jsonSheetData[i][0]).trigger('change');
                                            }
                                        } else { 
                                            break;
                                        }
                                    }
                                } else if (String(currentQuestion).indexOf('Full Company or person Name') > -1 && String(currentQuestion).indexOf('Full Company or person Name - Location of the lawsuit') == -1) {
                                    for (let i = index + 1; i < jsonSheetData.length; i++) { 
                                        if (jsonSheetData[i][25] && !jsonSheetData[i][24]) {
                                            if (jsonSheetData[i][0] !== undefined && jsonSheetData[i][0] !== '') {
                                                let repeatSection = questionRepeaterField.closest('.other-parties-involved').parents(".frm_form_field");
                                                repeatSection.find(".frm_repeat_sec:last").clone().insertAfter(repeatSection);
                                                repeatSection.find(".frm_repeat_sec:last input").val(jsonSheetData[i][0]).trigger('change');
                                                repeatSection.find(".frm_repeat_sec:last input").attr("value", jsonSheetData[i][0]).trigger('change');
                                            }
                                        } else { 
                                            break; 
                                        }
                                    }
                                } else if (String(currentQuestion).indexOf('Questions For The Provider') > -1) { 
                                    localCheckboxQuestionsCounter = 0;
                                    let indexQuestion = index + 1;
                                    
                                    // Set up radio buttons
                                    jQuery('.additional-questions').find('input[type="radio"][value="Yes"]').prop('checked', true).trigger('change');
                    
                                    for (let i = indexQuestion + 1; i < jsonSheetData.length; i++) { 
                                        if (jsonSheetData[i][25] && !jsonSheetData[i][24]) { 
                                            if (jsonSheetData[i][0] !== undefined && jsonSheetData[i][0] !== '' && jsonSheetData[i][0] != dataCheckboxQuestionProviders[localCheckboxQuestionsCounter] ) { 
                                                // check if there are any additional questions. 
                                                jQuery('[data-rfp-heading="Questions For The Provider"] .frm_opt_container').find('.checkbox input[type="checkbox"]:last').prop('checked', true).trigger('change');
                                                questionRepeaterField.next().next().find('.frm_repeat_inline').last().clone().insertAfter('.add-custom-questions');
                                                questionRepeaterField.next().next().find('.frm_repeat_inline').last().children('.frm_form_field').find('input').val(jsonSheetData[i][0]).trigger('change');
                                                questionRepeaterField.next().next().find('.frm_repeat_inline').last().children('.frm_form_field').find('input').attr("value", jsonSheetData[i][0]).trigger('change');
                                            } 
                                            localCheckboxQuestionsCounter++;
                                        } else { 
                                            break;
                                        }
                                    }
                                } else {  
                                    if( String(currentQuestion).indexOf('Questions For The Provider') == -1 &&  String(currentQuestion).indexOf('Full Company or person Name') == -1 && String(currentQuestion).indexOf('Full Company or person Name - Location of the lawsuit') == -1 ){
                                        
                                        if (!questionRepeaterField.parents('.frm_section_heading').hasClass('elseRepeaterFields')) {
                                            questionRepeaterField.parents('.frm_section_heading').addClass('elseRepeaterFields');
                                        }
                                        
                                        var fieldElementRepeater = jQuery(currentForm).find('[data-rfp-field] .form-label:contains("' + currentQuestion + '")').closest('.frm_form_field').find('input, select, textarea');
                                        
                                        if (fieldElementRepeater.length > 0) {
                                            var elementTypeRepeater = fieldElementRepeater[0].nodeName.toLowerCase();
                                            var inputTypeRepeater = fieldElementRepeater.attr('type');
                                        }
                                        // logic for input type text.
                                        if (inputTypeRepeater === 'text') {
                                            for (let i = index + 1; i < jsonSheetData.length; i++) {
                                                if (jsonSheetData[i][25] && !jsonSheetData[i][24]) {
                                                    if (jsonSheetData[i][0] !== undefined && jsonSheetData[i][0] !== '') {  
                                                        questionRepeaterField.parents('.frm_form_field').find('.frm_repeat_sec').last().clone().insertAfter(questionRepeaterField.parents('.frm_repeat_sec'));
                                                        questionRepeaterField.parents('.frm_form_field').find('.frm_repeat_sec').last().find('input.form-control').val(jsonSheetData[i][0]).trigger('change');
                                                        questionRepeaterField.parents('.frm_form_field').find('.frm_repeat_sec').last().find('input.form-control').attr("value", jsonSheetData[i][0]).trigger('change');
                                                    } 
                                                } else { 
                                                    break;
                                                }
                                            }
                                        }
                                       
                                        // logic for input type checkbox.
                                        if (inputTypeRepeater === 'checkbox') { 
                                            fieldElementRepeater.each(function() {
                                                for (let i = index + 1; i < jsonSheetData.length; i++) {
                                                    if (jsonSheetData[i] && jsonSheetData[i][25] && !jsonSheetData[i][24]) {
                                                        if (jsonSheetData[i][0] !== undefined && jsonSheetData[i][0] !== '') {
                                                            if (jQuery(this).val() == jsonSheetData[i][0]) {
                                                                jQuery(this).prop('checked', true).trigger('change');
                                                            }
                                                        }
                                                    } else {
                                                        break; // Break if condition is not met
                                                    }
                                                }
                                            });
                                        }
                                    }
                                }
                            }
                        });
                    }
                } else if (row && row.length > 0) {
                    var value = row[0];
                    
                    // Check if it's a question by searching in the form labels
                    var questionField = jQuery(currentForm).find('[data-rfp-field] .form-label:contains("' + value + '")');

                    if (value && String(value).indexOf('Additional Notes') > -1) {
                        jQuery('[data-rfp-field-label="Additional Notes"]').closest('.frm_form_field').find('textarea').val(jsonSheetData[index + 3][0]).trigger("change");
                    }

                    if (questionField.length) { 
                        // It's a question, reset the currentQuestion and prepare for the next answer
                        currentQuestion = value;
                        // If it's an answer, associate it with the current question
                        
                       fieldElementDefault = jQuery(currentForm).find('[data-rfp-field] .form-label:contains("' + currentQuestion + '")').parent().find('input, select, textarea');
                        if (fieldElementDefault.length) { 
                            var elementTypeDefault = fieldElementDefault[0].nodeName.toLowerCase();
                            
                            if (elementTypeDefault === 'input') {
    							var inputTypeDefault = fieldElementDefault.attr('type');
    						} else { 
    						    inputTypeDefault = elementTypeDefault;
    						}
    						
                            switch (elementTypeDefault) {
                                case 'input':
                                    if (inputTypeDefault === 'radio' || inputTypeDefault === 'checkbox') {
                                        if (value) { 
                                            if (inputTypeDefault === 'radio') { 
                                                var radioButton = fieldElementDefault.closest('.frm_form_field').find('input[type="radio"][value="' + jsonSheetData[index + 1][0] + '"]');
                                                if (radioButton.length) {
                                                    radioButton.prop('checked', true).trigger('change');
                                                }
                                            } else if (inputTypeDefault === 'checkbox') {
                                                fieldElementDefault.prop('checked', value === 'true').trigger('change');
                                            } 
                                        }
                                    } else if (inputTypeDefault === 'number') {
                                        fieldElementDefault.val(value).trigger('change');
                                    } else {
                                        fieldElementDefault.val(value).trigger('change');
                                    }
                                    break;

                                case 'select':
                                    if (String(value).indexOf('Other') === 0) {
                                        var otherField = fieldElementDefault.siblings('.frm_other_input');
                                        if (otherField.length) {
                                            var otherVal = value.split('—');
                                            if (otherVal.length === 2) {
                                                otherField.val(otherVal[1].trim()).trigger('change');
                                            }
                                        }
                                        value = 'Other';
                                    }
                                    fieldElementDefault.find('option[value="' + value + '"]').prop('selected', true).trigger('change');
                                    break;

                                case 'textarea': 
                                    if (value != undefined) { 
                                        fieldElementDefault.val(value); 
                                        fieldElementDefault.trigger('input');
                                        fieldElementDefault.trigger('change');
                                    }
                                    break;
                            }
                        } else {
                            console.warn('No form field found for question: ' + currentQuestion);
                        }
                    } else if (currentQuestion) { 
                        // If it's an answer, associate it with the current question
                        var fieldElement = jQuery(currentForm).find('[data-rfp-field] .form-label:contains("' + currentQuestion + '")').next('input, select, textarea');

                        if (fieldElement.length === 0) {
                            fieldElement = jQuery(currentForm).find('[data-rfp-field] .form-label:contains("' + currentQuestion + '")').parent().find('input, select, textarea');
                        }
                        
                        if (fieldElement.length) { 
                            var elementType = fieldElement[0].nodeName.toLowerCase();
                            
                            if (elementType === 'input') {
    							var inputType = fieldElement.attr('type');
    						} else { 
    						    inputType = elementType;
    						}
    						
                            switch (elementType) {
                                case 'input':
                                    if (inputType === 'radio' || inputType === 'checkbox') {
                                        if (value) { 
                                            if (inputType === 'radio') {
                                                fieldElement.parents('.frm_opt_container').children('.frm_radio').each(function() {
                                                    if( jQuery(this).find('input').val() == value ){ 
                                                        jQuery(this).find('input').prop('checked', true).trigger('change');    
                                                    }
                                                });
                                            } else if (inputType === 'checkbox') {
                                                fieldElement.parents('.frm_opt_container').children('.frm_checkbox').each(function() {
                                                    if( jQuery(this).find('input').val() == value ){ 
                                                        jQuery(this).find('input').prop('checked', true).trigger('change');    
                                                    }
                                                });
                                            } 
                                        }
                                    } else if (inputType === 'number') {
                                        fieldElement.val(value).trigger('change');
                                    } else {
                                        fieldElement.val(value).trigger('change');
                                    }
                                    break;

                                case 'select':
                                    if (String(value).indexOf('Other') === 0) {
                                        var otherField = fieldElement.siblings('.frm_other_input');
                                        if (otherField.length) {
                                            var otherVal = value.split('—');
                                            if (otherVal.length === 2) {
                                                otherField.val(otherVal[1].trim()).trigger('change');
                                            }
                                        }
                                        value = 'Other';
                                    }
                                    fieldElement.find('option[value="' + value + '"]').prop('selected', true).trigger('change');
                                    break;

                                case 'textarea': 
                                    if (value != undefined) { 
                                        fieldElement.val(value); 
                                        fieldElement.trigger('input');
                                        fieldElement.trigger('change');
                                    }
                                    break;
                            }
                        } else {
                            console.warn('No form field found for question: ' + currentQuestion);
                        }
                    } 
                }
            });
            if (!emptyFieldsRemoval) {
                emptyFieldsRemoval = true;
                jQuery(currentForm).find('.add-custom-questions').siblings('.frm_first_repeat').last().remove();
                jQuery(currentForm).find('.other-parties-involved').parents(".frm_form_field").siblings('.frm_repeat_sec').last().remove();
                jQuery(currentForm).find('.legal-service-conflict-person-name').parents('.frm_repeat_sec').last().remove();
                currentForm.find('.elseRepeaterFields').each(function() { 
                    jQuery(this).find('.frm_repeat_sec').each(function(){
                        var RepeaterFields = jQuery(this).find('input.form-control').val();
                        if(!RepeaterFields.length){
                            jQuery(this).remove();
                        }

                    });
                });
                
            }
        };

        reader.onerror = function (error) {
            console.error('Error reading Excel file: ', error);
        };

        reader.readAsArrayBuffer(file);
    });
});

// ******************  OLD CODEEE *******************
// document.getElementById('formidable-excel-upload-button').addEventListener('change', function (e) {
//     var file = e.target.files[0];
//     if (!file) {
//         return false;
//     }
//     jQuery(e.target).parents('.rfp-view').find('form').trigger('reset');
//     var reader = new FileReader();
//     reader.onload = readRfpDataFile;
//     reader.readAsText(file);
// });

// var incrementalIndex = {
//     'Common Questions': 1,
//     'Question For Providers': 1,
// };

// var incrementalNextIndex = {
//     'Additional Notes': 2,
//     'Question For Providers': 2,
// };

// function readRfpDataFile(e) {
//     var reader = e.target;
//     var rows = reader.result.split(/\r\n|\r|\n/g);
//     var parse = function (row) {
//         var insideQuote = false, entries = [], entry = [];
//         if (typeof row === 'undefined') {
//             return '';
//         }
//         row.split('').forEach(function (character) {
//             if (character === '"') {
//                 insideQuote = !insideQuote;
//             } else {
//                 if (character === FORMIDABLE_EXCEL.DELIMITER && !insideQuote) {
//                     entries.push(entry.join(''));
//                     entry = [];
//                 } else {
//                     entry.push(character);
//                 }
//             }
//         });
//         entries.push(entry.join(''));
//         entries = entries.filter(function (entry) {
//             return entry !== '';
//         });
//         return entries.length > 1 ? entries : entries[0];
//     };
//     var totalRows = rows.length;
//     var taxData = rows.filter(function (row) {
//         var column = parse(row);
//         return typeof column === 'object';
//     });
//     var index;
//     for (index = 0; index < totalRows; index++) {
//         var row = rows[index];

//         var column = parse(row);
//         if (typeof incrementalIndex[column] === 'number') {
//             index+= incrementalIndex[column];
//             break;
//         }
//         var nextIndex = index + 1;
//         if (typeof incrementalNextIndex[column] === 'number') {
//             nextIndex = index + incrementalNextIndex[column];
//         }
//         var nextColumn = parse(rows[nextIndex]);
//         if (nextColumn === '' || nextColumn === undefined) {
//             continue; // Treat as heading/answer of a question already processed
//         }
//         if (typeof column === 'string') {
//             column = column.replace(/"([^"]+(?="))"/g, '$1');
//             var value = nextColumn.replace(/"([^"]+(?="))"/g, '$1');
//             if (jQuery('[data-rfp-heading="' + value + '"]').length) {
//                 continue; // Matched the section heading
//             }
//             if (typeof nextColumn !== 'string') {
//                 continue; // Don't process tax data array
//             }
//             var field = jQuery('[data-rfp-field]:contains("' + column + '")');
//             if (!field.length) {
//                 continue; // Didn't find the field
//             }
//             var valueSet = false;
//             var setFieldVal = function (index, domElement) {
//                 if (valueSet) {
//                     return;
//                 }
//                 var $this = jQuery(domElement);
//                 var element = $this.prop('tagName').toString().toLowerCase();
//                 // console.log(`Element: ${element}, Value: ${nextColumn}`);

//                 switch (element) {
//                     case 'input':
//                         // console.log(`Checking ${$this.val()} === ${value}`);
//                         if ($this.attr('type') === 'radio' || $this.attr('type') === 'checkbox') {
//                             if ($this.val() === value) {
//                                 $this.prop('checked', true).trigger({ type: 'change', originalEvent: 'custom' });
//                                 valueSet = true;
//                             }
//                         } else {
//                             $this.val(value).trigger({ type: 'change', originalEvent: 'custom' });
//                             valueSet = true;
//                         }
//                         break;
//                     case 'select':
//                         if (value.indexOf('Other') === 0) {
//                             var otherField = $this.siblings('.frm_other_input');
//                             if (otherField) {
//                                 var otherVal = value.split('—');
//                                 if (otherVal.length === 2) {
//                                     otherField.val(otherVal[1].trim()).trigger({ type: 'change', originalEvent: 'custom' });
//                                 }
//                             }
//                             value = 'Other';
//                         }
//                         $this.find('option[value="' + value + '"]').prop('selected', true);
//                         $this.trigger({ type: 'change', originalEvent: 'custom' });
//                         valueSet = true;
//                         break;
//                     case 'textarea':
//                         $this.val(value).trigger({ type: 'change', originalEvent: 'custom' });
//                         valueSet = true;
//                         break;
//                 }
//             };
//             field.find('input, select, textarea').each(setFieldVal);
//         }
//     }
//     var additionalQuestionsSection = jQuery('.additional-questions');
//     var triggerQuestionSections = additionalQuestionsSection.find('.frm_trigger');
//     if (additionalQuestionsSection.find('[data-rfp-field]').is(':visible')) {
//         triggerQuestionSections.click(); // Show the additional questions section
//     }
//     var commonQuestionField = jQuery('[data-rfp-field]:contains("Common Questions")');
//     var customQuestionField = commonQuestionField.parents('.frm_form_field').siblings('.add-custom-questions');
//     var addMoreQuestionBtn = customQuestionField.find('.frm_add_form_row').first();
//     for (; index < totalRows; index++) {
//         var row = rows[index];
//         var column = parse(row);
//         if (!column || typeof column !== 'string') {
//             continue;
//         }
//         column = column.replace(/"([^"]+(?="))"/g, '$1');
//         var checkbox = commonQuestionField.find('.frm_checkbox:contains("' + column + '")');
//         if (checkbox.length) {
//             checkbox.find('input').prop('checked', true);
//             continue;
//         } else {
//             break;
//         }
//     }
//     jQuery('.frm_remove_form_row', customQuestionField).click();
//     function addCustomQuestion(index) {
//         function eventHandler() {
//             setTimeout(function () {
//                 var row = rows[index];
//                 if (typeof row === 'undefined') {
//                     console.log('returning, row is undefined', row);
//                     return;
//                 }
//                 var column = parse(row);
//                 if (!column || typeof column !== 'string') {
//                     console.log('returning', column);
//                     return;
//                 }
//                 var value = column.replace(/"([^"]+(?="))"/g, '$1');
//                 var customField = customQuestionField.children('.frm_repeat_inline').last();
//                 if (customField) {
//                     customField.find('input').val(value);
//                 }
//                 if (index <= rows.length - 1) {
//                     addMoreQuestionBtn.click();
//                     index++;
//                 } else {
//                     jQuery(document).off('frmAfterAddRow', eventHandler);
//                 }
//             }, 150);
//         }
//         jQuery(document).on('frmAfterAddRow', eventHandler);
//         addMoreQuestionBtn.click();
//     }
//     addCustomQuestion(index);

//     if (taxData.length) {
//         jQuery('#body-table table#form-table tbody').empty();
//         taxData.forEach(function (tax) {
//             var data = tax.split(FORMIDABLE_EXCEL.DELIMITER);
//             if (data[0] === '" "' && data[1] === '"Tax Years"') {
//                 return; // Header
//             }
//             addTableRow();
//             var newRow = jQuery('#body-table table#form-table tbody tr').last();
//             var cols = newRow.children('td');
//             data.forEach(function (value, index) {
//                 jQuery(cols[index + 1]).text(value.replace(/"([^"]+(?="))"/g, '$1'));
//             });
//         });
//     }
// }
