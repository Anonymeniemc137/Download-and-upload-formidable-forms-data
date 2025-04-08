function generateRfpData(formID) {
    var rfpData = {};
    jQuery('#frm_form_' + formID + '_container [data-rfp-field]').each(function (i, item) {
        var $this = jQuery(item);
        if (!$this.is(':visible')) {
            return;
        }
        var field = $this.data('rfp-field-label');
        var label = $this.find('.frm_primary_label .control-label');
        if (!label.length) {
            label = $this.find('.frm_primary_label');
        }
        if (!field) {
            if (label && label[0]) {
                field = (label[0].childNodes[0] ? label[0].childNodes[0].nodeValue : label[0].nodeValue).trim();
            }
        }
        if ($this.data('rfp-tax-history') !== undefined) {
            field = 'Tax History';
            var taxTables = $this.find('table.rbundle-html-table'); // Find all tables with the class rbundle-html-table
            var taxHistory = [];
        
            // Process each table individually
            taxTables.each(function(tableIndex, table) {
                var $table = jQuery(table);
                var headers = [];
        
                // Fetch the headings dynamically for the current table
                $table.find('thead th').each(function() {
                    headers.push(jQuery(this).text().trim());
                });
        
                // Iterate through table rows for the current table
                $table.find('tbody tr').each(function(rowIndex, row) {
                    var $row = jQuery(row);
                    var cols = $row.children('td');
                    var tax = {};
       
                    // Iterate through columns dynamically
                    cols.each(function(colIndex, col) {
                        var $col = jQuery(col);
                        var columnValue = '';
                     
                        // Handle select dropdowns by taking the selected option
                        if ($col.find('select').length) {
                            columnValue = $col.find('select option:selected').text();
                            if (columnValue === 'other') {
                                columnValue = $col.find('.form-control').val();
                                
                            } else {
                                columnValue = $col.find('select option:selected').text();
                            }
                        } else {
                            columnValue = $col.text().trim();
                        }
                    
                        tax[headers[colIndex]] = columnValue;
                    });
        
                    taxHistory.push(tax);
                });
            });
        
            // Update the rfpData object with collected data
            rfpData[$this.data('rfp-field')] = [field, taxHistory, $this.data('rfp-heading')];
            return;
        }



        // if ($this.data('rfp-tax-history') !== undefined) {
        //     field = 'Tax History';
        //     var taxTable = $this.find('table#form-table');
        //     var taxHistory = [];
        //     taxTable.find('tbody tr').each(function (index, row) {
        //         var $row = jQuery(row);
        //         var cols = $row.children('td');
        //         var tax = {
        //             ' ': cols[1].textContent,
        //             'Tax Years': cols[2].textContent,
        //             'Federal Income Tax Form': cols[3].textContent,
        //             'Legal Entity': cols[4].textContent,
        //             'Subject to BBA': cols[5].textContent,
        //         };
        //         taxHistory.push(tax);
        //     });
            
        //     console.log(taxHistory);
        //     return false;
        //     rfpData[$this.data('rfp-field')] = [field, taxHistory, $this.data('rfp-heading')];
        //     return;
        // }
        if (field === '') {
            return;
        }
        var fieldVal = '';
        var setFieldVal = function (index, input) {
            var $input = jQuery(input);
            var element = $input.prop('tagName').toString().toLowerCase();

            switch (element) {
                case 'input':
                    if ($input.attr('type') === 'radio') {
                        if ($input.prop('checked')) {
                            fieldVal = $input.val();
                        }
                    } else if ($input.attr('type') === 'checkbox') {
                        if ($input.prop('checked')) {
                            fieldVal += $input.val() + '\r\n\r\n';
                        }
                    } else {
                        fieldVal = $input.val();
                    }
                    break;
                case 'select':
                case 'textarea':
                    fieldVal = $input.val();
                    if (element === 'select' && $input.val() === 'Other') {
                        var otherField = $input.siblings('.frm_other_input');
                        if (otherField) {
                            var otherVal = otherField.val().trim();
                            if (otherVal === '') {
                                fieldVal = '';
                                return;
                            }
                            fieldVal = fieldVal + ' — ' + otherVal;
                        }
                    }
                    break;
            }
        };
        $this.find('input').each(setFieldVal);
        $this.find('textarea').each(setFieldVal);
        $this.find('select').each(setFieldVal);
      
        if (field && fieldVal) {
            fieldVal = String(fieldVal);  // Ensure fieldVal is a string
            if (typeof rfpData[$this.data('rfp-field')] === 'undefined') {
                rfpData[$this.data('rfp-field')] = [field, fieldVal.trim().split('\r\n\r\n'), $this.data('rfp-heading')];
            } else {
                var existing = rfpData[$this.data('rfp-field')];
                if (typeof existing[1] === 'string') {
                    existing[1] = [existing[1]];
                }
                existing[1].push(fieldVal.trim());
            }
        }
    });
    return rfpData;
}


function shouldIgnoreLabel(label) {
    var ignoreLabels = ['Additional Notes', 'Common Questions'];
    var total = ignoreLabels.length;
    for (var i = 0; i < total; i++) {
        if (ignoreLabels[i] === label) {
            return true;
        }
    }

    return false;
}


function generateRfpCsvFile(data) {
    
    let result = [];
    let currentObject = {};
    
    Object.keys(data).forEach((key, index) => {
        if (data[key][2] !== null && data[key][2] !== undefined) {
            if (Object.keys(currentObject).length > 0) {
                result.push(currentObject);
            }
            currentObject = {};
        }
        currentObject[key] = data[key];
    });
    
    if (Object.keys(currentObject).length > 0) {
        result.push(currentObject);
    }
    
    return result;   
}   


/* AJAX call for download - excel */
jQuery(document).ready(function() {
    jQuery(document).on('click', '.exportExcel', function() {
        var hiddenFieldValue;
        var formID = jQuery(this).parents(".frm_grid_container").siblings().find('.frm_fields_container').find('input[name="form_id"]').val();
        if (typeof formID === 'undefined') {
            formID = jQuery(this).parents('.frm_fields_container').find('input[name="form_id"]').val();
        }
        if( formID != 113 && jQuery('div.business-always').length ){
            var businessAlwaysQuestion = jQuery(this).parents(".frm_grid_container").siblings().find('.frm_fields_container .business-always .form-label span').html().split("?")[0] + "?";
        } else {
            businessAlwaysQuestion = '';
        }
        if (typeof formID === 'undefined') {
            formID = jQuery(this).parents('.frm_fields_container').find('input[name="form_id"]').val();
        } 
        
        let formDownloadName = jQuery(this).attr("data-form-name");
        let formTitle = jQuery('.frm_grid_container h1').text();
        var formData = generateRfpData(formID);
        if(jQuery("[required-rfp-field]").length>0){
            var hiddenFieldValue = jQuery('#rfp_hidden_field_' + formID).val().includes("IN PROGRESS") ? jQuery('#rfp_hidden_field_' + formID).val() : '';
        }
        var excelData = '';
        if (formData) {
            excelData = generateRfpCsvFile(formData);
        } 
        
        var stringifiedExcelData = JSON.stringify(excelData);
        
        console.log(stringifiedExcelData);
        jQuery.ajax({
            url: FORMIDABLE_EXCEL.AJAX_URL,
            method: 'POST',
            data : {
              security: FORMIDABLE_EXCEL.AJAX_NOUNCE,
              action: 'formidable_ajax_download',
              form_id: formID,
              form_title: formTitle,
              form_data: stringifiedExcelData,
              business_question : businessAlwaysQuestion,
              form_incomplete : hiddenFieldValue,
            },
            xhrFields: {
                responseType: 'blob'
            },
            success: function(data) {
                var blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var link = document.createElement('a');
                link.href = window.URL.createObjectURL(blob);
                link.download = formDownloadName +'.xlsx';
                document.body.appendChild(link);
                link.click(); // Trigger the download
                document.body.removeChild(link);
            },
            error: function(xhr, status, error) {
                console.error('Error:', error);
            }
        });
    });
});