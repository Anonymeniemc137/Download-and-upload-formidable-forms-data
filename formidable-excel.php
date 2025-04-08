<?php
/**
 * Formidable Excel
 *
 * @package     FormidableExcel
 * @author      Mihir Dave
 * @copyright   2024 Mihir Dave
 * @license     GPL-2.0-or-later
 *
 * @wordpress-plugin
 * Plugin Name: Formidable Excel
 * Description: Formidable form add-on for convert form to Excel vice-versa
 * Version:     1.0.0
 * Author:      Mihir Dave
 * Text Domain: Formidable-Excel
 * License:     GPL v2 or later
 * License URI: http://www.gnu.org/licenses/gpl-2.0.txt
 */


require WP_PLUGIN_DIR . '/formidable-excel/phpexcel/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

define( 'FORMIDABLE_EXCEL_DELIMITER', ',' );
define( 'FORMIDABLE_EXCEL_ATTRIBUTE', 'data-csv' );

/**
 *  Include js files.
 */
function formidable_excel_register_scripts() {
	wp_register_script( 'formidable-excel-upload', plugin_dir_url( __FILE__ ) . 'formidable-excel-upload.js', array() );
	wp_register_script( 'formidable-excel-download', plugin_dir_url( __FILE__ ) . 'formidable-excel-download.js', array() );
	wp_register_script( 'formidable-xlsx-reader', plugin_dir_url( __FILE__ ) . 'vendors/xlsx.full.min.js', array() );
}
add_action( 'wp_enqueue_scripts', 'formidable_excel_register_scripts' );

/**
 *  Create upload shortcode.
 */
add_shortcode(
	'formidable-excel-upload-button',
	function () {
		wp_enqueue_script( 'formidable-excel-upload' );
		wp_enqueue_script( 'formidable-xlsx-reader' );
		wp_localize_script(
			'formidable-excel-upload',
			'FORMIDABLE_EXCEL',
			array(
				'DELIMITER' => FORMIDABLE_EXCEL_DELIMITER,
				'ATTRIBUTE' => FORMIDABLE_EXCEL_ATTRIBUTE,
			)
		);
		return '<input type="file" id="formidable-excel-upload-button">';
	}
);

/**
 *  Create download shortcode.
 */
add_shortcode(
	'formidable-excel-download-button',
	function ( $atts ) {
		$atts = shortcode_atts(
			array(
				'file-name' => 'export',
				'button-name' => 'Save Progress',
                'form-id'   => '',
				'version'   => '1',
			),
			$atts
		);

		wp_enqueue_script( 'formidable-excel-download' );
        global $wpdb;
        if ( ! empty( $atts['form-id'] ) ){
            $form_name = $wpdb->get_var( $wpdb->prepare( "SELECT description FROM {$wpdb->prefix}frm_forms WHERE id=%d", $atts['form-id'] ) );
        } else {
            $form_name = $atts['file-name'];
        }
        
		wp_localize_script(
			'formidable-excel-download',
			'FORMIDABLE_EXCEL',
			array(
				'DELIMITER'   => FORMIDABLE_EXCEL_DELIMITER,
				'ATTRIBUTE'   => FORMIDABLE_EXCEL_ATTRIBUTE,
				'FILE_NAME'   => $atts['file-name'],
				'AJAX_URL'    => admin_url( 'admin-ajax.php' ),
				'AJAX_NOUNCE' => wp_create_nonce( 'download-excel' ),
			)
		);
        if( $atts['form-id'] == 113 ) {
            return '<button class="upload-btn custom-upload-btn exportExcel" data-form-name="RFP Profile - ' . do_shortcode( "[frm-field-value field_id='4490' user_id='current']" ) . '">' . __( $atts['button-name'], 'formidable' ) . '</button>';
        } else {
            return '<button class="upload-btn custom-upload-btn exportExcel" data-form-name="' . wp_strip_all_tags ( 'RFP - ' . $form_name ) . '">' . __( $atts['button-name'], 'formidable' ) . '</button>';
        }
	}
);


add_action( 'wp_ajax_nopriv_formidable_ajax_download', 'formidable_ajax_download' );
add_action( 'wp_ajax_formidable_ajax_download', 'formidable_ajax_download' );

/**
 *  Download excel by ajax.
 */
function formidable_ajax_download() {

	check_ajax_referer( 'download-excel', 'security' );

	global $wpdb;
	$full_company_name = array();
    $full_company_location = array(); 
	$get_form_id            = sanitize_text_field( $_POST['form_id'] );
	$get_form_title         = sanitize_text_field( $_POST['form_title'] );
	$business_question      = sanitize_text_field( $_POST['business_question'] );
	$is_form_incomplete     = sanitize_text_field( $_POST['form_incomplete'] );
	$get_form_data          = json_decode( html_entity_decode( stripslashes( sanitize_text_field( $_POST['form_data'] ) ) ) );
	$get_form_creation_date = gmdate( 'F d, Y' );

	$get_form_description = $wpdb->get_var( $wpdb->prepare( "SELECT description FROM {$wpdb->prefix}frm_forms WHERE id=%d", $get_form_id ) );
	
	// Create new Spreadsheet object.
	$spreadsheet        = new Spreadsheet();
	$sheet              = $spreadsheet->getActiveSheet();
	$cell_value_counter = 5;
	$key_questions_counter = 1; 
	
	// Remove gridlines.
    $sheet->setShowGridlines(false);

	// Set the width of columns.
	$sheet->getColumnDimension( 'A' )->setWidth( 80 );
	$sheet->getColumnDimension( 'B' )->setWidth( 15 );
	$sheet->getColumnDimension( 'C' )->setWidth( 25 );
	$sheet->getColumnDimension( 'D' )->setWidth( 20 );
	$sheet->getColumnDimension( 'E' )->setWidth( 15 );

	// Apply font style to sheet headings.
	$sheet_heading = array(
		'font' => array(
			'name' => 'Calibri',
			'size' => 24,
			'bold' => true,
		),
	);
    
    // Form incomplete text.
	$form_incomplete_text = array(
		'font' => array(
			'name' => 'Calibri',
			'size' => 18,
			'color' => array( 'rgb' => 'FF0000' ),
		),
	);
	
	// Apply font style to form answers.
	$common_form_answers = array(
		'font' => array(
			'name' => 'Calibri',
			'size' => 14,
			'color' => array( 'rgb' => 'FF0000' ),
		),
	);

    // Apply font style to form answers.
	$common_form_created_date = array(
		'font' => array(
			'name' => 'Calibri',
			'size' => 11,
			'color' => array( 'rgb' => '000000' ),
		),
	);
    
	// Apply font style to form questions.
	$common_form_questions = array(
		'font' => array(
			'name' => 'Calibri',
			'size' => 11,
			'bold' => true,
		),
	);

	// Apply font style to form headings.
	$common_form_headings = array(
		'fill' => array(
			'fillType' => Fill::FILL_SOLID,
			'color'    => array( 'rgb' => '1044CC' ),
		),
		'font' => array(
			'name'  => 'Calibri',
			'size'  => 14,
			'bold'  => true,
			'color' => array( 'rgb' => 'FFFFFF' ),
		),
	);

	// Define the style array to white borders.
	$black_border_style = array(
	    'font' => array(
			'name' => 'Calibri',
			'size' => 14,
			'color' => array( 'rgb' => 'FF0000' ),
		),
		'borders' => array(
			'allBorders' => array(
				'borderStyle' => Border::BORDER_THIN,
				'color'       => array( 'rgb' => '000000' ),
			),
		),
	);
	
	// Define the style array to white borders.
	$common_form_style = array(
	    'font' => array(
			'name' => 'Calibri',
			'size' => 14,
			'color' => array( 'rgb' => 'FF0000' ),
		),
	);
	
	// common form hidden text.
	$common_form_hidden_text = array(
		'font' => array(
			'name' => 'Calibri',
			'size' => 8,
			'color' => array( 'rgb' => 'FFFFFF' ),
		),
	);

	// spreadsheet heading.
	$sheet->setCellValue( 'A1', wp_strip_all_tags( $get_form_description ) );
	$sheet->setCellValue( 'A3', 'Created: ' . $get_form_creation_date );
	$sheet->setCellValue( 'Z1', wp_strip_all_tags( $get_form_id ) );
	$sheet->setCellValue( 'Z100', ' ' );

	// spreadsheet heading style.
	$sheet->getStyle( 'A1' )->applyFromArray( $sheet_heading )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
	$sheet->getStyle( 'Z1' )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
	$sheet->getStyle( 'A3' )->applyFromArray( $common_form_created_date )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
    
    if( $get_form_id != '113' ){
		if( !empty( $is_form_incomplete ) ){
			$sheet->setCellValue( 'A2', wp_strip_all_tags( $is_form_incomplete ) );
			$sheet->getStyle( 'A2' )->applyFromArray( $form_incomplete_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
		} else { 
			$sheet->setCellValue( 'A2', wp_strip_all_tags( 'Rbundle Anonymous Common RFP' ) );
			$sheet->getStyle( 'A2' )->applyFromArray( $common_form_created_date )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
		}
    }
    
	// Get form data and set its cell value.
	// $question_container[0] stands for all the questions.
	// $question_container[1] stands for all the answers.
	// $question_container[2] stands for all the RFP headings.
	
	if( ! empty( $get_form_data ) ){
    	foreach ( $get_form_data as $form_object ) { 
    	    if ( ! empty( $form_object ) ){
        		foreach ( $form_object as $question_container ) { 
        		    if ( $question_container[0] == "Full Company or person Name" ) {
						$sheet->setCellValue(  'Y' . $cell_value_counter, wp_strip_all_tags( 'isQuestion' ) );
						$sheet->setCellValue(  'Z' . $cell_value_counter, wp_strip_all_tags( 'isRepeater' ) );
						$sheet->getStyle( 'Y' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
						$sheet->getStyle( 'Z' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
					}
                    if ( $question_container[0] === 'Profile Nickname' ) {
                        $sheet->setCellValue( 'A2', wp_strip_all_tags( $question_container[1][0] ) );
                        $sheet->getStyle( 'A2' )->applyFromArray( $common_form_answers )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
                    }
        			if ( isset( $question_container[2] ) && ! empty( $question_container[2] ) ) {
        			    // Check for "Questions For The Provider".
						if ( $question_container[2] === 'Questions For The Provider' ) {
							$sheet->setCellValue(  'Y' . $cell_value_counter, wp_strip_all_tags( 'isQuestion' ) );
							$sheet->setCellValue( 'Z' . $cell_value_counter, wp_strip_all_tags( 'isRepeater' ) );
							$sheet->getStyle( 'Y' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
						    $sheet->getStyle( 'Z' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
						}
        				// Check for "Tax History".
                        if ($question_container[0] === 'Tax History') {
                            // Set the Tax History Heading
                            $sheet->setCellValue('A' . $cell_value_counter, wp_strip_all_tags($question_container[2]));
                            $sheet->getStyle('A' . $cell_value_counter)->applyFromArray($common_form_headings)->getAlignment()->setHorizontal('left')->setWrapText(true);
                            $cell_value_counter += 2;
                        
                            // Fetch headings dynamically from the first record
                            if (!empty($question_container[1]) && isset($question_container[1][0])) {
                                $first_record = (array)$question_container[1][0];
                                $column_letter = 'A'; // Start from column A
                                foreach (array_keys($first_record) as $heading) {
                                    $sheet->setCellValue($column_letter . $cell_value_counter, wp_strip_all_tags($heading));
                                    $sheet->getStyle($column_letter . $cell_value_counter)->applyFromArray($common_form_questions)->getAlignment()->setHorizontal('left')->setWrapText(true);
                                    $column_letter++;
                                }
                                $cell_value_counter++;
                            }
                        
                            // Loop for adding the table data
                            if (!empty($question_container[1])) {
                                foreach ($question_container[1] as $record) {
                                    $column_letter = 'A'; // Start from column A
                                    foreach ((array)$record as $value) {
                                        $table_style = $column_letter == 'A' ? $common_form_questions : $black_border_style;
                                        $sheet->setCellValue($column_letter . $cell_value_counter, wp_strip_all_tags($value));
                                        $sheet->getStyle($column_letter . $cell_value_counter)->applyFromArray($table_style)->getAlignment()->setHorizontal('left')->setWrapText(true);
                                        $column_letter++;
                                    }
                                    $cell_value_counter++;
                                }
                            }
                            $cell_value_counter++;
                        } else {
        				    if (!empty($question_container[1][1])) { 
								$sheet->setCellValue(  'Y' . $cell_value_counter, wp_strip_all_tags( 'isQuestion' ) );
								$sheet->setCellValue( 'Z' . $cell_value_counter, wp_strip_all_tags( 'isRepeater' ) );
								$sheet->getStyle( 'Y' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
						        $sheet->getStyle( 'Z' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
							}
        					$sheet->setCellValue( 'A' . $cell_value_counter, wp_strip_all_tags( $question_container[2] ) );
        					$sheet->getStyle( 'A' . $cell_value_counter )->applyFromArray( $common_form_headings )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
        					$cell_value_counter += 2;
        				}
        			}
        			if ($question_container[2] != 'Tax History' && $question_container[0] != 'Profile Nickname') {
                        if ($question_container[0] == 'Full Company or Person Name' || $question_container[0] == 'Location of the lawsuit' ) {
                            if ( $key_questions_counter == 1 ){
                                $sheet->setCellValue(  'Y' . $cell_value_counter, wp_strip_all_tags( 'isQuestion' ) );
								$sheet->setCellValue( 'Z' . $cell_value_counter, wp_strip_all_tags( 'isRepeater' ) );
								$sheet->getStyle( 'Y' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
						        $sheet->getStyle( 'Z' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
								
                                $sheet->setCellValue('A' . $cell_value_counter, wp_strip_all_tags('Full Company or person Name - Location of the lawsuit'));
                                $sheet->getStyle('A' . $cell_value_counter)->applyFromArray($common_form_questions)->getAlignment()->setHorizontal('left')->setWrapText(true);
                                ++$key_questions_counter;
                                ++$cell_value_counter;
                            }
                        } else {
                            if( $question_container[0] != 'Create your own question(s)' ){ 
                                if( $question_container[0] != 'Check to add any of the below questions.' ){
                                    if( $question_container[0] == 'Has the business always been a' ) {
                                        $sheet->setCellValue('A' . $cell_value_counter, wp_strip_all_tags($business_question));
                                        $sheet->getStyle('A' . $cell_value_counter)->applyFromArray($common_form_questions)->getAlignment()->setHorizontal('left')->setWrapText(true);
                                    } else{ 
                                        if (!empty($question_container[1][1])) { 
    										$sheet->setCellValue(  'Y' . $cell_value_counter, wp_strip_all_tags( 'isQuestion' ) );
    										$sheet->setCellValue( 'Z' . $cell_value_counter, wp_strip_all_tags( 'isRepeater' ) );
    										$sheet->getStyle( 'Y' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
    						                $sheet->getStyle( 'Z' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
    									}
                                        if( $question_container[0] != 'Additional Notes' && $question_container[0] != 'Tax History'){
                                            $sheet->setCellValue('A' . $cell_value_counter, wp_strip_all_tags($question_container[0]));
                                            $sheet->getStyle('A' . $cell_value_counter)->applyFromArray($common_form_questions)->getAlignment()->setHorizontal('left')->setWrapText(true);
                                        } else { 
                                            --$cell_value_counter;
                                        }
                                    }
                                    ++$cell_value_counter;
                                }
                            } else {
                                --$cell_value_counter;
                            }
                        }
                    } 
        			
        			if (!empty($question_container[1][1])) { 
                        if ($question_container[0] == "Full Company or Person Name" || $question_container[0] == "Location of the lawsuit") {
                            
                            if( ! empty( $question_container ) ){
                                foreach (array($question_container) as $question_container_inner) {  
                                    $sheet->setCellValue( 'Z' . $cell_value_counter, wp_strip_all_tags( 'isRepeater' ) );
                                    $sheet->getStyle( 'Z' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
                                    if ($question_container_inner[0] === "Full Company or Person Name") {
                                        $full_company_name[] = $question_container_inner[1];
                                    } elseif ($question_container_inner[0] === "Location of the lawsuit") {
                                        $full_company_location[] = $question_container_inner[1];
                                    }
                                }
                            }
                            if( isset( $full_company_name[0] ) && isset( $full_company_location[0] ) ){
                                --$cell_value_counter;
                                foreach ( $full_company_name[0] as $name_index => $name ){
                                    if( !empty( $full_company_location[0][$name_index] ) ){
                                        $sheet->setCellValue('A' . $cell_value_counter, $name . " - " . $full_company_location[0][$name_index] );
                                    } else {
                                        $sheet->setCellValue('A' . $cell_value_counter, $name);
                                    }
                                    $sheet->setCellValue( 'Z' . $cell_value_counter, wp_strip_all_tags( 'isRepeater' ) );
                                    $sheet->getStyle( 'Z' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
                                    $sheet->getStyle('A' . $cell_value_counter)->applyFromArray($common_form_answers)->getAlignment()->setHorizontal('left')->setWrapText(true);
                                    ++$cell_value_counter;
                                }
                            }
                        } else {
                            if( ! empty( $question_container[1] ) && $question_container[0] != 'Tax History' ){
                                foreach ($question_container[1] as $container_repeater_field) {
                                    if ( $container_repeater_field != "Create my own question(s)" ){
                                        $sheet->setCellValue( 'Z' . $cell_value_counter, wp_strip_all_tags( 'isRepeater' ) );
                                        $sheet->getStyle( 'Z' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
                                        $sheet->setCellValue('A' . $cell_value_counter, wp_strip_all_tags($container_repeater_field));
                                        $sheet->getStyle('A' . $cell_value_counter)->applyFromArray($common_form_answers)->getAlignment()->setHorizontal('left')->setWrapText(true);
                                        $cell_value_counter++;
                                    }
                                }
                            }
                        } ++$cell_value_counter;
                    } else { 
                        if ( $question_container[0] == "Full Company or Person Name" || $question_container[0] == "Location of the lawsuit" ){
                            if( ! empty( $question_container ) ){
                                foreach (array($question_container) as $question_container_inner) {  
                                    if ($question_container_inner[0] === "Full Company or Person Name") {
                                        $full_company_name[] = $question_container_inner[1];
                                    } elseif ($question_container_inner[0] === "Location of the lawsuit") {
                                        $full_company_location[] = $question_container_inner[1];
                                    }
                                }
                            }
                            if( isset( $full_company_location[0] ) ){
                                foreach ( $full_company_location[0] as $location_index => $location ){
                                    $sheet->setCellValue('A' . $cell_value_counter, $full_company_name[0][$location_index] . " - " . $location );
                                    $sheet->getStyle('A' . $cell_value_counter)->applyFromArray($common_form_answers)->getAlignment()->setHorizontal('left')->setWrapText(true);
                                    $cell_value_counter+= 2;
                                }
                            }
                        } elseif ( $question_container[0] != "Profile Nickname" ){
                            if( $question_container[1][0] == "Create my own question(s)" ){
                                $cell_value_counter++;
                            } else {  
                                if( $question_container[0] == "Create your own question(s)"){
                                    $sheet->setCellValue( 'A' . $cell_value_counter, wp_strip_all_tags( $question_container[1][0] ) );
                                    $sheet->setCellValue( 'Z' . $cell_value_counter, wp_strip_all_tags( 'isRepeater' ) );
                                    
                    				$sheet->getStyle( 'A' . $cell_value_counter )->applyFromArray( $common_form_answers )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
                    				$sheet->getStyle( 'Z' . $cell_value_counter )->applyFromArray( $common_form_hidden_text )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
                    				$cell_value_counter+= 2;
                                } else {
                    				$sheet->setCellValue( 'A' . $cell_value_counter, wp_strip_all_tags( $question_container[1][0] ) );
                    				$sheet->getStyle( 'A' . $cell_value_counter )->applyFromArray( $common_form_answers )->getAlignment()->setHorizontal( 'left' )->setWrapText( true );
                    				$cell_value_counter+= 2;
                                }
                            }
                        }
        			} 
        		}
    	    }
    	}
	}
	// Set headers to force download.
	header( 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
	header( 'Content-Disposition: attachment;filename="export.xlsx"' );
	header( 'Cache-Control: max-age=0' );

	// Create a writer instance and output the file to the browser.
	$sheet->getProtection()->setSheet( true );
	$writer = new Xlsx( $spreadsheet );
	$writer->save( 'php://output' );
	die;
}