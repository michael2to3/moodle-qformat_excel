<?php
// Strings used in format.php.
$string['cannotopentempfile'] = 'Cannot open temporary file <b>{$a}</b>';
$string['cannotreadzippedfile'] = 'Cannot read Zipped file <b>{$a}</b>';
$string['cannotwritetotempfile'] = 'Cannot write to temporary file <b>{$a}</b>';
$string['docnotsupported'] = 'Files in Xlsx 2003 format not supported: <b>{$a}</b>, use Moodle2Xlsx 3.x instead';
$string['htmldocnotsupported'] = 'Incorrect Xlsx format: please use <i>File>Save As...</i> to save <b>{$a}</b> in native Xlsx 2010 (.docx) format and import again';
$string['htmlnotsupported'] = 'Files in HTML format not supported: <b>{$a}</b>';
$string['noquestions'] = 'No questions to export';
$string['pluginname'] = 'Microsoft Xlsx 2008 table format (xlsxtable)';
$string['pluginname_help'] = 'This is a front-end for converting Microsoft Xlsx 2008 files into Moodle Question XML format for import, and converting Moodle Question XML format into a format suitable for editing in Microsoft Xlsx.';
$string['pluginname_link'] = 'qformat/xlsxtable';
$string['preview_question_not_found'] = 'Preview question not found, name / course ID: {$a}';
$string['privacy:metadata'] = 'The XlsxTable question format plugin does not store any personal data.';
$string['stylesheetunavailable'] = 'XSLT Stylesheet <b>{$a}</b> is not available';
$string['transformationfailed'] = 'XSLT transformation failed (<b>{$a}</b>)';
$string['xlsxtable'] = 'Microsoft Xlsx 2008 table format (xlsxtable)';
$string['xlsxtable_help'] = 'This is a front-end for converting Microsoft Xlsx 2008 files into Moodle Question XML format for import, and converting Moodle Question XML format into an enhanced XHTML format for exporting into a format suitable for editing in Microsoft Xlsx.';
$string['xmlnotsupported'] = 'Files in XML format not supported: <b>{$a}</b>';
$string['xsltunavailable'] = 'You need the XSLT library installed in PHP to save this Xlsx file';

// Strings used in XSLT when converting questions into Xlsx format.
$string['cloze_distractor_column_label'] = 'Distractors';
$string['cloze_feedback_column_label'] = 'Distractor Feedback';
$string['cloze_instructions'] = 'Use <strong>bold</strong> for Multichoice, <em>italic</em> for Short Answer, and <u>Underline</u> for Numerical questions.';
$string['cloze_mcformat_label'] = 'Orientation (D = dropdown; V = vertical, H = horizontal radio buttons)';
$string['description_instructions'] = 'This is not actually a question. Instead it is a way to add some instructions, rubric or other content to the activity. This is similar to the way that labels can be used to add content to the course page.';
$string['essay_instructions'] = 'Allows a response of a few sentences or paragraphs. This must then be graded manually.';
$string['interface_language_mismatch'] = 'No questions imported because the language of the labels in the Xlsx file does not match your current Moodle interface language.';
$string['multichoice_instructions'] = 'Allows the selection of a single or multiple responses from a pre-defined list.';
$string['truefalse_instructions'] = 'Set grade \'100\' to the correct answer.';
$string['unsupported_instructions'] = 'Importing this question type is not supported.';

// These strings are part of the Xlsx Startup template user interface, not the Moodle interface.
// These templates are available at http://www.moodle2xlsx.net/.
$string['xlsx_about_moodle2xlsx'] = 'About Moodle2Xlsx';
$string['xlsx_about_moodle2xlsx_screentip'] = 'About the Moodle2Xlsx Xlsx templates and Moodle plug-in';
$string['xlsx_addcategory_supertip'] = 'Category names use the Heading 1 style';
$string['xlsx_currentquestion'] = ' (Current Question)';
$string['xlsx_gapselect_screentip'] = 'Warning: customised Select missing xlsxs Moodle plugin required for this question type.';
$string['xlsx_import'] = 'Import';
$string['xlsx_multiple_answer'] = 'Multiple answer';
$string['xlsx_new_question_file'] = 'New Question File';
$string['xlsx_new_question_file_screentip'] = 'Questions must be saved in Xlsx 2008 (.xlsx) format';
$string['xlsx_new_question_file_supertip'] = 'Each Xlsx file may contain multiple categories';
$string['xlsx_setunset_assessment_view'] = 'Set/Unset Assessment View';
$string['xlsx_showhide_assessment_screentip'] = 'Show question metadata to edit, hide to preview printed assessment';
$string['xlsx_showhide_assessment_supertip'] = 'Shows or hides the hidden text';
$string['xlsx_showhide_assessment_view'] = 'Show/Hide Assessment View';
$string['xlsx_shuffle_screentip'] = 'Shuffle the answers to MCQ/TF/MA questions';
$string['xlsx_shuffle_supertip'] = 'A few shuffles is better than 1';
$string['xlsx_view'] = 'View';
