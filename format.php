<?php

defined('MOODLE_INTERNAL') || die();

require_once("$CFG->libdir/xmlize.php");
require_once("$CFG->dirroot/lib/uploadlib.php");
require_once("$CFG->dirroot/question/format/xml/format.php");
require_once("$CFG->dirroot/lib/excellib.class.php");

class qformat_xlsxtable extends qformat_xml
{
    /** @var array Lesson questions are stored here if importing a lesson Xlsx file. */
    private $lessonquestions = array();

    /**
     * Define required MIME-Type
     *
     * @return string MIME-Type
     */
    public function mime_type()
    {
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }

    /**
     * Validate the given file.
     *
     * Check that the file has a .docx suffix, should also check it's in Zip file format.
     *
     * @param stored_file $file the file to check
     * @return string the error message that occurred while validating the given file
     */
    public function validate_file(stored_file $file): string
    {
        if (!preg_match('#\.xlsx$#i', $file->get_filename())) {
            return get_string('errorfilenamemustbexlsx', 'qformat_xlsxtable');
        }
        return '';
    }


    // IMPORT FUNCTIONS START HERE.

    /**
     * Perform required pre-processing, i.e. convert Word file into Moodle Question XML
     *
     * Extract the WordProcessingML XML files from the .docx file, and use a sequence of XSLT
     * steps to convert it into Moodle Question XML
     *
     * @return bool Success
     */
    public function importpreprocess()
    {
        global $CFG;
        $realfilename = "";
        $filename = "";

        // Handle question imports in Lesson module by using mform, not the question/format.php qformat_default class.
        if (property_exists('qformat_default', 'realfilename')) {
            $realfilename = $this->realfilename;
        } else {
            global $mform;
            $realfilename = $mform->get_new_filename('questionfile');
        }
        if (property_exists('qformat_default', 'filename')) {
            $filename = $this->filename;
        } else {
            global $mform, $USER;

            if (property_exists('qformat_default', 'importcontext')) {
                // We have to check if this request is made from the lesson interface.
                $cm = get_coursemodule_from_id('lesson', $this->importcontext->instanceid);
                if ($cm) {
                    $draftid = optional_param('questionfile', '', PARAM_FILE);
                    $dir = make_temp_directory('forms');
                    $tempfile = tempnam($dir, 'tempup_');

                    $fs = get_file_storage();
                    $context = context_user::instance($USER->id);
                    if (!$files = $fs->get_area_files($context->id, 'user', 'draft', $draftid, 'id DESC', false)) {
                        throw new \moodle_exception(get_string('cannotwritetotempfile', 'qformat_xlsxtable', ''));
                    }
                    $file = reset($files);

                    $filename = $file->copy_content_to($tempfile);
                    $filename = $tempfile;
                } else {
                    $filename = "{$CFG->tempdir}/questionimport/{$realfilename}";
                }
            } else {
                $filename = "{$CFG->tempdir}/questionimport/{$realfilename}";
            }
        }
        $basefilename = basename($filename);
        $baserealfilename = basename($realfilename);


        if (!preg_match('#\.xlsx$#i', $realfilename)) {
            throw new moodle_exception(get_string('errorfilenamemustbexlsx', 'qformat_xlsxtable'));
        }

        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
        $spreadsheet = $reader->load($filename);

        $data = [];
        foreach ($spreadsheet->getActiveSheet()->getRowIterator() as $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);
            $rowData = [];
            foreach ($cellIterator as $cell) {
                $rowData[] = $cell->getValue();
            }
            $data[] = $rowData;
        }

        foreach ($data as $rowData) {
            $questionText = array_shift($rowData);
            $question = new stdClass();
            $question->qtype = 'multichoice';
            $question->questiontext = $questionText;
            $question->options = new stdClass();
            $question->options->answers = [];

            // Assuming the second cell in each row contains an image path
            $imagePath = array_shift($rowData);
            if (!empty($imagePath) && file_exists($imagePath)) {
                // Process the image file
                $question->image = base64_encode(file_get_contents($imagePath));
                // Include logic to attach the image to the question
            }

            foreach ($rowData as $cellData) {
                $answer = new stdClass();
                $answer->answer = $cellData;
                $answer->fraction = 0;
                $question->options->answers[] = $answer;
            }

            $this->add_question($question);
        }

        return true;
    }

    // EXPORT FUNCTIONS START HERE.

    /**
     * Use a .xlsx file extension when exporting, so that Excel is used to open the file
     * @return string file extension
     */
    public function export_file_extension()
    {
        return ".xlsx";
    }

    public function presave_process($content)
    {
        if (strpos($content, "</question>") === false) {
            throw new moodle_exception(get_string('noquestions', 'qformat_xlsxtable'));
        }

        $workbook = new MoodleExcelWorkbook($this->filename);
        $worksheet = $workbook->add_worksheet('Questions');

        foreach ($content as $rowIndex => $question) {
            $worksheet->write_string($rowIndex, 0, $questionText);
        }

        $workbook->close();
        return true;
    }


    public function readdata($filename)
    {
        if (property_exists('qformat_default', 'importcontext')) {
            $cm = get_coursemodule_from_id('lesson', $this->importcontext->instanceid);
            if ($cm) {
                return $this->lessonquestions;
            }
        }

        if (!preg_match('#\.xlsx$#i', $filename)) {
            return false;
        }

        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
        $spreadsheet = $reader->load($filename);
        $worksheet = $spreadsheet->getActiveSheet();

        $data = [];
        foreach ($worksheet->getRowIterator() as $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);
            $rowData = [];
            foreach ($cellIterator as $cell) {
                $rowData[] = $cell->getValue();
            }
            $data[] = $rowData;
        }

        return $data;
    }
}
