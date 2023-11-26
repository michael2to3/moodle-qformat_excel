<?php

defined('MOODLE_INTERNAL') || die();

require_once "$CFG->libdir/xmlize.php";
require_once "$CFG->dirroot/lib/uploadlib.php";
require_once "$CFG->dirroot/question/format/xml/format.php";
require_once "$CFG->dirroot/lib/excellib.class.php";

class qformat_xlsxtable extends qformat_xml
{
    private $lessonquestions = [];


    public function mime_type()
    {
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    }//end mime_type()


    public function validate_file(stored_file $file): string
    {
        if (!preg_match('#\.xlsx$#i', $file->get_filename())) {
            return get_string('errorfilenamemustbexlsx', 'qformat_xlsxtable');
        }

        return '';

    }//end validate_file()


    public function importpreprocess()
    {

    }//end importpreprocess()


    public function export_file_extension()
    {
        return '.xlsx';

    }//end export_file_extension()


    public function writequestion($question)
    {
        $this->lessonquestions[] = $question;
        return true;

    }//end writequestion()


    public function presave_process($content)
    {
        if (count($this->lessonquestions) == 0) {
            throw new moodle_exception('noquestions', 'qformat_xlsxtable');
        }

        $workbook  = new MoodleExcelWorkbook($this->filename);
        $worksheet = $workbook->add_worksheet('Questions');
        $answers   = $question->options->answers;
        foreach ($this->lessonquestions as $rowIndex => $question) {
            $worksheet->write($rowIndex, 0, $question->questiontext);
            foreach ($answers as $answer) {
                $worksheet->write($rowIndex, 1, $answer->answer);
            }
        }

        $workbook->close();
        return true;

    }//end presave_process()


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

        $reader      = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
        $spreadsheet = $reader->load($filename);
        $worksheet   = $spreadsheet->getActiveSheet();

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

    }//end readdata()


}//end class
