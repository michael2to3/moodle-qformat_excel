<?php

defined('MOODLE_INTERNAL') || die();

require_once "$CFG->libdir/xmlize.php";
require_once "$CFG->dirroot/lib/uploadlib.php";
require_once "$CFG->dirroot/question/format/xml/format.php";
require_once "$CFG->dirroot/lib/excellib.class.php";
use moodle_exception;

class qformat_xlsxtable extends qformat_default
{
    private $lessonquestions = [];


    public function provide_import()
    {
        return true;

    }//end provide_import()


    public function provide_export()
    {
        return true;

    }//end provide_export()


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


    public function readquestions($data)
    {
        $qa = [];
        foreach ($data as $i => $question) {
            $name         = $question[0];
            $questiontext = $question[1];
            $answer       = $question[2];
            if (empty($name) || empty($questiontext) || empty($answer)) {
                debugging('Skipping question '.$i, DEBUG_DEVELOPER);
                continue;
            }

            $q               = $this->defaultquestion();
            $q->id           = $i;
            $q->name         = $name;
            $q->questiontext = $questiontext;
            $q->qtype        = 'shortanswer';
            $q->feedback     = [
                0 => [
                    'text'   => ' ',
                    'format' => FORMAT_HTML,
                ]
            ];

            $q->fraction = [1];
            $q->answer   = [$answer];
            $qa[]        = $q;
        }//end foreach

        debugging('qa: '.print_r($qa, true), DEBUG_DEVELOPER);
        return $qa;

    }//end readquestions()


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
        foreach ($this->lessonquestions as $rowIndex => $question) {
            $worksheet->write($rowIndex, 0, $question->name);
            $worksheet->write($rowIndex, 1, $question->questiontext);
            $answers = $question->options->answers;
            foreach ($answers as $a) {
                $worksheet->write($rowIndex, 2, $a->answer);
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
            $cellIterator->setIterateOnlyExistingCells(true);
            $rowData  = [];
            $rowEmpty = true;
            foreach ($cellIterator as $cell) {
                $value = $cell->getValue();
                if (!empty($value)) {
                    $rowEmpty = false;
                }

                $rowData[] = $value;
            }

            if (!$rowEmpty) {
                $data[] = $rowData;
            }
        }

        return $data;

    }//end readdata()


}//end class
