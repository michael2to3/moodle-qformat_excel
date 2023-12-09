<?php

defined('MOODLE_INTERNAL') || die();

require_once "$CFG->libdir/xmlize.php";
require_once "$CFG->dirroot/lib/uploadlib.php";
require_once "$CFG->dirroot/question/format/xml/format.php";
require_once "$CFG->dirroot/lib/excellib.class.php";
use moodle_exception;

require_once "$CFG->libdir/phpspreadsheet/vendor/autoload.php";
use PhpOffice\PhpSpreadsheet\IOFactory;

class qformat_xlsxtable extends qformat_default
{
    private $lessonquestions = [];


    public function provide_import()
    {
        return true;

    }


    public function provide_export()
    {
        return true;

    }


    public function mime_type()
    {
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    }


    public function validate_file(stored_file $file): string
    {
        if (!preg_match('#\.xlsx$#i', $file->get_filename())) {
            return get_string('errorfilenamemustbexlsx', 'qformat_xlsxtable');
        }

        return '';

    }


    public function readquestions($data)
    {
        $allQuestions = [];
        foreach ($data['sheets'] as $sheetIndex => $sheetData) {
            $questions = $sheetData['rows'];
            $images = $sheetData['images'] ?? [];

            foreach ($questions as $i => $question) {
                if (count($question) < 3) {
                    debugging('Skipping question '.$i.' in sheet due to insufficient data', DEBUG_DEVELOPER);
                    continue;
                }

                $name         = $question[0];
                $questiontext = $question[1];
                $answer       = $question[2];
                if (empty($name) || empty($questiontext) || empty($answer)) {
                    debugging('Skipping question '.$i.' in sheet due to empty fields', DEBUG_DEVELOPER);
                    continue;
                }

                if (isset($images[0])) {
                    $imageData = $images[0];
                    $imageName = $imageData['name'];
                    $filePath = $imageData['path'];
                    $fileData = $imageData['base64'];
                    $mimeType = $imageData['mimetype'];

                    $questiontext .= '<img title="' . $imageName . '" src="data:' . $mimeType . ';base64,' . $fileData . '"/>';
          throw new moodle_exception(base64_encode(print_r($images, true)));
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
                    ],
                ];

                $q->fraction = [1];
                $q->answer   = [$answer];
                $allQuestions[] = $q;
            }
        }

        debugging('All questions: '.print_r($allQuestions, true), DEBUG_DEVELOPER);
        return $allQuestions;
    }


    public function export_file_extension()
    {
        return '.xlsx';

    }


    public function writequestion($question)
    {
        $this->lessonquestions[] = $question;
        return true;

    }


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

        // $imagestring = "";
        // foreach ($imagesforzipping as $imagename => $imagedata) {
        // $filetype = strtolower(pathinfo($imagename, PATHINFO_EXTENSION));
        // $base64data = base64_encode($imagedata);
        // $filedata = 'data:image/' . $filetype . ';base64,' . $base64data;
        // Embed the image name and data into the HTML.
        // $imagestring .= '<img title="' . $imagename . '" src="' . $filedata . '"/>';

    }

    public function readdata($filename)
    {
        if ($filename === null || !preg_match('#\.xlsx$#i', $filename)) {
            return false;
        }

        $reader = IOFactory::createReader('Xlsx');
        $spreadsheet = $reader->load($filename);

        $data = [];
        $sheetCount = $spreadsheet->getSheetCount();
        for ($sheetIndex = 0; $sheetIndex < $sheetCount; $sheetIndex++) {
            $worksheet = $spreadsheet->getSheet($sheetIndex);
            $sheetData = [];
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
                    $sheetData['rows'][] = $rowData;
                }
            }

            $drawings = $worksheet->getDrawingCollection();
            foreach ($drawings as $drawing) {
                $imageData                = [];
                $imagePath                = $drawing->getPath();
                $imageData['path']        = $imagePath;
                $imageData['coordinates'] = $drawing->getCoordinates();
                $imageData['name']        = $drawing->getName();
                $imageData['base64']      = base64_encode(file_get_contents($imagePath));
                $mimeType = 'image/png';
                if (function_exists('mime_content_type')) {
                    $mimeType = mime_content_type($imagePath);
                }
                $imageData['mimetype'] = $mimeType;

                $sheetData['images'][] = $imageData;
            }

            $data['sheets'][] = $sheetData;
        }

        return $data;
    }



}
