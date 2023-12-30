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
        foreach ($data['sheets'] as $sheetIndex => $sheetData) {
            $sheet          = $sheetData['rows'];
            $images         = ($sheetData['images'] ?? []);
            $name           = $this->bottom_value('_name', $sheet, 'Name not found');
            $fine           = $this->abs_val($this->bottom_value('_fine', $sheet, '0.1'));
            $penalty        = $this->abs_val($this->bottom_value('_penalty', $sheet, '0.333'));
            $fraction       = $this->abs_val($this->bottom_value('_fraction', $sheet, '1.0'));
            $qtype          = $this->bottom_value('_qtype', $sheet, 'numerical');
            $keys           = $this->process_data($sheet);
            $countQuestions = max(array_map(fn ($item) => count($item), $keys));

            debugging("Try import {$countQuestions} questions", DEBUG_DEVELOPER);

            for ($i = 0; $i < $countQuestions; $i++) {
                $qtext = $this->bottom_value('_text', $sheet, '');

                foreach ($keys as $key => $value) {
                    if ($key === '_answer') {
                        debugging('Skip _answer to be processed', DEBUG_DEVELOPER);
                        continue;
                    }

                    $qtext .= "<p><b>{$key}</b>: {$value[$i]}</p>";
                }

                foreach ($images as $image) {
                    $qtext .= "<img title=\"{$image['name']}\" src=\"data:{$image['mimetype']};base64,{$image['base64']}\"/>";
                }

                $q               = $this->defaultquestion();
                $q->id           = $i;
                $q->name         = $name;
                $q->tags         = [
                    'xlsxtable',
                    $sheetData['name'],
                ];
                $q->questiontext = $qtext;
                $q->qtype        = $qtype;
                $q->feedback     = [
                    0 => [
                        'text'   => ' ',
                        'format' => FORMAT_HTML,
                    ],
                ];

                $q->fraction  = [$fraction];
                $q->answer    = [$keys['_answer'][$i]];
                $q->tolerance = [$fine];
                $q->penalty   = $penalty;
                $qa[]         = $q;
            }//end for
        }//end foreach

        return $qa;

    }//end readquestions()


    private function abs_val($str)
    {
        if (str_contains($str, '%')) {
            return (floatval($str) / 100);
        }

        return floatval($str);

    }//end abs_val()


    private function process_data($data)
    {
        $count        = $this->get_counts($data);
        $countColumns = $count[0];
        $countRows    = $count[1];

        $keys = [];
        for ($j = 0; $j < $countColumns; $j++) {
            $keys[$data[0][$j]] = [];
        }

        for ($i = 1; $i < $countRows; $i++) {
            for ($j = 0; $j < $countColumns; $j++) {
                $keys[$data[0][$j]][] = $data[$i][$j];
            }
        }

        return $keys;

    }//end process_data()


    private function get_counts($data)
    {
        $countRows    = 0;
        $countColumns = 0;

        foreach ($data[0] as $i => $column) {
            if (empty($column)) {
                break;
            }

            $countColumns = ($i + 1);
        }

        foreach ($data as $i => $row) {
            $countEmpty = 0;
            for ($j = 0; $j < $countColumns; $j++) {
                $cell = $row[$j];
                if (empty($cell)) {
                    $countEmpty++;
                }
            }

            if ($countEmpty >= $countColumns) {
                debugging("Found empty row $countEmpty >= $countColumns, at row $i", DEBUG_DEVELOPER);
                break;
            }

            for ($j = 0; $j < $countColumns; $j++) {
                $cell = $row[$j];
                if (!empty($cell)) {
                    $countRows = ($i + 1);
                }
            }
        }//end foreach

        debugging('Columns: '.$countColumns, DEBUG_DEVELOPER);
        debugging('Rows: '.$countRows, DEBUG_DEVELOPER);
        return [
            $countColumns,
            $countRows,
        ];

    }//end get_counts()


    private function bottom_value($identifier, $data, $default = '')
    {
        foreach ($data as $i => $row) {
            foreach ($row as $j => $value) {
                if ($value === $identifier) {
                    $r = $data[($i + 1)][$j];
                    debugging("Found $identifier, bottom value: {$r}", DEBUG_DEVELOPER);
                    return $r;
                }
            }
        }

        debugging('Not found bottom value', DEBUG_DEVELOPER);
        return $default;

    }//end bottom_value()


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
        if ($filename === null || !preg_match('#\.xlsx$#i', $filename)) {
            return false;
        }

        $reader      = IOFactory::createReader('Xlsx');
        $spreadsheet = $reader->load($filename);

        $data       = [];
        $sheetCount = $spreadsheet->getSheetCount();
        for ($sheetIndex = 0; $sheetIndex < $sheetCount; $sheetIndex++) {
            $worksheet = $spreadsheet->getSheet($sheetIndex);
            $sheetData = ['name' => $worksheet->getTitle()];
            foreach ($worksheet->getRowIterator() as $row) {
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(false);
                $rowData = [];
                foreach ($cellIterator as $cell) {
                    $value     = $cell->getCalculatedValue();
                    $rowData[] = $value;
                }

                $sheetData['rows'][] = $rowData;
            }

            $drawings = $worksheet->getDrawingCollection();
            foreach ($drawings as $drawing) {
                $imageData                = [];
                $imagePath                = $drawing->getPath();
                $imageData['path']        = $imagePath;
                $imageData['coordinates'] = $drawing->getCoordinates();
                $imageData['name']        = $drawing->getName();
                $imageData['base64']      = base64_encode(file_get_contents($imagePath));
                $mimeType                 = 'image/png';
                if (function_exists('mime_content_type')) {
                    $mimeType = mime_content_type($imagePath);
                }

                $imageData['mimetype'] = $mimeType;

                $sheetData['images'][] = $imageData;
            }

            $data['sheets'][] = $sheetData;
        }//end for

        return $data;

    }//end readdata()


}//end class
