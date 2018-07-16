<?php

namespace Racklin\ExcelGenerator;

/**
 * Class ExcelGenerator
 *
 * @package Racklin\ExcelGenerator
 */
class ExcelGenerator
{
    protected $stEngine = null;

    public function __construct()
    {
        $this->stEngine = new \StringTemplate\Engine;
    }


    /**
     * Generate EXCEL xlsx
     *
     * @param $template
     * @param $data
     * @param $name
     * @param $desc 'I', 'F', 'FI'
     */
    public function generate($template, $data, $name = '', $desc = 'I')
    {
        $settings = json_decode(file_get_contents($template), true);
        $phpExcel = $this->initPHPExcel($settings);

        foreach ($settings['sheets'] as $sheetIndex => $sheetSettings) {
            $activeSheet = $this->initActiveSheet($phpExcel, $sheetIndex, $sheetSettings);
            $dataArray = [];

            // WTF?

            $dataArrayVariable = $sheetSettings['info']['data_array_variable'];
            if (is_array($data[$dataArrayVariable])) {
                $dataArray = $data[$dataArrayVariable];
            }

            $rowIndex = 1;

            // process table title
            $maxHeaderCol = count($sheetSettings['header']);
            $titleColNameStart = $this->getExcelColumnName(1) . $rowIndex;
            $titleColNameEnd = $this->getExcelColumnName($maxHeaderCol) . $rowIndex;

            // merge title column
            $activeSheet->mergeCells($titleColNameStart.":".$titleColNameEnd);
            $activeSheet->setCellValue($titleColNameStart, $sheetSettings['info']['table_title']);

            // set cell style
            if (isset($sheetSettings['info']['style']) && is_array($sheetSettings['info']['style'])) {
                $activeSheet->getStyle($titleColNameStart)->applyFromArray(array_merge(ExcelDefaultStyle::$TABLE_TITLE, $sheetSettings['info']['style']));
            } else {
                $activeSheet->getStyle($titleColNameStart)->applyFromArray(ExcelDefaultStyle::$TABLE_TITLE);
            }

            // set row height
            $rowHeight = isset($sheetSettings['info']['title_row_height']) ? $sheetSettings['info']['title_row_height'] : ExcelDefaultStyle::$TITLE_ROW_HEIGHT;
            $activeSheet->getRowDimension($rowIndex)->setRowHeight($rowHeight);

            // process header
            $rowIndex++;
            foreach ($sheetSettings['header'] as $headerIndex => $headerSettings) {
                $colName = $this->getExcelColumnName($headerIndex+1) . $rowIndex;
                $activeSheet->setCellValue($colName, $headerSettings['title']);

                // set cell style
                if (isset($headerSettings['style']) && is_array($headerSettings['style'])) {
                    $activeSheet->getStyle($colName)->applyFromArray(array_merge(ExcelDefaultStyle::$HEADER, $headerSettings['style']));
                } else {
                    $activeSheet->getStyle($colName)->applyFromArray(ExcelDefaultStyle::$HEADER);
                }

                // column width
                if (isset($headerSettings['width'])) {
                    $activeSheet->getColumnDimension($this->getExcelColumnName($headerIndex+1))->setWidth($headerSettings['width']);
                }
            }
            // set row height
            $rowHeight = isset($sheetSettings['info']['header_row_height']) ? $sheetSettings['info']['header_row_height'] : ExcelDefaultStyle::$HEADER_ROW_HEIGHT;
            $activeSheet->getRowDimension($rowIndex)->setRowHeight($rowHeight);


            // process data
            foreach ($dataArray as $dataValue) {
                $rowIndex++;
                foreach ($sheetSettings['data'] as $dataIndex => $dataSettings) {
                    $colName = $this->getExcelColumnName($dataIndex+1) . $rowIndex;
                    $val = $this->renderText($dataSettings['value'], $dataValue);
                    $activeSheet->setCellValue($colName, $val);

                    // set cell style
                    if (isset($dataSettings['style']) && is_array($dataSettings['style'])) {
                        $activeSheet->getStyle($colName)->applyFromArray(array_merge(ExcelDefaultStyle::$DATA, $dataSettings['style']));
                    } else {
                        $activeSheet->getStyle($colName)->applyFromArray(ExcelDefaultStyle::$DATA);
                    }
                }
                // set row height
                $rowHeight = isset($sheetSettings['info']['data_row_height']) ? $sheetSettings['info']['data_row_height'] : ExcelDefaultStyle::$DATA_ROW_HEIGHT;
                $activeSheet->getRowDimension($rowIndex)->setRowHeight($rowHeight);
            }

            // set table borders
            $activeSheet->getStyle('A1:'.$this->getExcelColumnName($maxHeaderCol).$rowIndex)->applyFromArray(ExcelDefaultStyle::$TABLE_BORDER);
        }

        // save or output
        $this->write($phpExcel, $name, $desc);
    }


    /**
     * @param $settings
     * @return \PHPExcel
     */
    protected function initPHPExcel($settings)
    {
        $phpExcel = new \PHPExcel();

        // set document information
        $phpExcel->getProperties()->setCreator($settings['info']['creator']);
        $phpExcel->getProperties()->setTitle($settings['info']['title']);
        $phpExcel->getProperties()->setSubject($settings['info']['subject']);
        $phpExcel->getProperties()->setKeywords($settings['info']['keywords']);

        return $phpExcel;
    }


    /**
     * @param \PHPExcel $phpExcel
     * @param $sheetIndex
     * @param $sheetSettings
     * @return \PHPExcel_Worksheet
     */
    protected function initActiveSheet(\PHPExcel $phpExcel, $sheetIndex, $sheetSettings)
    {
        $phpExcel->setActiveSheetIndex($sheetIndex);
        $activeSheet = $phpExcel->getActiveSheet();

        // set sheet information
        $activeSheet->setTitle($sheetSettings['info']['title']);

        return $activeSheet;
    }


    /**
     * @param $columnNumber
     * @return string
     */
    protected function getExcelColumnName($columnNumber)
    {
        $dividend = $columnNumber;
        $columnName = "";
        $modulo = 0;

        while ($dividend > 0) {
            $modulo = ($dividend - 1) % 26;
            $columnName = chr(65 + $modulo) . $columnName;
            $dividend = (int)(($dividend - $modulo) / 26);
        }

        return $columnName;
    }


    /**
     * @param $template
     * @param $data
     * @return mixed|string
     */
    protected function renderText($template, $data)
    {
        $text = $this->stEngine->render($template, $data);
        // empty undefined variable
        $text = preg_replace("/{[\w.]+}/", "", $text);
        return $text;
    }


    /**
     * @param \PHPExcel $phpExcel
     * @param string $name
     * @param string $desc
     */
    protected function write(\PHPExcel $phpExcel, $name='sheet.xlsx', $desc='I')
    {
        $writer = \PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');

        switch ($desc) {
            default:
            case 'I':
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Cache-Control: private, must-revalidate, post-check=0, pre-check=0, max-age=1');
                //header('Cache-Control: public, must-revalidate, max-age=0'); // HTTP/1.1
                header('Pragma: public');
                header('Expires: Sat, 26 Jul 1997 05:00:00 GMT'); // Date in the past
                header('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT');
                header('Content-Disposition: attachment; filename="'.basename($name).'.xlsx"');
                $writer->save("php://output");
                break;

            case 'F':
                $writer->save($name);
                break;

            case 'FI':
                $writer->save($name);

                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Cache-Control: private, must-revalidate, post-check=0, pre-check=0, max-age=1');
                //header('Cache-Control: public, must-revalidate, max-age=0'); // HTTP/1.1
                header('Pragma: public');
                header('Expires: Sat, 26 Jul 1997 05:00:00 GMT'); // Date in the past
                header('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT');
                header('Content-Disposition: attachment; filename="'.basename($name).'.xlsx"');
                header('Content-Length: ' . filesize($name));
                echo file_get_contents($name);
                break;


        }
    }
}
