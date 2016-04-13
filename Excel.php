<?php
/**
 * @author chervand <chervand@gmail.com>
 */

/**
 * Class Excel provides phpoffice/phpexcel wrapper for Yii Framework.
 * It works with IDataProvider objects, array of models and raw data arrays.
 *
 * Export example:
 *
 *  (new Excel)
 *      ->worksheet('Worksheet #1', [['col1', 'col2'], ['cell11', 'cell12'], ['cell21', 'cell22']],
 *          function (\PHPExcel_Worksheet $worksheet, array &$data) {
 *              $worksheet->fromArray($data);
 *          }
 *      )
 *      ->scenario('export')
 *      ->worksheet('Worksheet #2', new \CActiveDataProvider('User'))
 *      ->export('/tmp/', 'export.xlsx');
 *
 * @todo: formatting support
 */
class Excel extends CComponent
{
    /**
     * Default output path.
     */
    const OUTPUT_DEFAULT = 'php://output';
    const FORMAT_XLS = '.xls';
    const FORMAT_XLSX = '.xlsx';
    const FORMAT_HTML = '.html';
    const FORMAT_CSV = '.csv';

    /**
     * @var string model scenario for retrieving safe attributes, defaults to 'search'
     */
    private $_scenario = 'search';

    /**
     * @var PHPExcel workbook object
     */
    private $_workbook;


    /**
     * Excel constructor.
     * Initializes a workbook and removes default worksheet.
     */
    public function __construct()
    {
        $this->_workbook = new PHPExcel();
        $this->_workbook->removeSheetByIndex(0);
    }

    /**
     * Adds a new worksheet from a data array or a CDataProvider object.
     *
     * @param string $title
     * @param array|IDataProvider $data
     * @param null|callable $callback PHPExcel worksheet render callback
     *
     * @see https://github.com/PHPOffice/PHPExcel
     *
     * @return $this
     * @throws PHPExcel_Exception
     *
     * @todo: worksheet available characters and max length
     */
    public function worksheet($title, $data, $callback = null)
    {
        $workbook = &$this->_workbook;
        $worksheet = new PHPExcel_Worksheet($this->_workbook, $title);

        if (!is_callable($callback)) {
            $callback = [$this, 'defaultCallback'];
        }

        try {
            call_user_func($callback, $worksheet, $data);
            $workbook->addSheet($worksheet);
        } catch (PHPExcel_Exception $e) {
            throw $e;
        }

        return $this;
    }

    /**
     * Exports the workbook.
     *
     * @param null|string $filename defaults to Export_{timestamp}.csv,
     * @param null|string $path output path without filename
     * if extension is not .xsl, .xlsx, .html, .csv it defaults to .csv
     *
     * @return bool
     */
    public function export($filename = null, $path = self::OUTPUT_DEFAULT)
    {
        if (!is_string($filename)) {
            $filename = 'Export_' . time();
        }

        if (is_string($path) && $path != self::OUTPUT_DEFAULT) {
            $path .= $filename;
        }

        $format = '.' . end(explode('.', $filename));
        if (!in_array($format, [
            self::FORMAT_XLS,
            self::FORMAT_XLSX,
            self::FORMAT_HTML,
            self::FORMAT_CSV,
        ])
        ) {
            $format = self::FORMAT_CSV;
            $filename .= $format;
        }

        return $this->output($this->_workbook, $path, $filename, $format);
    }

    protected function output(&$workbook, $path, $filename, $format)
    {
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

        switch ($format) {
            case self::FORMAT_XLS:
                header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
                $writer = PHPExcel_IOFactory::createWriter($workbook, 'Excel5');
                break;
            case self::FORMAT_XLSX:
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=UTF-8');
                $writer = PHPExcel_IOFactory::createWriter($workbook, 'Excel2007');
                break;
            case self::FORMAT_HTML:
                header('Content-type: text/html; charset=UTF-8');
                $writer = PHPExcel_IOFactory::createWriter($workbook, 'HTML');
                break;
            case self::FORMAT_CSV:
            default:
                header('Content-type: text/csv; charset=UTF-8');
                $writer = PHPExcel_IOFactory::createWriter($workbook, 'CSV');
        }

        try {
            $writer->save($path);
        } catch (Exception $e) {
            return false;
        }

        return true;
    }

    /**
     * @param $scenario
     * @return $this
     */
    public function scenario($scenario)
    {
        $this->_scenario = $scenario;
        return $this;
    }

    /**
     * Default callback worksheet which simply exports raw data without any formatting.
     *
     * @param PHPExcel_Worksheet $worksheet
     * @param array|IDataProvider $dataProvider
     *
     * @return PHPExcel_Worksheet
     * @throws PHPExcel_Exception
     */
    protected function defaultCallback(\PHPExcel_Worksheet $worksheet, $dataProvider)
    {
        $_data = [];

        if ($dataProvider instanceof \CActiveDataProvider) {
            $_data[] = $dataProvider->model->attributeNames();
            foreach ($dataProvider->getData() as $model) {
                if ($model instanceof \CActiveRecord) {
                    $model->setScenario($this->_scenario);
                    $_names = $model->getSafeAttributeNames();
                    $_data[] = $model->getAttributes($_names);
                }
            }
            return $worksheet->fromArray($_data);
        }

        if ($dataProvider instanceof \IDataProvider) {
            $_data = &$dataProvider->getData();
        } elseif (is_array($dataProvider)) {
            $_data = $dataProvider;
        } else {
            $_data = [$dataProvider];
        }

        foreach ($_data as $index => $value) {
            if ($value instanceof \CActiveRecord) {
                $value->setScenario($this->_scenario);
                $_names = $value->getSafeAttributeNames();
                $_data[$index] = $value->getAttributes($_names);
            } elseif (
                !is_array($value)
                && !is_scalar($value)
                && !is_null($value)
            ) {
                unset($_data[$index]);
            }
        }

        return $worksheet->fromArray($_data);
    }
}
