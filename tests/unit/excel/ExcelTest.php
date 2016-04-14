<?php

namespace excel;

use Codeception\Specify;
use Codeception\TestCase\Test;

class ExcelTest extends Test
{
    use Specify;

    /** @noinspection PhpUndefinedClassInspection */

    /**
     * @var \UnitTester
     */
    protected $tester;
    protected $modelClass = 'User';
    protected $savePath = '/tmp/';


    protected function _before()
    {
        $_modelClass = getenv('modelClass');
        if ($_modelClass) {
            $this->modelClass = $_modelClass;
        }
    }

    public function testClassExistence()
    {
        $excel = new \Excel();
        $this->assertInstanceOf('CComponent', $excel);
    }

    public function testOutput()
    {
        $this->specifyConfig()
            ->cloneOnly(['excel', 'path']);

        $excel = new \Excel();
        $excel->worksheet('Worksheet', ['-']);
        $path = $this->savePath;

        $this->specify('save .xls to file', function () use ($excel, $path) {
            $isExported = $excel->export('export.xls', $path);
            $this->assertTrue($isExported);
            $this->assertFileExists($path . 'export.xls');
        });

        $this->specify('save .xlsx to file', function () use ($excel, $path) {
            $isExported = $excel->export('export.xlsx', $path);
            $this->assertTrue($isExported);
            $this->assertFileExists($path . 'export.xlsx');
        });

        $this->specify('save .html to file', function () use ($excel, $path) {
            $isExported = $excel->export('export.html', $path);
            $this->assertTrue($isExported);
            $this->assertFileExists($path . 'export.html');
        });

        $this->specify('save .csv to file', function () use ($excel, $path) {
            $isExported = $excel->export('export.csv', $path);
            $this->assertTrue($isExported);
            $this->assertFileExists($path . 'export.csv');
        });

        $this->specify('save .csv to not existing dir', function () use ($excel, $path) {
            $isExported = $excel->export('export.csv', '/pmt/');
            $this->assertFalse($isExported);
            $this->assertFileNotExists('/pmt/export.csv');
        });

        $isExported = $excel->export();
        $this->assertTrue($isExported);
        $this->assertNotEmpty(ob_get_contents());
    }

    public function testWorksheets()
    {
        $arrayOfValues = [['col1', 'col2'], ['item11', 'item12'], ['item21', 'item22']];
        $arrayOfModels = \CActiveRecord::model($this->modelClass)->findAll();

        $isExported = (new \Excel)
            ->worksheet('Array of values', $arrayOfValues)
            ->worksheet('Array of values + callback', $arrayOfValues,
                function (\PHPExcel_Worksheet $worksheet, array $data) {
                    $worksheet->fromArray($data);
                }
            )
            ->worksheet('Array of models', $arrayOfModels)
            ->worksheet('Array of models + callback', $arrayOfModels,
                function (\PHPExcel_Worksheet $worksheet, array $data) {
                    $_data = [];
                    foreach ($data as $model) {
                        if ($model instanceof \CActiveRecord) {
                            $_data[] = $model->getAttributes();
                        }
                    }
                    $worksheet->fromArray($_data);
                }
            )
            ->worksheet('CArrayDataProvider of raw data', new \CArrayDataProvider($arrayOfValues))
            ->worksheet('CArrayDataProvider of models', new \CArrayDataProvider($arrayOfModels))
            ->worksheet('CActiveDataProvider + callback', new \CActiveDataProvider($this->modelClass),
                function (\PHPExcel_Worksheet $worksheet, \CActiveDataProvider $dataProvider) {
                    $_data[] = $dataProvider->model->attributeNames();
                    foreach ($dataProvider->getData() as $model) {
                        if ($model instanceof \CActiveRecord) {
                            $_data[] = $model->getAttributes();
                        }
                    }
                    $worksheet->fromArray($_data);
                }
            )
            ->export('worksheets.xlsx', $this->savePath);

        $this->assertTrue($isExported);
        $this->assertFileExists($this->savePath . 'worksheets.xlsx');
    }

}
