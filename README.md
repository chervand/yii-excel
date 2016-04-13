# [PHPExcel](https://github.com/PHPOffice/PHPExcel) wrapper for [Yii Framework](https://github.com/yiisoft/yii)

It supports export for IDataProvider objects, array of models and raw data arrays to .xls, .xlsx, .html and .csv formats.

## Export

### Usage

```php
    (new Excel)
        ->worksheet('Worksheet #1', [['col1', 'col2'], ['cell11', 'cell12'], ['cell21', 'cell22']],
            function (\PHPExcel_Worksheet $worksheet, array &$data) {
                $worksheet->fromArray($data);
            }
        )
        ->scenario('export')
        ->worksheet('Worksheet #2', new \CActiveDataProvider('User'))
        ->export('/tmp/', 'export.xlsx');
```

### Output

`export()` has 2 optional arguments: 

 - save path without filename, defaults to `php://output`
 - filename with extension, defaults to `Export_{timestamp}.csv`

```php
    (new Excel)
        ->worksheet('Worksheet #1', new \CActiveDataProvider('User'))
        ->export('/tmp/', 'export.xlsx');
```

Supported formats:

 - BIFF 8 (`.xls`) Excel 95 and above
 - Office Open XML (`.xlsx`) Excel 2007 and above
 - HTML (`.html`)
 - CSV (`.csv`)
 
### Worksheets

To add a sheet to the workbook call `worksheet()` with arguments:

 - worksheet title, required
 - data to be exported, which could be a CActiveDataProvider, CArrayDataProvider, array of models or a raw data array
 - callback for custom configuration of PHPExcel_Worksheet object (see [PHPExcel](https://github.com/PHPOffice/PHPExcel) documentation), params:
    - PHPExcel_Worksheet object
    - variable passed as data to `worksheet()`

### Scenario

For retrieving safe attributes of exported models `search` scenario is used. It could be changed by calling `scenario()`.
 
```php
    (new Excel)
        ->scenario('export')
        ->worksheet('Worksheet #1', new \CActiveDataProvider('User'))
        ->export('/tmp/', 'export.xlsx');
```

### Complete example

```php
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
        ->scenario('export')
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
```

See `tests` for more examples.
