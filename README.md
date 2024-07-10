# xlswriterplus
This class EXTENDS XLSXWriter for PHP

The pourpose of this class is extending https://github.com/mk-j/PHP_XLSXWriter to allow inserting images in a sheet

New method:
~~~php
public function addImage($sheet_name, $imagePath, $imageName, $imageId, $imageOptions = [])

    string $sheet_name - Name of the sheet in which to insert the image
    string $imagePath - Path to image
    string $imageName - Image name (not used, only for backward compatibility)
    int $imageId - Image unique identifier 
    array $imageOptions - Options for image, all optional: startColNum=0, startRowNum=0, endColNum=0, endRowNum=0. If not end coordinates given, image will keep its original size
~~~
Simple example:
~~~php
include_once("xlsxwriter.class.php");
include_once("xlsxwriterplus.php");

$data = array(
    array('year','month','amount'),
    array('2003','1','220'),
    array('2003','2','153.5'),
);

$writer = new XLSWriterPlus()

$writer->writeSheet($data);
$writer->addImage('Sheet1', realpath('./media/logo1.jpeg'), '', 1, [ 'startColNum' => 4, 'startRowNum' => 2, ]);
$writer->writeToFile('output_simple.xlsx');
~~~
Advanced example:
~~~php
include_once("xlsxwriter.class.php");
include_once("xlsxwriterplus.php");

$data = array(
    array('year','month','amount'),
    array('2003','1','220'),
    array('2003','2','153.5'),
);

$data2 = array(
    array('year','month','amount'),
    array('2024','1','223'),
    array('2024','2','198.7'),
    array('2024','3','221.6'),
    array('2024','4','209'),
);

$writer = new XLSWriterPlus()

$writer->writeSheet($data);
$writer->addImage('Sheet1', realpath('./media/logo1.jpeg'), '', 1, [ 'startColNum' => 4, 'startRowNum' => 2, ]);
$writer->addImage('Sheet1', realpath('./media/logo1.jpeg'), '', 2, [ 'startColNum' => 12, 'startRowNum' => 4 ]);
$writer->addImage('Sheet1', realpath('./media/logo2.png'), '', 3, [ 'startColNum' => 1, 'startRowNum' => 12 ]);
// adding data to a new sheet
$writer->writeSheet($data2, 'Sheet2');
$writer->addImage('Sheet2', realpath('./media/logo2.png'), '', 4, [ 'startColNum' => 2, 'startRowNum' => 4 ]);
$writer->writeToFile('output_advanced.xlsx');
~~~
