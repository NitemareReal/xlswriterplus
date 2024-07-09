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
