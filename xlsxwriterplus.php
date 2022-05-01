<?php

class XLSWriterPlus extends XLSXWriter
{
    /**
     * @var array
     */
    private $images = [];

    /**
     * @var array
     */
    private $imageOptions = [];

    /**
     * @var array
     */
    private $ignoredErrorsCells = [];

	protected $DPIx = 0;
	protected $DPIy = 0;
    /**
     * @return array
     */
    public function getIgnoredErrorsCells()
    {
        return $this->ignoredErrorsCells;
    }

    /**
     * @param array $ignoredErrorsCells
     */
    public function setIgnoredErrorsCells($ignoredErrorsCells)
    {
        $this->ignoredErrorsCells = $ignoredErrorsCells;
    }

    /**
     * @return array
     */
    public function getSheets()
    {
        return $this->sheets;
    }

    /**
     * @param string $sheet_name - Name of the sheet in which to insert the image
	 * @param string $imagePath - Path to image
     * @param string $imageName - Image name (not used, only for backward compatibility)
     * @param int $imageId - Image unique identifier 
     * @param array $imageOptions - Options for image, all optional: startColNum=0, startRowNum=0, endColNum=0, endRowNum=0. If not end coordinates given, image will keep its original size
     * @throws Exception
     */
    public function addImage($sheet_name, $imagePath, $imageName, $imageId, $imageOptions = [])
    {
		if (empty($sheet_name))
			return;

        if(!file_exists($imagePath)){
            throw new \Exception(sprintf('File %s not found.', $imagePath));
        }

		$this->initializeSheet($sheet_name);

        $this->images[$sheet_name][$imageId] = array(
            'path' => $imagePath,
			'name' => $imageName,
			//'uniq_name' => uniqid() . '.' . pathinfo($imagePath, PATHINFO_EXTENSION),
			'uniq_name' => sha1_file($imagePath) . '.' . pathinfo($imagePath, PATHINFO_EXTENSION),	// "sha1_file" allows to reuse image files
        );

        $this->imageOptions[$sheet_name][$imageId] = array_merge([
            'startColNum' => 0,
            'endColNum' => 0,
            'startRowNum' => 0,
			'endRowNum' => 0,
		], $imageOptions);

		$this->current_sheet = $sheet_name;
    }

    public function writeToString()
    {
        $temp_file = $this->tempFilename();
        $this->writeToFile($temp_file);
        $string = file_get_contents($temp_file);

        return $string;
    }

    /**
     * @param string $filename
     * @throws Exception
     */
    public function writeToFile($filename)
    {
        foreach ($this->sheets as $sheet_name => $sheet) {
            $this->finalizeSheet($sheet_name);
        }

        if (file_exists($filename)) {
            if (is_writable($filename)) {
                @unlink($filename);
            } else {
                throw new \Exception("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", file is not writeable.");
            }
        }

        $zip = new \ZipArchive();
        if (empty($this->sheets)) {
            throw new \Exception("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", no worksheets defined.");
        }
        if (!$zip->open($filename, \ZipArchive::CREATE)) {
            throw new \Exception("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", unable to create zip.");
        }

		$zip->addEmptyDir("docProps/");
		$zip->addFromString("docProps/app.xml" , $this->buildAppXML() );
		$zip->addFromString("docProps/core.xml", $this->buildCoreXML());

        $zip->addEmptyDir("_rels/");
        $zip->addFromString("_rels/.rels", $this->buildRelationshipsXML());

		if (\count($this->images) > 0) {	// if there is any image, on any sheet
			$img = \imagecreatetruecolor(100, 100); // a dummy image, just to get system DPIs
			list($this->DPIx, $this->DPIy) = \imageresolution($img);

            $zip->addEmptyDir("xl/media/");

            $zip->addEmptyDir("xl/drawings");
            $zip->addEmptyDir("xl/drawings/_rels");

			foreach ($this->images as $sheet_name => $images){
				foreach($images as $imageId => $image) {
					$zip->addFile($image['path'], 'xl/media/' . $image['uniq_name']);
				}
			}
            $i = 1;
			foreach ($this->sheets as $sheet_name => $sheet) {
				if(isset($this->images[$sheet_name])){
					$zip->addFromString("xl/drawings/drawing" . $i . ".xml", $this->buildDrawingXML($sheet_name));
					$zip->addFromString("xl/drawings/_rels/drawing" . $i . ".xml.rels", $this->buildDrawingRelationshipXML($sheet_name));
				}
				$i++;
			}
        }

        $zip->addEmptyDir("xl/worksheets/");
        $zip->addEmptyDir("xl/worksheets/_rels/");

        $i = 1;
        foreach ($this->sheets as $sheet_name => $sheet) {
			$zip->addFile($sheet->filename, "xl/worksheets/" . $sheet->xmlname);
			if(isset($this->images[$sheet_name]) && \count($this->images[$sheet_name]) > 0){	// Drawing relationship only needed if there is any picture in this sheet
				$zip->addFromString("xl/worksheets/_rels/" . $sheet->xmlname . '.rels', $this->buildSheetRelationshipXML($i));
			}
			$i++;
        }

        $zip->addFromString("xl/workbook.xml", $this->buildWorkbookXML());
        $zip->addFile($this->writeStylesXML(), "xl/styles.xml");
        $zip->addFromString("[Content_Types].xml", $this->buildContentTypesXML());

        $zip->addEmptyDir("xl/_rels/");
        $zip->addFromString("xl/_rels/workbook.xml.rels", $this->buildWorkbookRelsXML());
        $zip->close();
    }

    /**
	 * @param string $sheet_name - Name of the sheet
     * @return string
     */
    public function buildDrawingXML($sheet_name)
    {
        $imageRelationshipXML = '';

        if(\count($this->images[$sheet_name]) > 0) {

            $imageRelationshipXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>' . "\n" . 
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
        }
		$i = 1;
        foreach($this->images[$sheet_name] as $imageId => $image) {         
            $imageOptions = $this->imageOptions[$sheet_name][$imageId];
			
			$oneCellAnchor = false;
			$anchorType = "twoCellAnchor";
			
            if($imageOptions['endColNum'] == 0 && $imageOptions['endRowNum'] == 0) {	// if no end coordinates (row and col cells) given
				$oneCellAnchor = true;													// it is a "one cell anchor" image, just top left cells
				$anchorType = "oneCellAnchor";
				list($imageWidth, $imageHeight) = \getimagesize($image['path']);		// get image width and height in pixels
				$EMUx = $imageWidth * 914400 / $this->DPIx;								// conversion from pixels to EMUs for "xdr:ext" tag, used instead of "xdr:to"
				$EMUy = $imageHeight * 914400 / $this->DPIy;
			}
                
            $endColOffset = 0; //round($width * 1);
            $endRowOffset = 0; //round($height * 1);

            $imageRelationshipXML .= '
	<xdr:' . $anchorType;
			if(!$oneCellAnchor){
				$imageRelationshipXML .= ' editAs="oneCell"';
			}
			$imageRelationshipXML .= '>
		<xdr:from>
			<xdr:col>' . $imageOptions['startColNum'] . '</xdr:col>
			<xdr:colOff>0</xdr:colOff>
			<xdr:row>' . $imageOptions['startRowNum'] . '</xdr:row>
			<xdr:rowOff>0</xdr:rowOff>
		</xdr:from>';
			if($oneCellAnchor){
				$imageRelationshipXML .= '
		<xdr:ext cx="' . $EMUx . '" cy="' . $EMUy .'"/>';
			}else {
				$imageRelationshipXML .= '
		<xdr:to>
			<xdr:col>' . $imageOptions['endColNum'] . '</xdr:col>
			<xdr:colOff>' . $endColOffset . '</xdr:colOff>
			<xdr:row>' . $imageOptions['endRowNum'] . '</xdr:row>
			<xdr:rowOff>' . $endRowOffset . '</xdr:rowOff>
		</xdr:to>';
			}
			$imageRelationshipXML .= '
		<xdr:pic>
			<xdr:nvPicPr>
				<xdr:cNvPr id="' . $imageId . '" name="Picture ' . $imageId . '">
					<a:extLst>
						<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
							<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" id="{D536D061-A3D2-4F2B-ACAF-CD70361876FA}"/>
						</a:ext>
					</a:extLst>
				</xdr:cNvPr>
				<xdr:cNvPicPr>
					<a:picLocks noChangeAspect="1" />
				</xdr:cNvPicPr>
			</xdr:nvPicPr>
			<xdr:blipFill>
				<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' . $i . '"/>
				<a:stretch/>
			</xdr:blipFill>
			<xdr:spPr>
				<a:prstGeom prst="rect">
					<a:avLst/>
				</a:prstGeom>
			</xdr:spPr>
		</xdr:pic>
		<xdr:clientData/>
	</xdr:' . $anchorType . '>';
			$i++;
        }
        
        if($imageRelationshipXML != '') {
            $imageRelationshipXML .= "\n" . '</xdr:wsDr>';
        }

        return $imageRelationshipXML;
    }

    /**
     * @return string
     */
    protected function buildContentTypesXML()
    {
        $content_types_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $content_types_xml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $content_types_xml .= '<Default Extension="jpeg" ContentType="image/jpeg"/>';
        $content_types_xml .= '<Default Extension="png" ContentType="image/png"/>';
        $content_types_xml .= '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $content_types_xml .= '<Default Extension="xml" ContentType="application/xml"/>';

		$content_types_drawing_xml = "";
		$i = 1;
        foreach ($this->sheets as $sheet_name => $sheet) {
			$content_types_xml .= '<Override PartName="/xl/worksheets/' . ($sheet->xmlname) . '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
			if(isset($this->images[$sheet_name]) && \count($this->images[$sheet_name]) > 0){
				$content_types_drawing_xml .= '<Override PartName="/xl/drawings/drawing' . $i . '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>';
			}
			$i++;
        }
        $content_types_xml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
		$content_types_xml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
		$content_types_xml .= $content_types_drawing_xml;
// 		$content_types_xml .= "\n";
		$content_types_xml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
		$content_types_xml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
		$content_types_xml .= '</Types>';

        return $content_types_xml;
    }

    /**
     * @param string $sheet_name - Name of the sheet to finalize
     */
    protected function finalizeSheet($sheet_name)
    {
        if (empty($sheet_name) || $this->sheets[$sheet_name]->finalized)
            return;

        $sheet = &$this->sheets[$sheet_name];

        $sheet->file_writer->write('</sheetData>');

        if (!empty($sheet->merge_cells)) {
            $sheet->file_writer->write('<mergeCells>');
            foreach ($sheet->merge_cells as $range) {
                $sheet->file_writer->write('<mergeCell ref="' . $range . '"/>');
            }
            $sheet->file_writer->write('</mergeCells>');
		}
        
        $max_cell = self::xlsCell($sheet->row_count - 1, count($sheet->columns) - 1);

		if ($sheet->auto_filter) {
			$sheet->file_writer->write('<autoFilter ref="A' . ($sheet->auto_filter). ':' . $max_cell . '"/>'); // from original source and improvements by okatse (https://github.com/okatse)
		}

        $sheet->file_writer->write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
        $sheet->file_writer->write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
        $sheet->file_writer->write('<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
        $sheet->file_writer->write('<headerFooter differentFirst="false" differentOddEven="false">');
        $sheet->file_writer->write('<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
        $sheet->file_writer->write('<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
        $sheet->file_writer->write('</headerFooter>');
        if(count($this->getIgnoredErrorsCells()) > 0){
            $sheet->file_writer->write('<ignoredErrors>');
            foreach($this->getIgnoredErrorsCells() as $ignoredErrorsCell) {
                $sheet->file_writer->write('<ignoredError sqref="' . $ignoredErrorsCell . '" numberStoredAsText="1"/>');
            }
            $sheet->file_writer->write('</ignoredErrors>');
        }
        if (isset($this->images[$sheet_name]) && \count($this->images[$sheet_name]) > 0) {
            $sheet->file_writer->write('<drawing r:id="rId1" />');
        }
        $sheet->file_writer->write('</worksheet>');

        $max_cell_tag = '<dimension ref="A1:' . $max_cell . '"/>';
        $padding_length = $sheet->max_cell_tag_end - $sheet->max_cell_tag_start - strlen($max_cell_tag);
        $sheet->file_writer->fseek($sheet->max_cell_tag_start);
        $sheet->file_writer->write($max_cell_tag . str_repeat(" ", $padding_length));
        $sheet->file_writer->close();
        $sheet->finalized = true;
    }

    /**
	 * @param string $sheet_name - Name of the sheet
     * @return string
     */
    public function buildDrawingRelationshipXML($sheet_name)
    {
        $drawingXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>' . "\n";
        $drawingXML .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

		$i=1;
        foreach ($this->images[$sheet_name] as $imageId => $image) {
            $drawingXML .= '<Relationship Id="rId' . ($i++) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' . $image['uniq_name'] . '"/>';
        }

        $drawingXML .= "\n" . '</Relationships>';

        return $drawingXML;
    }

    /**
     * @param int $sheetId - The sheet id
     * @return string
     */
    public function buildSheetRelationshipXML($sheetId)
    {
        $lastRelationshipId = 0;

        $rels_xml = "";
        $rels_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>' . "\n";
        $rels_xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$rels_xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' . $sheetId . '.xml"/>';
        $rels_xml .= "\n";
        $rels_xml .= '</Relationships>';

        return $rels_xml;
    }
}