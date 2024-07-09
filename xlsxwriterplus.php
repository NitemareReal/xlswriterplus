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
    private $extensions = [];

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

        $this->extensions[\strtolower(\pathinfo($imagePath, PATHINFO_EXTENSION))] = 1;  // save extension in a (simulated) set to add mime in Content Types XML
		$this->current_sheet = $sheet_name;
    }
    /**
     * @override
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

        $zip->addEmptyDir("xl/theme/");
        $zip->addFromString("xl/theme/theme1.xml", $this->buildThemeXML() );

		if (\count($this->images) > 0) {	            // if there is any image, on any sheet
			$img = \imagecreatetruecolor(100, 100);     // a dummy image, just to get system DPIs
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
     * @override
     * Saves a sheet (BUT it doesn't finalize it. It will be finalized just before creating document!)
     *
     * @param array $data
     * @param string $sheet_name
     * @param array $header_types
     * @return void
     */
    public function writeSheet(array $data, $sheet_name='', array $header_types=array())
	{
		$sheet_name = empty($sheet_name) ? 'Sheet1' : $sheet_name;
		$data = empty($data) ? array(array('')) : $data;
		if (!empty($header_types))
		{
			$this->writeSheetHeader($sheet_name, $header_types);
		}
		foreach($data as $i=>$row)
		{
			$this->writeSheetRow($sheet_name, $row);
		}
	}
    /**
     * @override
	 * @param string $sheet_name - Name of the sheet
     * @return string
     */
    public function buildDrawingXML($sheet_name)
    {
        $imageRelationshipXML = '';

        if(\count($this->images[$sheet_name]) > 0) {

            $imageRelationshipXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n" . 
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
				<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' . $i . '">
					<a:extLst>
						<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
							<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
						</a:ext>
					</a:extLst>
				</a:blip>
				<a:stretch>
					<a:fillRect/>
				</a:stretch>
			</xdr:blipFill>
			<xdr:spPr>
                <a:xfrm>
					<a:ext cx="' . $EMUx . '" cy="' . $EMUy .'"/>
				</a:xfrm>
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
     * @override
     * @return string
     */
    protected function buildContentTypesXML()
    {
        $content_types_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $content_types_xml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        if(\array_key_exists("jpg", $this->extensions) || \array_key_exists("jpeg", $this->extensions)){
            $content_types_xml .= '<Default Extension="jpeg" ContentType="image/jpeg"/>';
        }
        if(\array_key_exists("png", $this->extensions)){
            $content_types_xml .= '<Default Extension="png" ContentType="image/png"/>';
        }
        if(\array_key_exists("gif", $this->extensions)){
            $content_types_xml .= '<Default Extension="png" ContentType="image/gif"/>';
        }
        if(\array_key_exists("tiff", $this->extensions) || \array_key_exists("tif", $this->extensions)){
            $content_types_xml .= '<Default Extension="png" ContentType="image/tiff"/>';
        }
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
		$content_types_xml .= '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
        $content_types_xml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
		$content_types_xml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
//		$content_types_xml .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
		$content_types_xml .= $content_types_drawing_xml;
// 		$content_types_xml .= "\n";
		$content_types_xml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
		$content_types_xml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
		$content_types_xml .= '</Types>';

        return $content_types_xml;
    }

    /**
     * @override
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

        if($max_cell > 0){
        	$max_cell_tag = '<dimension ref="A1:' . $max_cell . '"/>';
		} else {
			$max_cell_tag = '<dimension ref="A1"/>';
		}
        $padding_length = $sheet->max_cell_tag_end - $sheet->max_cell_tag_start - strlen($max_cell_tag);
        $sheet->file_writer->fseek($sheet->max_cell_tag_start);
        $sheet->file_writer->write($max_cell_tag . str_repeat(" ", $padding_length));
        $sheet->file_writer->close();
        $sheet->finalized = true;
    }

    /**
     * @override
	 * @param string $sheet_name - Name of the sheet
     * @return string
     */
    public function buildDrawingRelationshipXML($sheet_name)
    {
        $drawingXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $drawingXML .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

		$i=1;
        foreach ($this->images[$sheet_name] as $imageId => $image) {
            $drawingXML .= '<Relationship Id="rId' . ($i++) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' . $image['uniq_name'] . '"/>';
        }

        $drawingXML .= "\n" . '</Relationships>';

        return $drawingXML;
    }

    /**
     * @override
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
    public function buildThemeXML(){
		$themeXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>';
		return $themeXML;
	}

}