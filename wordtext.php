<?php
class WordTEXT
{
	private $debug = false;
	private $file;
	private $styles_xml;
	private $numb_xml;
	private $rels_xml;
	private $doc_xml;
	private $doc_media = [];
	private $last = 'none';
	private $encoding = 'ISO-8859-1';
	
	/**
	 * CONSTRUCTOR
	 * 
	 * @param Boolean $debug Debug mode or not
	 * @param String $encoding selects alternative encoding if required
	 * @return void
	 */
	public function __construct($debug_=null, $encoding=null)
	{
		if($debug_ != null) {
			$this->debug = $debug_;
		}
		if ($encoding != null) {
			$this->encoding = $encoding;
		}
	}


	/**
	 * READS The Document and Relationships into separated XML files
	 * 
	 * @param var $object The class variable to set as DOMDocument 
	 * @param var $xml The xml file
	 * @param string $encoding The encoding to be used
	 * @return void
	 */
	private function setXmlParts(&$object, $xml, $encoding)
	{
		$object = new DOMDocument();
		$object->encoding = $encoding;
		$object->preserveWhiteSpace = false;
		$object->formatOutput = true;
		$object->loadXML($xml);
		$object->saveXML();
	}



	/**
	 * READS The Document and Relationships into separated XML files
	 * 
	 * @param String $filename The filename
	 * @return void
	 */
	private function readZipPart($filename)
	{
		$zip = new ZipArchive();
		$_xml = 'word/document.xml';
		$_xml_rels = 'word/_rels/document.xml.rels';
		$_xml_numb = 'word/numbering.xml';
		$_xml_styles = 'word/styles.xml';
		$_xml_foot = 'word/footnotes.xml';
		$_xml_end = 'word/endnotes.xml';
		
		if (true === $zip->open($filename)) {
			//Get the main word document file
			if (($index = $zip->locateName($_xml)) !== false) {
				$xml = $zip->getFromIndex($index);
			}
			//Get the relationships
			if (($index = $zip->locateName($_xml_rels)) !== false) {
				$xml_rels = $zip->getFromIndex($index);
			}
			//Get the list references from the word numbering file
			if (($index = $zip->locateName($_xml_numb)) !== false) {
				$xml_numb = $zip->getFromIndex($index);
			}
			//Get the style references from the word styles file
			if (($index = $zip->locateName($_xml_styles)) !== false) {
				$xml_styles = $zip->getFromIndex($index);
			}
			//Get the footnotes from the word fonts file
			if (($index = $zip->locateName($_xml_foot)) !== false) {
				$xml_foot = $zip->getFromIndex($index);
			}
			//Get the endnotes from the word fonts file
			if (($index = $zip->locateName($_xml_end)) !== false) {
				$xml_end = $zip->getFromIndex($index);
			}
			$zip->close();
		} else die('non zip file');

		$enc = mb_detect_encoding($xml);
		$this->setXmlParts($this->doc_xml, $xml, $enc);
		$this->setXmlParts($this->rels_xml, $xml_rels, $enc);
		$this->setXmlParts($this->numb_xml, $xml_numb, $enc);
		$this->setXmlParts($this->styles_xml, $xml_styles, $enc);
		$this->setXmlParts($this->foot_xml, $xml_foot, $enc);
		$this->setXmlParts($this->end_xml, $xml_end, $enc);
		
		if($this->debug) {
			echo "XML File : word/document.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->doc_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/_rels/document.xml.rels<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->rels_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/numbering.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->numb_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/styles.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->styles_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/footnotes.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->foot_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/endnotes.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->end_xml->saveXML();
			echo "</textarea>";
		}
	}



	/**
	 * Looks up the footnotes XML file and returns the footnotes if any exist
	 * 
	 * @returns Array - All the footnotes 
	 */
	private function footnotes()
	{
		$reader1 = new XMLReader();
		$reader1->XML($this->foot_xml->saveXML());
		$Ftext = array();
		$hyper = '';
		while ($reader1->read()) {
		// look for required style
			$znum = 1;
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:footnote') { //Get footnote
				$Footnum = $reader1->getAttribute("w:id");
				$st2 = new XMLReader;
				$st2->xml(trim($reader1->readOuterXML()));
				while ($st2->read()) {
					if ($st2->name == 'w:t') {
						$Ftext[$Footnum] .= htmlentities($st2->expand()->textContent);
					}
				}
					
			}
		}
		return $Ftext;
	}


	/**
	 * Looks up the endnotes XML file and returns the endnotes if any exist
	 * 
	 * @returns Array - All the endnotes
	 */
	private function endnotes()
	{
		$reader1 = new XMLReader();
		$reader1->XML($this->end_xml->saveXML());
		$Etext = array();
		while ($reader1->read()) {
		// look for required style
			$znum = 1;
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:endnote') { //Get endnote
				$Endnum = $reader1->getAttribute("w:id");
				$st2 = new XMLReader;
				$st2->xml(trim($reader1->readOuterXML()));
				while ($st2->read()) {
					if ($st2->name == 'w:t') {
						$Etext[$Endnum] .= htmlentities($st2->expand()->textContent);
					}
				}
					
			}
		}
		return $Etext;
	}


	/**
	 * Looks up a style in the styles XML file and returns the paragraph numbering ref if it exists
	 * 
	 * @param String $style - The name of the style
	 * @returns String - The Paragraph numbering reference
	 */
	private function findstyles($style)
	{
		$reader1 = new XMLReader();
		$reader1->XML($this->styles_xml->saveXML());
		$FontTheme = '';
		$parnum = '';
		while ($reader1->read()) {
		// look for required style
			$znum = 1;
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:style') { //Get style settings
				if ($reader1->getAttribute("w:styleId") == $style){
					$st1 = new XMLReader;
					$st1->xml(trim($reader1->readOuterXML()));
					while ($st1->read()) {

						if($st1->name == "w:numId") { // Get paragraph numbering ref
							$parnum = $st1->getAttribute("w:val");
						}
					}
				}
			}
		}
		return $parnum;
	}
	
	

	/**
	 * CONVERTS A NUMBER TO ITS ROMAN PRESENTATION
	 * @param String/Integer $num - The number to be converted
	 * @return String - Roman number
	**/ 
	function numberToRoman($num)  
	{ 
		// Be sure to convert the given parameter into an integer
		$n = intval($num);
		$result = ''; 
 
		// Declare a lookup array that we will use to traverse the number: 
		$lookup = array(
			'm' => 1000, 'cm' => 900, 'd' => 500, 'cd' => 400, 
			'c' => 100, 'xc' => 90, 'l' => 50, 'xl' => 40, 
			'x' => 10, 'ix' => 9, 'v' => 5, 'iv' => 4, 'i' => 1
		); 
 
		foreach ($lookup as $roman => $value)  
		{
			// Look for number of matches
			$matches = intval($n / $value); 
 
			// Concatenate characters
			$result .= str_repeat($roman, $matches); 
 
			// Substract that from the number 
			$n = $n % $value; 
		} 

		return $result; 
	} 

	/**
	 * GETS THE TEXT AND FOOTNOTE AND ENDNOTE REFERENCES FOR A GIVEN ELEMENT
	 * 
	 * @param XML $xml - The XML node
	 * @return String - The element's text including footnote and endnote references
	 */
	private function checkFormating(&$xml)
	{	
		$node = trim($xml->readOuterXML());
		$text = '';
		$reader = new XMLReader();
		$reader->XML($node);
		$FEref = '';
		$Ttmp = '';

		while ($reader->read()) {
			if ($reader->name === 'w:tab') {
				$Ttmp .= " ";
			}

			if($reader->name == "w:footnoteReference") {
				$Ftmp = $reader->getAttribute("w:id");
				$FEref = "[".$Ftmp."]";
			}
			if($reader->name == "w:endnoteReference") {
				$Ftmp = $reader->getAttribute("w:id");
				$FEref = "[".$this->numberToRoman($Ftmp)."]";
			}
			if($reader->name == "w:t") {
				$Ttmp .= htmlentities($reader->expand()->textContent);
			}
		}
		
		if ($FEref <> ''){
			$text = $FEref;
		} else {
			$text =  $Ttmp;
		}
		return $text;
	}
	

	
	/**
	 * CHECKS THE ELEMENT FOR List ELEMENTS and their numbering
	 * 
	 * @param XML $xml - The XML node
	 * @return Array - The list/paragraph numbering
	 */
	private function getListFormating(&$xml)
	{
		static $Listcount = array();
		$node = trim($xml->readOuterXML());

		$reader = new XMLReader();
		$reader->XML($node);
		$PSret= array();
		$LnumA = array();
		$ListnumId = '';
		
		while ($reader->read()){
			if($reader->name == "w:pStyle" && $reader->hasAttributes ) {
				$Pstyle = $reader->getAttribute("w:val");
				$parnum = $this->findstyles($Pstyle); // get the defined styles for this paragraph
				if (substr($Pstyle,0,7) == 'numpara'){
					$ListnumId = $parnum;
					$Listlevel = 0;
				}
			}
			if($reader->name == "w:ilvl" && $reader->hasAttributes) { // List formating - list level
				$Listlevel = $reader->getAttribute("w:val");
			}
			if($reader->name == "w:numId" && $reader->hasAttributes) { // List formating - List cross reference
				$ListnumId = $reader->getAttribute("w:val");
			}

		}
		
		if ($ListnumId){
			// look for the List reference number of this element
			$reader1 = new XMLReader();
			$reader1->XML($this->numb_xml->saveXML());
			while ($reader1->read()) {
				if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:num' && $reader1->getAttribute("w:numId") == $ListnumId) {
					$st1 = new XMLReader;
					$st1->xml(trim($reader1->readOuterXML()));
					while ($st1->read()) {
						if($st1->name == 'w:abstractNumId') {
							$ListAbsNo = $st1->getAttribute("w:val");
						}
					}
				}
			}
			// look for the List details of this element
			$reader2 = new XMLReader();
			$reader2->XML($this->numb_xml->saveXML());
			while ($reader2->read()) {
				if ($reader2->nodeType == XMLREADER::ELEMENT && $reader2->name == 'w:abstractNum' && $reader2->getAttribute("w:abstractNumId") == $ListAbsNo) {
					$st2 = new XMLReader;
					$st2->xml(trim($reader2->readOuterXML()));
					while ($st2->read()) {
						if($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:lvl') {
							$Rlvl = $st2->getAttribute("w:ilvl");
						}
						if($st2->name == 'w:start') {
						$Rstart[$Rlvl] = $st2->getAttribute("w:val");
						}
						if($st2->name == 'w:numFmt') {
							$Rnumfmt[$Rlvl] = $st2->getAttribute("w:val");
						}
						if($st2->name == 'w:lvlText') {
							$Rlvltxt[$Rlvl] = $st2->getAttribute("w:val");
						}
					}
				}
			}
		}
		
		$alphabet = array( 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y'. 'z');
		
  
		if ($ListnumId){ // If the element is a list element get its number
			if (substr($Rlvltxt[$Listlevel],0,1) <> '%'){ // Check if there is a character before the list number and if so get it
				$LNfirst = substr($Rlvltxt[$Listlevel],0,1);
			} else {
				$LNfirst = '';
			}
			$LNlast = substr($Rlvltxt[$Listlevel],-1); // The last character of a list number
			if ($Listcount[$ListnumId][$Listlevel] == ''){ // Get the list number of the list element
				$Listcount[$ListnumId][$Listlevel] = $Rstart[$Listlevel];
				$Listcount[$ListnumId][$Listlevel + 1] = '';
			} else {
				$Listcount[$ListnumId][$Listlevel] = $Listcount[$ListnumId][$Listlevel] + 1;
				$Listcount[$ListnumId][$Listlevel + 1] = '';
			}
			if (strlen($Rlvltxt[$Listlevel]) > 4){
				$Lcount = 0;
			} else {
				$Lcount = $Listlevel;
			}
			while ($Lcount <= $Listlevel){ // produce the list element number
				$LnumA[$Lcount] = $Listcount[$ListnumId][$Lcount]; // The number of the list element
				if ($Rnumfmt[$Lcount] == 'lowerLetter'){
					$LnumA[$Lcount] = $LNfirst.$alphabet[$LnumA[$Lcount]-1].$LNlast;
				} else if ($Rnumfmt[$Lcount] == 'upperLetter'){
					$LnumA[$Lcount] = $LNfirst.strtoupper($alphabet[$LnumA[$Lcount]-1].$LNlast);
				} else if ($Rnumfmt[$Lcount] == 'lowerRoman'){
					$LnumA[$Lcount] = $LNfirst.$this->numberToRoman($LnumA[$Lcount]).$LNlast;
				} else if ($Rnumfmt[$Lcount] == 'upperRoman'){
					$LnumA[$Lcount] = $LNfirst.strtoupper($this->numberToRoman($LnumA[$Lcount])).$LNlast;
				} else if ($Rnumfmt[$Lcount] == 'bullet'){
					$LnumA[$Lcount] = "•";
				} else {
					$LnumA[$Lcount] = $LNfirst.$LnumA[$Lcount].$LNlast;
				}
				$Lnum .= $LnumA[$Lcount];
				$Lcount++;
			}
			$Lnum = $Lnum." ";
		
		}
		
		$PSret['Lnum'] = $Lnum;  // return the element's list number
		$PSret['listnum'] = $ListnumId;
		return $PSret;

	}
	

	/**
	 * CHECKS IF ELEMENT IS AN HYPERLINK
	 *  
	 * @param XML $xml - The XML node
	 * @return String - The Hyperlink text
	 */
	private function getHyperlink(&$xml)
	{
		$ret = array('open'=>'<ul>','close'=>'</ul>');
		$link ='';
		if($xml->hasAttributes) {
			$attribute = "";
			while($xml->moveToNextAttribute()) {
				if($xml->name == "r:id"){  // check for external hyperlinks
					$attribute = $xml->value;
				}
			}
			
			if($attribute != "") {
				$reader = new XMLReader();
				$reader->XML($this->rels_xml->saveXML());
				
				while ($reader->read()) {
					if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name=='Relationship') {
						if($reader->getAttribute("Id") == $attribute) {
							$link = $reader->getAttribute('Target');
							break;
						}
					}
				}
			}
		}
		
		$ret = $link;
		
		return $ret;
	}




	/**
	 * PROCESS PARAGRAPH CONTENT
	 *  
	 * @param XML $xml - The XML node
	 * @return String - The text of the paragraph
	 */
	private function getParagraph(&$paragraph)
	{
		$text = '';
		$list_format=array();
		$Pformat = 'N';
		// loop through paragraph dom
		while ($paragraph->read()) {
			// look for elements
			if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:r') {
				if ($Pformat == 'Y'){
					if($list_format['listnum']){
						$text .= $list_format['Lnum'];
					}
					$Pformat = 'N';
				}
				$text .= $this->checkFormating($paragraph); // Get text				
			} else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:pPr') { // Get list and paragraph formatting
				$list_format = $this->getListFormating($paragraph);
				$Pformat = 'Y';
			} 
			else if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:hyperlink') {
				$hyperlink = $this->getHyperlink($paragraph); // Add in hyperlink ref
				if (substr($hyperlink,0,6) == "mailto"){
					$hyperlink = substr($hyperlink,7);
				}
				$text .= $hyperlink;
				$paragraph->next();
			}
		}
		return $text;
	}
			


	/**
	 * READS THE GIVEN DOCX FILE INTO HTML FORMAT
	 *  
	 * @param String $filename - The DOCX file name
	 * @return Array - The text of the DOCX file
	 */
	public function readDocument($filename)
	{
		
		$this->file = $filename;
		$this->readZipPart($filename);
		$reader = new XMLReader();
		$reader->XML($this->doc_xml->saveXML());
		$text = array();
		$para = 1;
		$maxlen = 0;
		while ($reader->read()) {
		// look for new paragraphs
			$paragraph = new XMLReader;
			$p = $reader->readOuterXML();
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:p') {
				// finds and gets paragraphs			
				$paragraph->xml($p); // set up new instance of XMLReader for parsing paragraph independantly	
				$Rtext = $this->getParagraph($paragraph);
				if ($Rtext <> '' AND $Rtext <> ' '){
					$Tlen = strlen($Rtext);
					if ($Tlen > $maxlen){
						$maxlen = $Tlen;
					}
					$text[$para] = $Rtext;
					$para++;
				}
				$reader->next();
			}
			
		}
		$Foot = $this->footnotes(); // Get any Footnotes in the document
		if ($Foot[1]) {
			$text[$para] = "FOOTNOTES";
			$para++;
			$Fcount = 1;
			while ($Foot[$Fcount]){
				$Ftext = "[".$Fcount."] ".$Foot[$Fcount];
				$Tlen = strlen($Ftext);
				if ($Tlen > $maxlen){
					$maxlen = $Tlen;
				}
				$text[$para] = $Ftext;
				++$Fcount;
				$para++;
			}
		}
		
		$Endn = $this->endnotes(); //Get any Endnotes in the document
		if ($Endn[1]) {
			$text[$para] = "ENDNOTES";
			$para++;
			$Fcount = 1;
			while ($Endn[$Fcount]){
				$Etext = "[".$this->numberToRoman($Fcount)."] ".$Endn[$Fcount];
				$Tlen = strlen($Etext);
				if ($Tlen > $maxlen){
					$maxlen = $Tlen;
				}
				$text[$para] = $Etext;
				++$Fcount;
				$para++;
			}
		}
		$Rcount = $para-1;
		$text[0] = $Rcount.":".$maxlen;
		$reader->close();
		if($this->debug) {  // if in DEBUG mode, display the text of the DOCX document along with the paragraph element numbers of the text array
			echo "<div style='width:100%;'>";
			$det = explode(':',$text[0]);
			echo "No of text elements in the array - ".$det[0]."<br>";
			echo "Max length of a text element in the array - ".$det[1]."<br>&nbsp;<br>";
			$LC = 1;
			while ($LC < $para){
				echo "Element ".$LC." : ";
				echo mb_convert_encoding($text[$LC], $this->encoding);
				echo "<br>";
				$LC++;
			}
			echo "</div>";
		}
		return mb_convert_encoding($text, $this->encoding); // Output the text of the DOCX document as an array
	}
}






