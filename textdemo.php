<!DOCTYPE html>
<html lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />

</head>

<body>
<?php
require_once('wordtext.php');
$rt = new WordTEXT(false,'UTF-8');
$text = $rt->readDocument('sample.docx');

$det = explode(':',$text[0]);
echo "No of text elements in the array - ".$det[0]."<br>";
echo "Max length of a text element in the array - ".$det[1]."<br>&nbsp;<br>";
$LC = 1;
while ($LC <= $det[0]){
	echo "Element ".$LC." : ".$text[$LC]."<br>";
	$LC++;
}

?>
</body>
