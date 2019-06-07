<?php
function getRow($index,$worksheet) {
  $lastColumn= $worksheet->getHighestColumn();
  return $worksheet->rangeToArray('A' .$index .':' . $lastColumn . $index,NULL,TRUE,FALSE);
}

function getHeadings($worksheet){
  $headings = array();
  for ($i=0; $i < count(getRow(1,$worksheet)[0]); $i++) {
   if(getRow(1,$worksheet)[0][$i]){
     array_push($headings,getRow(1,$worksheet)[0][$i]. '');
   }
  }
  return $headings;
}
function getColumn($index,$worksheet) {
  $lastRow = $worksheet->getHighestRow();
  $columString = PHPExcel_Cell::stringFromColumnIndex($index);
  return $worksheet->rangeToArray($columString .'1:' . $columString . $lastRow,NULL,TRUE,FALSE);
}
function getColumnData($index,$worksheet) {
  $lastRow = $worksheet->getHighestRow();
  $columString = PHPExcel_Cell::stringFromColumnIndex($index);
  return $worksheet->rangeToArray($columString .'2:' . $columString . $lastRow,NULL,TRUE,FALSE);
}
function getColumnByName($name,$worksheet) {
  $lastRow = $worksheet->getHighestRow();
  $headings = getHeadings($worksheet);
  $index = -1;
  foreach ($headings as $key => $value) {
     if(strtoupper($name)==strtoupper($value[0])){
       $index = $key;
       break;
     }
  }
  if($index==-1){
    exit();
  }
  $columString = PHPExcel_Cell::stringFromColumnIndex(null);
  return json_encode($worksheet->rangeToArray($columString .'1:' . $columString . $lastRow,NULL,TRUE,FALSE));
}

function obtainCategoriesPerColumn($index,$worksheet){
  $rowsInColumn = getColumnData($index,$worksheet);
  $categories = array();
  foreach ($rowsInColumn as $key => $value) {
    if (!in_array($value[0], $categories)) {
       array_push($categories,$value[0] .'');
    }
  }
  sort($categories);
  return $categories;
}

function obtainValueAtCoordinate($x,$y,$worksheet){
  $columString = PHPExcel_Cell::stringFromColumnIndex($x);
  $cellValue = $worksheet->getCell($columString.$y)->getValue();
  return $cellValue;
}


function deleteUploads(){
  $files = glob('../../uploads/*tmp');
   foreach($files as $file) {
    unlink($file);
   }
}

function getUploadedFile(){
  $files = glob('../../uploads/*tmp');
  foreach($files as $file) {
   $tmpfname =  $file;
  }
  return $tmpfname;
}
 ?>
