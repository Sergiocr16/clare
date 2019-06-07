<?php
require_once "../Classes/PHPExcel.php";
include 'globalMethods.php';
$tmpfname = getUploadedFile();
$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
$excelObj = $excelReader->load($tmpfname);
$worksheet = $excelObj->getSheet(0);
$lastRow = $worksheet->getHighestRow();
$lastColumn= $worksheet->getHighestColumn();
$lastColumnIndex = PHPExcel_Cell::columnIndexFromString($lastColumn);
$headings = getHeadings($worksheet);
$radioVal = $_POST["type"];
 if (isset($_POST['selected']) && isset($_POST['groupVariable'])){
  $sv = $_POST['selected'];
  $gv= $_POST['groupVariable'];
}else{
header("Location: /clare/php/clare/readExcel.php");
}
$groupVariable = new \stdClass();
$groupVariable->title = $headings[$gv];
$groupVariable->index = $gv;
$groupVariable->categories = obtainCategoriesPerColumn($groupVariable->index,$worksheet);


// Se obtienen las distintas categorias por columna y se crea y da formato a cada una
function getCategoriesPerColumn($column,$worksheet,$categoriesPerGroupVariable){
  $categories = obtainCategoriesPerColumn($column,$worksheet);
  $categoriesPerColum = array();
  // Se recorren las distintas categorias por columna
for ($i=0; $i < count($categories); $i++) {
  array_push($categoriesPerColum,createCategoryRow($categories[$i],$categoriesPerGroupVariable));
}
array_push($categoriesPerColum,createCategoryRow("Total",$categoriesPerGroupVariable));
return $categoriesPerColum;
}

function createCategoryRow($title,$categoriesPerGroupVariable){
  $category = new \stdClass();
  $category->title = $title;
  $category->groups = array();
  $total = new \stdClass();
  $total->number = 0;
  $total->percentage = "";
  $total->acumulated = "";
  $category->total = $total;
  // Se crea un arreglo para almacenar cada cuenta de de la variable categoria
  for ($j=0; $j < count($categoriesPerGroupVariable); $j++) {
    $item = new \stdClass();
    $item->number = 0;
    $item->percentage = "";
    $item->acumulated = "";
    $item->confidenceInterval = "";
    array_push($category->groups,$item);
  }
  return $category;
}


$formatedSelected = array();
for ($i=0; $i < count($sv); $i++) {
  array_push($formatedSelected,getCategoriesPerColumn($sv[$i],$worksheet,$groupVariable->categories));
}


function countVariablesPerGroup($selected,$formatedSelected,$worksheet,$groupVariable){
// echo json_encode($formatedSelected[0]);
  for ($i=0; $i < count($selected); $i++) {
    $columnData = new \stdClass();
    $columnData = getColumnData($selected[$i],$worksheet);
    for ($j=0; $j < count($columnData); $j++) {
        compareTitlePerColumn($j,$columnData[$j][0],$formatedSelected,$groupVariable,$worksheet);
    }
  }
}
function calculatePercentage($total,$num){
  return $num/$total*100;
}

function calculateConfidenceInterval($total,$num){
 $x = $num;
 $n = $total;
 $confidence = 95;
 $desviation = 13.7;
 $zetaAlphaMedio = 1.96;
 $alphaMedio = ((1 - ($confidence/100))/2);
 $mediaMinor = $x - ($zetaAlphaMedio*($desviation/sqrt($n)));
 $mediaMayor = $x + ($zetaAlphaMedio*($desviation/sqrt($n)));
 return "[".number_format((float)$mediaMinor, 1, '.', '').",".number_format((float)$mediaMayor, 1, '.', '')."]";
}
function calculateStandarVariation(){

$nums = array(58.8,41.2);
$sum=0;
for($i=0;$i<count($nums);$i++){
	$sum+=$nums[$i];
}
$media = $sum/count($nums);
$sum2=0;
for($i=0;$i<count($nums);$i++){
	$sum2+=($nums[$i]-$media)*($nums[$i]-$media);
}
$vari = $sum2/count($nums);
$sq = sqrt($vari);
// echo "La varianza es: $vari <br>";
// echo "La desviacion estandar es: ".$sq;
}

function calculatePercentageAndAcum($formatedValue){
for ($i=0; $i < count($formatedValue->groups); $i++) {
  $item = $formatedValue->groups[$i];
  $item->percentage =calculatePercentage($formatedValue->total->number,$item->number);
  if($i==0){
    $item->acumulated = $item->percentage;
  }else{
     $item->acumulated = $item->percentage + $formatedValue->groups[$i-1]->acumulated;
  }
  $item->percentageFormatted = number_format((float)$item->percentage, 1, '.', '').'%';
  $item->acumulatedFormatted = number_format((float)$item->acumulated, 1, '.', '').'%';
  $item->confidenceInterval = calculateConfidenceInterval(100,$item->percentage);
}
}

function calculatePercentageAndAcumTotal($formatedSelected,$itemTotal){
  for ($i=0; $i < count($formatedSelected)-1; $i++) {
    $formatedSelected[$i]->total->percentageFormatted = number_format((float)100, 1, '.', '').'%';
    $formatedSelected[$i]->total->acumulatedFormatted = number_format((float)100, 1, '.', '').'%';
  }
  for ($i=0; $i < count($itemTotal->groups); $i++) {
    $itemTotal->groups[$i]->percentage =  calculatePercentage($itemTotal->total->number,$itemTotal->groups[$i]->number);
    if($i==0){
      $itemTotal->groups[$i]->acumulated = $itemTotal->groups[$i]->percentage;
    }else{
        $itemTotal->groups[$i]->acumulated = $itemTotal->groups[$i]->percentage + $itemTotal->groups[$i-1]->acumulated;
    }
    $itemTotal->groups[$i]->percentageFormatted = number_format((float)$itemTotal->groups[$i]->percentage, 1, '.', '').'%';
    $itemTotal->groups[$i]->acumulatedFormatted = number_format((float)$itemTotal->groups[$i]->acumulated, 1, '.', '').'%';
    $itemTotal->groups[$i]->confidenceInterval = calculateConfidenceInterval(100,$itemTotal->groups[$i]->percentage);
  }
}



function calculateTotalPerSelected($formatedSelected,$itemTotal){
  for ($i=0; $i < count($formatedSelected)-1; $i++) {
   for ($x=0; $x < count($formatedSelected[$i]->groups); $x++) {
        $total = $itemTotal->groups[$x]->number+$formatedSelected[$i]->groups[$x]->number;
     $itemTotal->groups[$x]->number = $total;
   }
   $itemTotal->total->number = $itemTotal->total->number +$formatedSelected[$i]->total->number;
  }
  calculatePercentageAndAcumTotal($formatedSelected,$itemTotal);
}



function compareTitlePerColumn($indexY,$data,$formatedSelected,$groupVariable,$worksheet){
  for ($i=0; $i < count($formatedSelected); $i++) {
    $formatedValue = new \stdClass();
    for ($x=0; $x < count($formatedSelected[$i])-1; $x++) {
      $formatedValue = $formatedSelected[$i][$x];
      if(strtoupper($data)==strtoupper($formatedValue->title)){
        for ($j=0; $j < count($groupVariable->categories); $j++) {
          $valueInGroupColumn = new \stdClass();
          $valueInGroupColumn = obtainValueAtCoordinate($groupVariable->index,$indexY+2,$worksheet). "";
          if(strtoupper($valueInGroupColumn)==strtoupper($groupVariable->categories[$j])){
             $formatedValue->groups[$j]->number = $formatedValue->groups[$j]->number + 1;
             $formatedValue->total->number = $formatedValue->total->number+1;
           }
        }
        calculatePercentageAndAcum($formatedValue);
      }
    }
  }
}



function printGroupTableWord($formatedSelected, $groupVariable, $worksheet, $selected,$radioVal){
  // $headings = getHeadings($worksheet);
  echo '<div id="exportContent" style="width:100%">';
  echo "<table style='width:100%;    width:100.0%;border-collapse:collapse;border:none;mso-border-alt:solid #BFBFBF .5pt;
     mso-border-themecolor:background1;mso-border-themeshade:191;mso-yfti-tbllook:1184;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>";
  echo "<thead >";
  echo "<tr style=''>";
  // echo '<th style="background-color:#26A69A!important;"></th>';
  if($radioVal==1){
    echo '<th style="background-color:#26A69A!important;"" colspan="'.(count($groupVariable->categories)+2).'">';
  }
  if($radioVal==2){
    echo '<th style="background-color:#26A69A!important;"" colspan="'.((count($groupVariable->categories)*3)+3).'">';
  }
  if($radioVal==3){
    echo '<th style="background-color:#26A69A!important;"" colspan="'.((count($groupVariable->categories)*4)+3).'">';
  }
     echo '<b style="color:white!important;">';
     echo $groupVariable->title;
     echo '</b>';
     echo '</th>';
   // echo '<th style="background-color:#26A69A!important; "></th>';
   echo "</thead>";
   echo "<tbody>";
   for ($row = 0; $row < count($selected); $row++) {
   echo "<tr style='background-color:#E0F2F1;'>";
       echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
       mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
       mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
     echo '<b>';
      echo getHeadings($worksheet)[$selected[$row]];
      echo '</b>';
       echo "</td>";
       for ($j = 0; $j < count($groupVariable->categories); $j++) {
         if($radioVal==1){
           echo '<td colspan="1" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
           mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
           mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
           mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
           mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
         }
         if($radioVal==2){
           echo '<td colspan="3" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
           mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
           mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
           mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
           mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
         }
         if($radioVal==3){
           echo '<td colspan="4" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
           mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
           mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
           mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
           mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
         }
         echo '<b>';
         echo $groupVariable->categories[$j];
         echo '</b>';
         echo "</td>";
       }
       if($radioVal==1){
         echo '<td colspan="1" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
         mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
         mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
       }
       if($radioVal==2){
         echo '<td colspan="2" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
         mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
         mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
       }
       if($radioVal==3){
         echo '<td colspan="2" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
         mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
         mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';


       }
       echo '<b>';
       echo "Total general";
       echo '</b>';
       echo "</td>";
   echo "</tr>";
   if($radioVal==2){
   echo "<tr style='background-color:#E0F2F1;'>";
   echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
   mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
   mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
     echo '<b>';
      echo '</b>';
       echo "</td>";
       for ($j = 0; $j < count($groupVariable->categories); $j++) {
         echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
         mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
         mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
           echo 'Num.';
         echo "</td>";
         echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
         mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
         mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
         mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
 echo '%.';
       echo "</td>";
       echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
       mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
       mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
 echo 'Acum.';
     echo "</td>";
       }
       echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
       mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
       mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
       echo 'Num.';
     echo "</td>";
     echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
     mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
     mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
echo '%.';
   echo "</td>";

   echo "</tr>";
 }
 if($radioVal==3){
 echo "<tr style='background-color:#E0F2F1;'>";
 echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
 mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
 mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
 mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
 mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
   echo '<b>';
    echo '</b>';
     echo "</td>";
     for ($j = 0; $j < count($groupVariable->categories); $j++) {
       echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
       mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
       mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
         echo 'Num.';
       echo "</td>";
       echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
       mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
       mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
echo '%.';
     echo "</td>";
     echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
     mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
     mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
echo 'Acum.';
   echo "</td>";
   echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
   mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
   mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
echo 'IC';
 echo "</td>";
     }
     echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
     mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
     mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
     echo 'Num.';
   echo "</td>";
   echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
   mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
   mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
echo '%.';
 echo "</td>";

 echo "</tr>";
}
   if($radioVal==1){
   for ($i=0; $i < count($formatedSelected[$row]); $i++) {
     echo "<tr>";
     echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
     mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
     mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
     echo $formatedSelected[$row][$i]->title;
     echo "</td>";
     for ($x=0; $x < count($formatedSelected[$row][$i]->groups) ; $x++) {
       echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
       mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
       mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
       mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
       echo $formatedSelected[$row][$i]->groups[$x]->number;
       echo "</td>";
     }
     echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
     mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
     mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
     if($i!=count($formatedSelected[$row])-1){
     echo $formatedSelected[$row][$i]->total->number;
   }else{
     echo $formatedSelected[$row][$i]->total->number;

   }
     // echo $formatedSelected[$row][$i]->total->number;
     echo "</td>";
     echo "</tr>";
   }
 }

 if($radioVal==2){
 for ($i=0; $i < count($formatedSelected[$row]); $i++) {
   echo "<tr>";
   echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
   mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
   mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
   echo $formatedSelected[$row][$i]->title;
   echo "</td>";
   for ($x=0; $x < count($formatedSelected[$row][$i]->groups) ; $x++) {
     echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
     mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
     mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
     echo $formatedSelected[$row][$i]->groups[$x]->number;
     echo "</td>";
     echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
     mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
     mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
     echo $formatedSelected[$row][$i]->groups[$x]->percentageFormatted;
     echo "</td>";
     echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
     mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
     mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
     mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
     echo $formatedSelected[$row][$i]->groups[$x]->acumulatedFormatted;
     echo "</td>";
   }
   echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
   mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
   mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
   if($i!=count($formatedSelected[$row])-1){
   echo $formatedSelected[$row][$i]->total->number;
 }else{
   echo $formatedSelected[$row][$i]->total->number;

 }
   // echo $formatedSelected[$row][$i]->total->number;
   echo "</td>";
   echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
   mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
   mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
   mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
   echo "100.0%";
   echo "</td>";
   echo "</tr>";
 }
}

if($radioVal==3){
for ($i=0; $i < count($formatedSelected[$row]); $i++) {
  echo "<tr>";
  echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
  echo $formatedSelected[$row][$i]->title;
      echo "</td>";

  for ($x=0; $x < count($formatedSelected[$row][$i]->groups) ; $x++) {
    echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
    mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
    mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
    echo $formatedSelected[$row][$i]->groups[$x]->number;
    echo "</td>";
    echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
    mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
    mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
    echo $formatedSelected[$row][$i]->groups[$x]->percentageFormatted;
    echo "</td>";
    echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
    mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
    mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
    echo $formatedSelected[$row][$i]->groups[$x]->acumulatedFormatted;
    echo "</td>";
    echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
    mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
    mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
    echo $formatedSelected[$row][$i]->groups[$x]->confidenceInterval;
    echo "</td>";
  }
  echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
  if($i!=count($formatedSelected[$row])-1){
  echo $formatedSelected[$row][$i]->total->number;
}else{
  echo $formatedSelected[$row][$i]->total->number;

}
  // echo $formatedSelected[$row][$i]->total->number;
  echo "</td>";
  echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
  echo "100.0%";
  echo "</td>";
  echo "</tr>";
}
}


 }
  //
  // for ($row = 2; $row <= $lastRow; $row++) {
  //    echo "<tr>";
  //     for ($column = 0; $column <= $lastColumnIndex; $column++) {
  //       $columString = PHPExcel_Cell::stringFromColumnIndex($column);
  //       $value = $worksheet->getCell($columString .$row)->getValue();
  //       if($value!=null){
  //       echo "<td>";
  //       echo $value;
  //       echo "</td>";
  //     }
  //     }
  //      echo "</tr>";
  // }
  echo "</tbody>";
  echo "</table>";
  echo '</div>';
}



function printGroupTable($formatedSelected, $groupVariable, $worksheet, $selected,$radioVal){
  // $headings = getHeadings($worksheet);
  echo '<div id="" style="width:100%">';
  echo "<table style='width:100%; '>";
  echo "<thead >";
  echo "<tr style=''>";
  // echo '<th style="background-color:#26A69A!important;"></th>';
  if($radioVal==1){
    echo '<th style="background-color:#26A69A!important;"" colspan="'.(count($groupVariable->categories)+2).'">';
  }
  if($radioVal==2){
    echo '<th style="background-color:#26A69A!important;"" colspan="'.((count($groupVariable->categories)*3)+3).'">';
  }
  if($radioVal==3){
    echo '<th style="background-color:#26A69A!important;"" colspan="'.((count($groupVariable->categories)*4)+3).'">';
  }
     echo '<b style="color:white!important;">';
     echo $groupVariable->title;
     echo '</b>';
     echo '</th>';
   // echo '<th style="background-color:#26A69A!important; "></th>';
   echo "</thead>";
   echo "<tbody>";
   for ($row = 0; $row < count($selected); $row++) {
   echo "<tr style='background-color:#E0F2F1;'>";
       echo '<td >';
     echo '<b>';
      echo getHeadings($worksheet)[$selected[$row]];
      echo '</b>';
       echo "</td>";
       for ($j = 0; $j < count($groupVariable->categories); $j++) {
         if($radioVal==1){
           echo '<td colspan="1">';
         }
         if($radioVal==2){
           echo '<td colspan="3">';
         }
         if($radioVal==3){
           echo '<td colspan="4">';
         }
         echo '<b>';
         echo $groupVariable->categories[$j];
         echo '</b>';
         echo "</td>";
       }
       if($radioVal==1){
         echo '<td colspan="1">';
       }
       if($radioVal==2){
         echo '<td colspan="2">';
       }
       if($radioVal==3){
         echo '<td colspan="2">';
       }
       echo '<b>';
       echo "Total general";
       echo '</b>';
       echo "</td>";
   echo "</tr>";
   if($radioVal==2){
   echo "<tr style='background-color:#E0F2F1;'>";
       echo '<td>';
     echo '<b>';
      echo '</b>';
       echo "</td>";
       for ($j = 0; $j < count($groupVariable->categories); $j++) {
           echo '<td>';
           echo 'Num.';
         echo "</td>";
         echo '<td>';
 echo '%.';
       echo "</td>";
       echo '<td>';
 echo 'Acum.';
     echo "</td>";
       }
       echo '<td>';
       echo 'Num.';
     echo "</td>";
     echo '<td>';
echo '%.';
   echo "</td>";

   echo "</tr>";
 }
 if($radioVal==3){
 echo "<tr style='background-color:#E0F2F1;'>";
     echo '<td>';
   echo '<b>';
    echo '</b>';
     echo "</td>";
     for ($j = 0; $j < count($groupVariable->categories); $j++) {
         echo '<td>';
         echo 'Num.';
       echo "</td>";
       echo '<td>';
echo '%.';
     echo "</td>";
     echo '<td>';
echo 'Acum.';
   echo "</td>";
   echo '<td>';
echo 'IC';
 echo "</td>";
     }
     echo '<td>';
     echo 'Num.';
   echo "</td>";
   echo '<td>';
echo '%.';
 echo "</td>";

 echo "</tr>";
}
   if($radioVal==1){
   for ($i=0; $i < count($formatedSelected[$row]); $i++) {
     echo "<tr>";
       echo '<td >';
     echo $formatedSelected[$row][$i]->title;
     echo "</td>";
     for ($x=0; $x < count($formatedSelected[$row][$i]->groups) ; $x++) {
       echo "<td>";
       echo $formatedSelected[$row][$i]->groups[$x]->number;
       echo "</td>";
     }
     echo "<td>";
     if($i!=count($formatedSelected[$row])-1){
     echo $formatedSelected[$row][$i]->total->number;
   }else{
     echo $formatedSelected[$row][$i]->total->number;

   }
     // echo $formatedSelected[$row][$i]->total->number;
     echo "</td>";
     echo "</tr>";
   }
 }

 if($radioVal==2){
 for ($i=0; $i < count($formatedSelected[$row]); $i++) {
   echo "<tr>";
     echo '<td >';
   echo $formatedSelected[$row][$i]->title;
   echo "</td>";
   for ($x=0; $x < count($formatedSelected[$row][$i]->groups) ; $x++) {
     echo "<td>";
     echo $formatedSelected[$row][$i]->groups[$x]->number;
     echo "</td>";
     echo "<td>";
     echo $formatedSelected[$row][$i]->groups[$x]->percentageFormatted;
     echo "</td>";
     echo "<td>";
     echo $formatedSelected[$row][$i]->groups[$x]->acumulatedFormatted;
     echo "</td>";
   }
   echo "<td>";
   if($i!=count($formatedSelected[$row])-1){
   echo $formatedSelected[$row][$i]->total->number;
 }else{
   echo $formatedSelected[$row][$i]->total->number;

 }
   // echo $formatedSelected[$row][$i]->total->number;
   echo "</td>";
   echo "<td>";
   echo "100.0%";
   echo "</td>";
   echo "</tr>";
 }
}

if($radioVal==3){
for ($i=0; $i < count($formatedSelected[$row]); $i++) {
  echo "<tr>";
    echo '<td >';
  echo $formatedSelected[$row][$i]->title;
  echo "</td>";
  for ($x=0; $x < count($formatedSelected[$row][$i]->groups) ; $x++) {
    echo "<td>";
    echo $formatedSelected[$row][$i]->groups[$x]->number;
    echo "</td>";
    echo "<td>";
    echo $formatedSelected[$row][$i]->groups[$x]->percentageFormatted;
    echo "</td>";
    echo "<td>";
    echo $formatedSelected[$row][$i]->groups[$x]->acumulatedFormatted;
    echo "</td>";
    echo "<td>";
    echo $formatedSelected[$row][$i]->groups[$x]->confidenceInterval;
    echo "</td>";
  }
  echo "<td>";
  if($i!=count($formatedSelected[$row])-1){
  echo $formatedSelected[$row][$i]->total->number;
}else{
  echo $formatedSelected[$row][$i]->total->number;

}
  // echo $formatedSelected[$row][$i]->total->number;
  echo "</td>";
  echo "<td>";
  echo "100.0%";
  echo "</td>";
  echo "</tr>";
}
}


 }
  //
  // for ($row = 2; $row <= $lastRow; $row++) {
  //    echo "<tr>";
  //     for ($column = 0; $column <= $lastColumnIndex; $column++) {
  //       $columString = PHPExcel_Cell::stringFromColumnIndex($column);
  //       $value = $worksheet->getCell($columString .$row)->getValue();
  //       if($value!=null){
  //       echo "<td>";
  //       echo $value;
  //       echo "</td>";
  //     }
  //     }
  //      echo "</tr>";
  // }
  echo "</tbody>";
  echo "</table>";
  echo '</div>';
}

countVariablesPerGroup($sv,$formatedSelected,$worksheet,$groupVariable);
for ($i=0; $i < count($formatedSelected); $i++) {
  calculateTotalPerSelected($formatedSelected[$i],$formatedSelected[$i][count($formatedSelected[$i])-1]);
}

 // echo json_encode($formatedSelected);
?>

<!DOCTYPE html>
<html lang="es">
<head>
    <title>Bienvenido</title>
    <!--Optimizacion en móbiles-->
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <meta charset="utf-8">
    <meta name="theme-color" content="#0097A7">
    <meta name="MobileOptimized" content="width" >
    <meta name="HandheldFriendly" content="true">
    <meta name="apple-mobile-web-app-capable" content="yes" />
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent" />
    <meta name="format-detection" content="telephone=no">
    <meta name="apple-mobile-web-app-title" content="PWA demo" />
    <!--Links para cs-->
    <link rel="manifest" href="./manifest.json">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <link type="text/css" rel="stylesheet" href="../../css/materialize.min.css"/>
    <link rel="stylesheet" type="text/css" href="../../stylesheet.css">
    <link rel="shortcut icon" href="images/favicon.png" type="image/x-icon">
    <link rel="apple-touch-icon" href="images/favicon.png" sizes="192x192">
</head>

<body class="" >
<!--Barra de navegación-->
   <nav class="white" role="navigation">
    <div class="nav-wrapper container">
      <a id="logo-container" href="index.html" class="brand-logo"><img src="../../res/logo1.jpg" style="margin-top:7px;" alt="Smiley face" height="50" > </a>
      <ul class="right hide-on-med-and-down">
      <span style="color:black;">Dr. Roy Wong</span>
      </ul>
<!--
      <ul id="nav-mobile" class="side-nav">
        <li><a href="php/correo/contacto.php"><i style="font-size: 50px;" class="material-icons">contact_mail</i>Contacto</a></li>
      </ul>
      <a href="#" data-activates="nav-mobile" class="button-collapse"><i class="material-icons">menu</i></a> --> -->
    </div>
   </nav>
<!--Primer Banner-->
<div id="index-banner" class="parallax-container" style="min-height:80px!important;margin-bottom:20px">
  <div class="parallax"><img src="../../res/primera.jpg" alt="Unsplashed background img 1"></div>
  <div class="section no-pad-bot">
    <div class="container" >
      <div class="row center">
        <h5 class="header center white-text text-lighten-2">Agrupamiento</h1>
      </div>
    </div>
  </div>
  </div>
</div>

  <div class="container" style="width:100%!important;">
<?php printGroupTable($formatedSelected,$groupVariable, $worksheet,$sv,$radioVal) ?>
  </div>
  <div class="container " style="width:90%!important;display:none">
    <?php printGroupTableWord($formatedSelected,$groupVariable, $worksheet,$sv,$radioVal) ?>
  </div>

  <a class="btn-floating btn-large waves-effect waves-light  teal lighten-2"  onClick="javascript:history.go(-2)" style="float:left; position:fixed; bottom:10px;left:10px;"><i class="material-icons">replay</i></a>

  <a onclick="Export2Doc('exportContent', 'word-export');" class="btn-floating btn-large waves-effect waves-light deep-orange lighten-1 pulse" style="float:right; position:fixed; bottom:10px;right:10px;"><i class="material-icons">cloud_download</i></a>

    <script type="text/javascript" src="../../js/jquery-3.1.1.min.js"></script>
    <script type="text/javascript" src="../../js/materialize.min.js"></script>
    <script type="text/javascript" src="../../js/init.js"></script>
    <script type="text/javascript">
    function Export2Doc(element, filename = ''){
        var preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
        var postHtml = "</body></html>";
        var html = preHtml+document.getElementById(element).innerHTML+postHtml;

        var blob = new Blob(['\ufeff', html], {
            type: 'application/msword'
        });

        // Specify link url
        var url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(html);

        // Specify file name
        filename = filename?filename+'.doc':'document.doc';

        // Create download link element
        var downloadLink = document.createElement("a");

        document.body.appendChild(downloadLink);

        if(navigator.msSaveOrOpenBlob ){
            navigator.msSaveOrOpenBlob(blob, filename);
        }else{
            // Create a link to the file
            downloadLink.href = url;

            // Setting the file name
            downloadLink.download = filename;

            //triggering the function
            downloadLink.click();
        }

        document.body.removeChild(downloadLink);
    }
    </script>
</body>
</html>
