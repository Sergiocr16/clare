<?php
require_once "../Classes/PHPExcel.php";
include 'globalMethods.php';
$tmpfname = getUploadedFile();
$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
$excelObj = $excelReader->load($tmpfname);
$worksheet = $excelObj->getSheet(0);
$lastRow = $worksheet->getHighestRow();
$lastColumn = $worksheet->getHighestColumn();
$lastColumnIndex = PHPExcel_Cell::columnIndexFromString($lastColumn);
$headings = getHeadings($worksheet);
$radioVal = $_POST["type"];
if (isset($_POST['acumulated'])) {
    $acum = true;
} else {
    $acum = false;
}
if (isset($_POST['selected']) && isset($_POST['groupVariable'])) {
    $sv = $_POST['selected'];
    $gv = $_POST['groupVariable'];
    if ($gv != - 1) {
        $groupVariable = new \stdClass();
        $groupVariable->title = $headings[$gv];
        $groupVariable->index = $gv;
        $groupVariable->categories = obtainCategoriesPerColumn($groupVariable->index, $worksheet);
    }
} else {
    header("Location: /clare/php/clare/readExcel.php");
}
// Se obtienen las distintas categorias por columna y se crea y da formato a cada una
function getCategoriesPerColumn($column, $worksheet, $categoriesPerGroupVariable) {
    $categories = obtainCategoriesPerColumn($column, $worksheet);
    $categoriesPerColum = array();
    // Se recorren las distintas categorias por columna
    for ($i = 0;$i < count($categories);$i++) {
        array_push($categoriesPerColum, createCategoryRow($categories[$i], $categoriesPerGroupVariable));
    }
    array_push($categoriesPerColum, createCategoryRow("Total", $categoriesPerGroupVariable));
    return $categoriesPerColum;
}
function getSelectedFormatted($column, $worksheet) {
    $categories = obtainCategoriesPerColumn($column, $worksheet);
    $categoriesPerColum = array();
    // Se recorren las distintas categorias por columna
    for ($i = 0;$i < count($categories);$i++) {
        array_push($categoriesPerColum, createCategoryRowSingle($categories[$i]));
    }
    array_push($categoriesPerColum, createCategoryRowSingle("Total"));
    return $categoriesPerColum;
}
function createCategoryRowSingle($title) {
    $category = new \stdClass();
    $category->title = $title;
    $category->groups = array();
    $total = new \stdClass();
    $total->number = 0;
    $total->percentage = "";
    $total->acumulated = "";
    $category->total = $total;
    // Se crea un arreglo para almacenar cada cuenta de de la variable categoria
    $item = new \stdClass();
    $item->number = 0;
    $item->percentage = "";
    $item->acumulated = "";
    $item->confidenceInterval = "";
    array_push($category->groups, $item);
    return $category;
}
function createCategoryRow($title, $categoriesPerGroupVariable) {
    $category = new \stdClass();
    $category->title = $title;
    $category->groups = array();
    $total = new \stdClass();
    $total->number = 0;
    $total->percentage = "";
    $total->acumulated = "";
    $category->total = $total;
    // Se crea un arreglo para almacenar cada cuenta de de la variable categoria
    for ($j = 0;$j < count($categoriesPerGroupVariable);$j++) {
        $item = new \stdClass();
        $item->number = 0;
        $item->percentage = "";
        $item->acumulated = "";
        $item->confidenceInterval = "";
        array_push($category->groups, $item);
    }
    return $category;
}
$formatedSelected = array();
if ($gv != - 1) {
    for ($i = 0;$i < count($sv);$i++) {
        array_push($formatedSelected, getCategoriesPerColumn($sv[$i], $worksheet, $groupVariable->categories));
    }
} else {
    for ($i = 0;$i < count($sv);$i++) {
        array_push($formatedSelected, getSelectedFormatted($sv[$i], $worksheet));
    }
}
function countVariablesPerGroup($selected, $formatedSelected, $worksheet, $groupVariable) {
    // echo json_encode($formatedSelected[0]);
    for ($i = 0;$i < count($selected);$i++) {
        $columnData = new \stdClass();
        $columnData = getColumnData($selected[$i], $worksheet);
        for ($j = 0;$j < count($columnData);$j++) {
            compareTitlePerColumn($j, $columnData[$j][0], $formatedSelected, $groupVariable, $worksheet);
        }
    }
}
function countVariables($selected, $formatedSelected, $worksheet) {
    for ($i = 0;$i < count($selected);$i++) {
        $columnData = new \stdClass();
        $columnData = getColumnData($selected[$i], $worksheet);
        for ($j = 0;$j < count($columnData);$j++) {
            compareTitlePerColumnWithoutGroup($j, $columnData[$j][0], $formatedSelected, $selected[$i], $worksheet);
        }
    }
}
function calculatePercentage($total, $num) {
    if ($total != 0) {
        return $num / $total * 100;
    } else {
        return 0;
    }
}
function calculateConfidenceInterval($total, $num) {
    $confidence = 95;
    $proportion = $num/$total;
    $percentaje = $proportion*100;
    $ee = sqrt(($proportion)*((1-$proportion)/$total));
    $zetaAlphaMedio = 1.96;
    $correction = 0.5/$total;
    $eeZ = $ee*$zetaAlphaMedio;
    $limitMinor = ($proportion-$eeZ-$correction)*100;
    $limitMayor = ($proportion+$eeZ+$correction)*100;
    // $x = $num;
    // $n = $total;
    // $desviation = 3.17;
    // $zetaAlphaMedio = 1.96;
    // $alphaMedio = ((1 - ($confidence / 100)) / 2);
    // $mediaMinor = $x - ($zetaAlphaMedio * ($desviation / sqrt($n)));
    // $mediaMayor = $x + ($zetaAlphaMedio * ($desviation / sqrt($n)));
    // return "[" . number_format((float)$mediaMinor, 1, '.', '') . "," . number_format((float)$mediaMayor, 1, '.', '') . "]";
    return "[" . number_format((float)$limitMinor, 1, '.', '') . "," . number_format((float)$limitMayor, 1, '.', '') . "]";

}
function calculateStandarVariation() {
    $nums = array(58.8, 41.2);
    $sum = 0;
    for ($i = 0;$i < count($nums);$i++) {
        $sum+= $nums[$i];
    }
    $media = $sum / count($nums);
    $sum2 = 0;
    for ($i = 0;$i < count($nums);$i++) {
        $sum2+= ($nums[$i] - $media) * ($nums[$i] - $media);
    }
    $vari = $sum2 / count($nums);
    $sq = sqrt($vari);
    // echo "La varianza es: $vari <br>";
    // echo "La desviaICon estandar es: ".$sq;

}
function calculatePercentageAndAcum($formatedValue) {
    for ($i = 0;$i < count($formatedValue->groups);$i++) {
        $item = $formatedValue->groups[$i];
        $item->percentage = calculatePercentage($formatedValue->total->number, $item->number);
        if ($i == 0) {
            $item->acumulated = $item->percentage;
        } else {
            $item->acumulated = $item->percentage + $formatedValue->groups[$i - 1]->acumulated;
        }
        $item->percentageFormatted = number_format((float)$item->percentage, 1, '.', '') . '%';
        $item->acumulatedFormatted = number_format((float)$item->acumulated, 1, '.', '') . '%';
        $item->confidenceInterval = calculateConfidenceInterval($formatedValue->total->number, $item->number);
    }
}
function calculatePercentageAndAcumNoGroup($formatedValue, $itemTotal) {
    for ($i = 0;$i < count($formatedValue->groups);$i++) {
        $item = $formatedValue->groups[$i];
        $item->percentage = calculatePercentage($itemTotal->total->number, $formatedValue->total->number);
        if ($i == 0) {
            $item->acumulated = $item->percentage;
        } else {
            $item->acumulated = $item->percentage + $formatedValue->groups[$i - 1]->acumulated;
        }
        $item->percentageFormatted = number_format((float)$item->percentage, 1, '.', '') . '%';
        $item->acumulatedFormatted = number_format((float)$item->acumulated, 1, '.', '') . '%';
        $item->confidenceInterval = calculateConfidenceInterval($itemTotal->total->number, $item->number);
    }
}
function calculatePercentageAndAcumTotal($formatedSelected, $itemTotal) {
    for ($i = 0;$i < count($formatedSelected) - 1;$i++) {
        $formatedSelected[$i]->total->percentageFormatted = number_format((float)100, 1, '.', '') . '%';
        $formatedSelected[$i]->total->acumulatedFormatted = number_format((float)100, 1, '.', '') . '%';
    }
    for ($i = 0;$i < count($itemTotal->groups);$i++) {
        $itemTotal->groups[$i]->percentage = calculatePercentage($itemTotal->total->number, $itemTotal->groups[$i]->number);
        if ($i == 0) {
            $itemTotal->groups[$i]->acumulated = $itemTotal->groups[$i]->percentage;
        } else {
            $itemTotal->groups[$i]->acumulated = $itemTotal->groups[$i]->percentage + $itemTotal->groups[$i - 1]->acumulated;
        }
        $itemTotal->groups[$i]->percentageFormatted = number_format((float)$itemTotal->groups[$i]->percentage, 1, '.', '') . '%';
        $itemTotal->groups[$i]->acumulatedFormatted = number_format((float)$itemTotal->groups[$i]->acumulated, 1, '.', '') . '%';
        $itemTotal->groups[$i]->confidenceInterval = calculateConfidenceInterval(100, $itemTotal->groups[$i]->percentage);
    }
}
function calculatePercentageAndAcumTotalNoGroup($formatedSelected, $itemTotal) {
    for ($i = 0;$i < count($formatedSelected);$i++) {
        $formatedSelected[$i]->total->percentageFormatted = number_format((float)100, 1, '.', '') . '%';
        $formatedSelected[$i]->total->acumulatedFormatted = number_format((float)100, 1, '.', '') . '%';
        calculatePercentageAndAcumNoGroup($formatedSelected[$i], $itemTotal);
    }
    for ($i = 0;$i < count($itemTotal->groups);$i++) {
        $itemTotal->groups[$i]->percentage = calculatePercentage($itemTotal->total->number, $itemTotal->groups[$i]->number);
        if ($i == 0) {
            $itemTotal->groups[$i]->acumulated = $itemTotal->groups[$i]->percentage;
        } else {
            $itemTotal->groups[$i]->acumulated = $itemTotal->groups[$i]->percentage + $itemTotal->groups[$i - 1]->acumulated;
        }
        $itemTotal->groups[$i]->percentageFormatted = number_format((float)$itemTotal->groups[$i]->percentage, 1, '.', '') . '%';
        $itemTotal->groups[$i]->acumulatedFormatted = number_format((float)$itemTotal->groups[$i]->acumulated, 1, '.', '') . '%';
        $itemTotal->groups[$i]->confidenceInterval = calculateConfidenceInterval(100, $itemTotal->groups[$i]->percentage);
    }
}
function calculateTotalPerSelected($formatedSelected, $itemTotal) {
    for ($i = 0;$i < count($formatedSelected) - 1;$i++) {
        for ($x = 0;$x < count($formatedSelected[$i]->groups);$x++) {
            $total = $itemTotal->groups[$x]->number + $formatedSelected[$i]->groups[$x]->number;
            $itemTotal->groups[$x]->number = $total;
        }
        $itemTotal->total->number = $itemTotal->total->number + $formatedSelected[$i]->total->number;
    }
    calculatePercentageAndAcumTotal($formatedSelected, $itemTotal);
}
function calculateTotalPerSelectedWithoutGroup($formatedSelected, $itemTotal) {
    for ($i = 0;$i < count($formatedSelected) - 1;$i++) {
        for ($x = 0;$x < count($formatedSelected[$i]->groups);$x++) {
            $total = $itemTotal->groups[$x]->number + $formatedSelected[$i]->groups[$x]->number;
            $itemTotal->groups[$x]->number = $total;
        }
        $itemTotal->total->number = $itemTotal->total->number + $formatedSelected[$i]->total->number;
    }
    calculatePercentageAndAcumTotalNoGroup($formatedSelected, $itemTotal);
}
function compareTitlePerColumn($indexY, $data, $formatedSelected, $groupVariable, $worksheet) {
    for ($i = 0;$i < count($formatedSelected);$i++) {
        $formatedValue = new \stdClass();
        for ($x = 0;$x < count($formatedSelected[$i]) - 1;$x++) {
            $formatedValue = $formatedSelected[$i][$x];
            if (compareStrings($data, $formatedValue->title)) {
                for ($j = 0;$j < count($groupVariable->categories);$j++) {
                    $valueInGroupColumn = new \stdClass();
                    $valueInGroupColumn = obtainValueAtCoordinate($groupVariable->index, $indexY + 2, $worksheet) . "";
                    if (compareStrings($valueInGroupColumn, $groupVariable->categories[$j])) {
                        $formatedValue->groups[$j]->number = $formatedValue->groups[$j]->number + 1;
                        $formatedValue->total->number = $formatedValue->total->number + 1;
                    }
                }
                calculatePercentageAndAcum($formatedValue);
            }
        }
    }
}
function compareTitlePerColumnWithoutGroup($indexY, $data, $formatedSelected, $indexX, $worksheet) {
    $total = 0;
    for ($i = 0;$i < count($formatedSelected);$i++) {
        $formatedValue = new \stdClass();
        for ($x = 0;$x < count($formatedSelected[$i]);$x++) {
            $formatedValue = $formatedSelected[$i][$x];
            if (compareStrings($data, $formatedValue->title)) {
                $valueInSelectedColumn = new \stdClass();
                $valueInSelectedColumn = obtainValueAtCoordinate($indexX, $indexY + 2, $worksheet) . "";
                $formatedValue->groups[0]->number = $formatedValue->groups[0]->number + 1;
                $formatedValue->total->number = $formatedValue->total->number + 1;
                $total = $total + $formatedValue->total->number;
            }
        }
    }
}

function printNoGroupTableWord($formatedSelected, $worksheet, $selected, $radioVal) {
  echo '<div id="exportContentNoGroup" style="width:100%">';
  echo "<table style='width:100%;width:100.0%;border-collapse:collapse;border:none;mso-border-alt:solid #BFBFBF .5pt;
   mso-border-themecolor:background1;mso-border-themeshade:191;mso-yfti-tbllook:1184;mso-padding-alt:0cm 5.4pt 0cm 5.4pt;font-family:Arial'>";
    echo "<tbody>";
    for ($row = 0;$row < count($selected);$row++) {
        echo "<tr style='background-color:#E0F2F1;'>";
        echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
        echo '<b>';
        echo getHeadings($worksheet) [$selected[$row]];
        echo '</b>';
        echo "</td>";
        echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
        echo '<b>';
        echo "Num";
        echo '</b>';
        echo "</td>";
        if ($radioVal == 2) {
          echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
            echo '<b>';
            echo "%";
            echo '</b>';
            echo "</td>";
        }
        if ($radioVal == 3) {
          echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
            echo '<b>';
            echo "%";
            echo '</b>';
            echo "</td>";
            echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
            echo '<b>';
            echo "IC 95%";
            echo '</b>';
            echo "</td>";
        }
        echo "</tr>";
        if (true) {
            for ($i = 0;$i < count($formatedSelected[$row]);$i++) {
                echo "<tr>";
                echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                echo $formatedSelected[$row][$i]->title;
                echo "</td>";
                for ($x = 0;$x < count($formatedSelected[$row][$i]->groups);$x++) {
                  echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
    mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
    mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                    echo $formatedSelected[$row][$i]->groups[$x]->number;
                    echo "</td>";
                    if ($radioVal == 2) {
                      echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
        mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
        mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
        mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
        mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                        echo $formatedSelected[$row][$i]->groups[$x]->percentageFormatted;
                        echo "</td>";
                    }
                    if ($radioVal == 3) {
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
                        echo $formatedSelected[$row][$i]->groups[$x]->confidenceInterval;
                        echo "</td>";
                    }
                }
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


function printNoGroupTable($formatedSelected, $worksheet, $selected, $radioVal) {
    echo '<div id="" style="width:100%">';
    echo "<table style='width:100%; '>";
    echo "<tbody>";
    for ($row = 0;$row < count($selected);$row++) {
        echo "<tr style='background-color:#E0F2F1;'>";
        echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
        echo '<b>';
        echo getHeadings($worksheet) [$selected[$row]];
        echo '</b>';
        echo "</td>";
        echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
        echo '<b>';
        echo "Num";
        echo '</b>';
        echo "</td>";
        if ($radioVal == 2) {
          echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
            echo '<b>';
            echo "%";
            echo '</b>';
            echo "</td>";
        }
        if ($radioVal == 3) {
          echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
            echo '<b>';
            echo "%";
            echo '</b>';
            echo "</td>";
            echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
            echo '<b>';
            echo "IC 95%";
            echo '</b>';
            echo "</td>";
        }
        echo "</tr>";
        if (true) {
            for ($i = 0;$i < count($formatedSelected[$row]);$i++) {
                echo "<tr>";
                echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                echo $formatedSelected[$row][$i]->title;
                echo "</td>";
                for ($x = 0;$x < count($formatedSelected[$row][$i]->groups);$x++) {
                  echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
    mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
    mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                    echo $formatedSelected[$row][$i]->groups[$x]->number;
                    echo "</td>";
                    if ($radioVal == 2) {
                      echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
        mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
        mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
        mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
        mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                        echo $formatedSelected[$row][$i]->groups[$x]->percentageFormatted;
                        echo "</td>";
                    }
                    if ($radioVal == 3) {
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
                        echo $formatedSelected[$row][$i]->groups[$x]->confidenceInterval;
                        echo "</td>";
                    }
                }
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
function printGroupTableWord($formatedSelected, $groupVariable, $worksheet, $selected, $radioVal, $acum) {
    // $headings = getHeadings($worksheet);
    echo '<div id="exportContent" style="width:100%">';
    echo "<table style='width:100%;    width:100.0%;border-collapse:collapse;border:none;mso-border-alt:solid #BFBFBF .5pt;
     mso-border-themecolor:background1;mso-border-themeshade:191;mso-yfti-tbllook:1184;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>";
    echo "<thead >";
    echo "<tr style=''>";
    // echo '<th style="background-color:#26A69A!important;"></th>';
    if ($radioVal == 1) {
        echo '<th style="background-color:#26A69A!important;"" colspan="' . (count($groupVariable->categories) + 2) . '">';
    }
    if ($radioVal == 2) {
        if ($acum) {
            echo '<th style="background-color:#26A69A!important;"" colspan="' . ((count($groupVariable->categories) * 3) + 3) . '">';
        } else {
            echo '<th style="background-color:#26A69A!important;"" colspan="' . ((count($groupVariable->categories) * 2) + 3) . '">';
        }
    }
    if ($radioVal == 3) {
        echo '<th style="background-color:#26A69A!important;"" colspan="' . ((count($groupVariable->categories) * 4) + 3) . '">';
    }
    echo '<b style="color:white!important;">';
    echo $groupVariable->title;
    echo '</b>';
    echo '</th>';
    // echo '<th style="background-color:#26A69A!important; "></th>';
    echo "</thead>";
    echo "<tbody>";
    for ($row = 0;$row < count($selected);$row++) {
        echo "<tr style='background-color:#E0F2F1;'>";
        echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
        echo '<b>';
        echo getHeadings($worksheet) [$selected[$row]];
        echo '</b>';
        echo "</td>";
        for ($j = 0;$j < count($groupVariable->categories);$j++) {
            if ($radioVal == 1) {
                echo '<td colspan="1" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
            }
            if ($radioVal == 2) {
                if ($acum) {
                    echo '<td colspan="3" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
      mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
      mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
      mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
      mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                } else {
                    echo '<td colspan="2" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
      mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
      mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
      mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
      mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                }
            }
            if ($radioVal == 3) {
                if ($acum) {
                    echo '<td colspan="4" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
      mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
      mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
      mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
      mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                } else {
                    echo '<td colspan="3" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
      mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
      mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
      mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
      mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                }
            }
            echo '<b>';
            echo $groupVariable->categories[$j];
            echo '</b>';
            echo "</td>";
        }
        if ($radioVal == 1) {

            echo '<td colspan="1" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
        }
        if ($radioVal == 2) {

            echo '<td colspan="2" style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
        }
        if ($radioVal == 3) {

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
        if ($radioVal == 2) {
            echo "<tr style='background-color:#E0F2F1;'>";
            echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
            echo '<b>';
            echo '</b>';
            echo "</td>";
            for ($j = 0;$j < count($groupVariable->categories);$j++) {
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
                if ($acum) {
                  echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
    mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
    mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
    mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                    echo 'Acum.';
                    echo "</td>";
                }
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
        if ($radioVal == 3) {
            echo "<tr style='background-color:#E0F2F1;'>";
            echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
            echo '<b>';
            echo '</b>';
            echo "</td>";
            for ($j = 0;$j < count($groupVariable->categories);$j++) {
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
                if ($acum) {
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
                echo 'IC 95%';
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
        if ($radioVal == 1) {
            for ($i = 0;$i < count($formatedSelected[$row]);$i++) {
                echo "<tr>";
                echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                echo $formatedSelected[$row][$i]->title;
                echo "</td>";
                for ($x = 0;$x < count($formatedSelected[$row][$i]->groups);$x++) {
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
                if ($i != count($formatedSelected[$row]) - 1) {
                    echo $formatedSelected[$row][$i]->total->number;
                } else {
                    echo $formatedSelected[$row][$i]->total->number;
                }
                // echo $formatedSelected[$row][$i]->total->number;
                echo "</td>";
                echo "</tr>";
            }
        }
        if ($radioVal == 2) {
            for ($i = 0;$i < count($formatedSelected[$row]);$i++) {
                echo "<tr>";
                echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                echo $formatedSelected[$row][$i]->title;
                echo "</td>";
                for ($x = 0;$x < count($formatedSelected[$row][$i]->groups);$x++) {
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
                    if ($acum) {
                        echo "<td>";
                        echo $formatedSelected[$row][$i]->groups[$x]->acumulatedFormatted;
                        echo "</td>";
                    }
                }
                echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                if ($i != count($formatedSelected[$row]) - 1) {
                    echo $formatedSelected[$row][$i]->total->number;
                } else {
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
        if ($radioVal == 3) {
            for ($i = 0;$i < count($formatedSelected[$row]);$i++) {
                echo "<tr>";
                echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                echo $formatedSelected[$row][$i]->title;
                echo "</td>";
                for ($x = 0;$x < count($formatedSelected[$row][$i]->groups);$x++) {
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
                    if ($acum) {
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
                    echo $formatedSelected[$row][$i]->groups[$x]->confidenceInterval;
                    echo "</td>";
                }
                echo '<td style="border:solid #BFBFBF 1.0pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;border-top:none;mso-border-top-alt:solid #BFBFBF .5pt;
  mso-border-top-themecolor:background1;mso-border-top-themeshade:191;
  mso-border-alt:solid #BFBFBF .5pt;mso-border-themecolor:background1;
  mso-border-themeshade:191;padding:0cm 5.4pt 0cm 5.4pt">';
                if ($i != count($formatedSelected[$row]) - 1) {
                    echo $formatedSelected[$row][$i]->total->number;
                } else {
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
function printGroupTable($formatedSelected, $groupVariable, $worksheet, $selected, $radioVal, $acum) {
    // $headings = getHeadings($worksheet);
    echo '<div id="" style="width:100%">';
    echo "<table style='width:100%; '>";
    echo "<thead >";
    echo "<tr style=''>";
    // echo '<th style="background-color:#26A69A!important;"></th>';
    if ($radioVal == 1) {
        echo '<th style="background-color:#26A69A!important;"" colspan="' . (count($groupVariable->categories) + 2) . '">';
    }
    if ($radioVal == 2) {
        if ($acum) {
            echo '<th style="background-color:#26A69A!important;"" colspan="' . ((count($groupVariable->categories) * 3) + 3) . '">';
        } else {
            echo '<th style="background-color:#26A69A!important;"" colspan="' . ((count($groupVariable->categories) * 2) + 3) . '">';
        }
    }
    if ($radioVal == 3) {
        echo '<th style="background-color:#26A69A!important;"" colspan="' . ((count($groupVariable->categories) * 4) + 3) . '">';
    }
    echo '<b style="color:white!important;">';
    echo $groupVariable->title;
    echo '</b>';
    echo '</th>';
    // echo '<th style="background-color:#26A69A!important; "></th>';
    echo "</thead>";
    echo "<tbody>";
    for ($row = 0;$row < count($selected);$row++) {
        echo "<tr style='background-color:#E0F2F1;'>";
        echo '<td >';
        echo '<b>';
        echo getHeadings($worksheet) [$selected[$row]];
        echo '</b>';
        echo "</td>";
        for ($j = 0;$j < count($groupVariable->categories);$j++) {
            if ($radioVal == 1) {
                echo '<td colspan="1">';
            }
            if ($radioVal == 2) {
                if ($acum) {
                    echo '<td colspan="3">';
                } else {
                    echo '<td colspan="2">';
                }
            }
            if ($radioVal == 3) {
                if ($acum) {
                    echo '<td colspan="4">';
                } else {
                    echo '<td colspan="3">';
                }
            }
            echo '<b>';
            echo $groupVariable->categories[$j];
            echo '</b>';
            echo "</td>";
        }
        if ($radioVal == 1) {
            echo '<td colspan="1">';
        }
        if ($radioVal == 2) {
            echo '<td colspan="2">';
        }
        if ($radioVal == 3) {
            echo '<td colspan="2">';
        }
        echo '<b>';
        echo "Total general";
        echo '</b>';
        echo "</td>";
        echo "</tr>";
        if ($radioVal == 2) {
            echo "<tr style='background-color:#E0F2F1;'>";
            echo '<td>';
            echo '<b>';
            echo '</b>';
            echo "</td>";
            for ($j = 0;$j < count($groupVariable->categories);$j++) {
                echo '<td>';
                echo 'Num.';
                echo "</td>";
                echo '<td>';
                echo '%.';
                echo "</td>";
                if ($acum) {
                    echo '<td>';
                    echo 'Acum.';
                    echo "</td>";
                }
            }
            echo '<td>';
            echo 'Num.';
            echo "</td>";
            echo '<td>';
            echo '%.';
            echo "</td>";
            echo "</tr>";
        }
        if ($radioVal == 3) {
            echo "<tr style='background-color:#E0F2F1;'>";
            echo '<td>';
            echo '<b>';
            echo '</b>';
            echo "</td>";
            for ($j = 0;$j < count($groupVariable->categories);$j++) {
                echo '<td>';
                echo 'Num.';
                echo "</td>";
                echo '<td>';
                echo '%.';
                echo "</td>";
                if ($acum) {
                    echo '<td>';
                    echo 'Acum.';
                    echo "</td>";
                }
                echo '<td>';
                echo 'IC 95%';
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
        if ($radioVal == 1) {
            for ($i = 0;$i < count($formatedSelected[$row]);$i++) {
                echo "<tr>";
                echo '<td >';
                echo $formatedSelected[$row][$i]->title;
                echo "</td>";
                for ($x = 0;$x < count($formatedSelected[$row][$i]->groups);$x++) {
                    echo "<td>";
                    echo $formatedSelected[$row][$i]->groups[$x]->number;
                    echo "</td>";
                }
                echo "<td>";
                if ($i != count($formatedSelected[$row]) - 1) {
                    echo $formatedSelected[$row][$i]->total->number;
                } else {
                    echo $formatedSelected[$row][$i]->total->number;
                }
                // echo $formatedSelected[$row][$i]->total->number;
                echo "</td>";
                echo "</tr>";
            }
        }
        if ($radioVal == 2) {
            for ($i = 0;$i < count($formatedSelected[$row]);$i++) {
                echo "<tr>";
                echo '<td >';
                echo $formatedSelected[$row][$i]->title;
                echo "</td>";
                for ($x = 0;$x < count($formatedSelected[$row][$i]->groups);$x++) {
                    echo "<td>";
                    echo $formatedSelected[$row][$i]->groups[$x]->number;
                    echo "</td>";
                    echo "<td>";
                    echo $formatedSelected[$row][$i]->groups[$x]->percentageFormatted;
                    echo "</td>";
                    if ($acum) {
                        echo "<td>";
                        echo $formatedSelected[$row][$i]->groups[$x]->acumulatedFormatted;
                        echo "</td>";
                    }
                }
                echo "<td>";
                if ($i != count($formatedSelected[$row]) - 1) {
                    echo $formatedSelected[$row][$i]->total->number;
                } else {
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
        if ($radioVal == 3) {
            for ($i = 0;$i < count($formatedSelected[$row]);$i++) {
                echo "<tr>";
                echo '<td >';
                echo $formatedSelected[$row][$i]->title;
                echo "</td>";
                for ($x = 0;$x < count($formatedSelected[$row][$i]->groups);$x++) {
                    echo "<td>";
                    echo $formatedSelected[$row][$i]->groups[$x]->number;
                    echo "</td>";
                    echo "<td>";
                    echo $formatedSelected[$row][$i]->groups[$x]->percentageFormatted;
                    echo "</td>";
                    if ($acum) {
                        echo "<td>";
                        echo $formatedSelected[$row][$i]->groups[$x]->acumulatedFormatted;
                        echo "</td>";
                    }
                    echo "<td>";
                    echo $formatedSelected[$row][$i]->groups[$x]->confidenceInterval;
                    echo "</td>";
                }
                echo "<td>";
                if ($i != count($formatedSelected[$row]) - 1) {
                    echo $formatedSelected[$row][$i]->total->number;
                } else {
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
if ($gv != - 1) {
    countVariablesPerGroup($sv, $formatedSelected, $worksheet, $groupVariable);
    for ($i = 0;$i < count($formatedSelected);$i++) {
        calculateTotalPerSelected($formatedSelected[$i], $formatedSelected[$i][count($formatedSelected[$i]) - 1]);
    }
} else {
    countVariables($sv, $formatedSelected, $worksheet);
    for ($i = 0;$i < count($formatedSelected);$i++) {
        calculateTotalPerSelectedWithoutGroup($formatedSelected[$i], $formatedSelected[$i][count($formatedSelected[$i]) - 1]);
    }
}
// echo json_encode($formatedSelected);

?>

<!DOCTYPE html>
<html lang="es">
<head>
    <title>Bienvenido</title>
    <!--OptimizaICon en móbiles-->
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
<!--Barra de navegaICón-->
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

  <?php
if ($gv != - 1) {
    printGroupTable($formatedSelected, $groupVariable, $worksheet, $sv, $radioVal, $acum);
} else {
    printNoGroupTable($formatedSelected, $worksheet, $sv, $radioVal, $acum);
}
?>
  </div>
  <div class="container " style="width:90%!important;display:none">
    <?php
if ($gv != - 1) {
    printGroupTableWord($formatedSelected, $groupVariable, $worksheet, $sv, $radioVal, $acum);
}else{
  printNoGroupTableWord($formatedSelected, $worksheet, $sv, $radioVal, $acum);
}
?>
  </div>

  <a class="btn-floating btn-large waves-effect waves-light  teal lighten-2"  onClick="javascript:history.go(-2)" style="float:left; position:fixed; bottom:10px;left:10px;"><i class="material-icons">replay</i></a>

<?php
if($gv != - 1){
echo '<a onclick="Export2DocGroup()" class="btn-floating btn-large waves-effect waves-light deep-orange lighten-1 pulse" style="float:right; position:fixed; bottom:10px;right:10px;"><i class="material-icons">cloud_download</i></a>';
}else{
  echo '<a onclick="Export2DocNoGroup()" class="btn-floating btn-large waves-effect waves-light deep-orange lighten-1 pulse" style="float:right; position:fixed; bottom:10px;right:10px;"><i class="material-icons">cloud_download</i></a>';
}
 ?>
    <script type="text/javascript" src="../../js/jquery-3.1.1.min.js"></script>
    <script type="text/javascript" src="../../js/materialize.min.js"></script>
    <script type="text/javascript" src="../../js/init.js"></script>
    <script type="text/javascript">
    function Export2DocGroup(){
      var postHtml = "</body></html>";
     var elem = document.getElementById("exportContent").innerHTML+postHtml;
     Export2Doc(elem,'word-content');
    }
    function Export2DocNoGroup(){
      var postHtml = "</body></html>";
     var elem = document.getElementById("exportContentNoGroup").innerHTML+postHtml;
     Export2Doc(elem,'word-content');
    }
    function Export2Doc(element, filename = ''){
        var preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
        var postHtml = "</body></html>";
        var html = preHtml+element;

        var blob = new Blob(['\ufeff', html], {
            type: 'application/msword'
        });

        // SpeICfy link url
        var url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(html);

        // SpeICfy file name
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
