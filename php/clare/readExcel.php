<!DOCTYPE html>
<html lang="es">
<head>
    <title>Clare</title>
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

<body class="">
<!--Barra de navegación-->
   <nav class="white" role="navigation">
    <div class="nav-wrapper container">
      <a id="logo-container" href="/" class="brand-logo"><img src="../../res/logo1.jpg" style="margin-top:7px;" alt="Smiley face" height="50" > </a>
      <ul class="right hide-on-med-and-down">
      <span style="color:black;">Dr. Roy Wong</span>
      </ul>
      <!-- <ul id="nav-mobile" class="side-nav">
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
        <h5 class="header center white-text text-lighten-2">Lectura de tabla</h1>
      </div>
    </div>
  </div>
  </div>
</div>

   <div>
       <?php
require_once "PHPExcel.php";
include 'globalMethods.php';
deleteUploads();
function printForm($headings) {
    echo '<div class="container">';
    echo '<form  action="/php/clare/formatExcel.php" method="POST" enctype="multipart/form-data">';
    echo '<h6>Variable de grupo:</h6>';
    echo '<select  class="mdl-textfield__input validate" id="groupVariable" name="groupVariable" >';
    echo '<option value="" disabled selected>Seleccione la variable de grupo</option>';
    echo '<option value="-1">Sin variable de grupo</option>';
    for ($i = 0;$i < count($headings);$i++) {
        if ($headings[$i] != null) {
            echo '<option value="' . $i . '">' . $headings[$i] . '</option>';
        }
    }
    echo '</select>';
    echo '<h6>Variables de seleccion:</h6>';
    for ($i = 0;$i < count($headings);$i++) {
        if ($headings[$i] != null) {
            echo '<select  class="mdl-textfield__input validate" id="selected" name="selected[]" required multiple>';
            echo '<option value="" disabled selected>Seleccione las variables de selección</option>';
            for ($i = 0;$i < count($headings);$i++) {
                if ($headings[$i] != null) {
                    echo '<option value="' . $i . '">' . $headings[$i] . '</option>';
                }
            }
            echo '</select>';
        }
    }
    echo '<h6>Filtrar por:</h6>';
    echo '<div class="left">';
    echo ' <p>';
    echo '   <label>';
    echo '     <input class="with-gap" value="1" name="type" type="radio" checked />';
    echo '     <span>Solo datos</span>';
    echo '   </label>';
    echo ' </p>';
    echo ' <p>';
    echo '   <label>';
    echo '     <input class="with-gap" value="2" name="type" type="radio" />';
    echo '     <span>Porcentaje</span>';
    echo '   </label>';
    echo ' </p>';
    echo ' <p>';
    echo '   <label>';
    echo '     <input class="with-gap" value="3" name="type" type="radio"  />';
    echo '     <span>Porcentaje e intervalo</span>';
    echo '   </label>';
    echo '</div>';
    echo '<div class="right">';
    echo '<p >';
    echo '<label>';
    echo '<input type="checkbox" value="false" class="filled-in" id="acumulated" name="acumulated"/>';
    echo '<span>Mostrar acumulado</span>';
    echo '</label>';
    echo ' <p>';
    echo '   <label>';
    echo '     <input class="with-gap" value="1" min="0" name="decimalNumber" type="number"/>';
    echo '     <span>Número de decimales</span>';
    echo '   </label>';
    echo ' </p>';
    echo ' </p>';
    echo '</div>';
    echo '<div class="col s12" style="margin-top:20px;padding:0!important">';
    echo '<button id="btnsubmit" class="btn-large waves-effect waves-light teal lighten-1 right" type="submit">Procesar</button>';
    echo '<button class="btn-large waves-effect waves-light left deep-orange lighten-1" onClick="javascript:history.go(-1)" type="button" >Atrás</button>';
    echo '</div>';
    echo '</form>';
    echo '</div>';
}
function printTable($worksheet) {
    $lastRow = $worksheet->getHighestRow();
    $lastColumn = $worksheet->getHighestColumn();
    $lastColumnIndex = PHPExcel_Cell::columnIndexFromString($lastColumn);
    $headings = getHeadings($worksheet);
    echo '<div style="overflow-x:auto;max-height:400px";margin:30px;>';
    echo "<table class='responsive-table'>";
    echo "<thead>";
    echo "<tr>";
    for ($i = 0;$i <= count($headings) - 1;$i++) {
        if ($headings[$i] != null) {
            echo '<th>';
            echo $headings[$i];
            echo '</th>';
        }
    }
    echo "</thead>";
    echo "<tbody>";
    for ($row = 2;$row <= $lastRow;$row++) {
        echo "<tr>";
        for ($column = 0;$column <= $lastColumnIndex;$column++) {
            $columString = PHPExcel_Cell::stringFromColumnIndex($column);
            $value = $worksheet->getCell($columString . $row)->getValue();
            echo "<td>";
            echo $value;
            echo "</td>";
        }
        echo "</tr>";
    }
    echo "</tbody>";
    echo "</table>";
    echo '</div>';
}
if (isset($_FILES['excelFile']) && !empty($_FILES['excelFile']['tmp_name'])) {
    $tmpfname = $_FILES['excelFile']['tmp_name'];
    $excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
    $excelObj = $excelReader->load($tmpfname);
    $worksheet = $excelObj->getSheet(0);
    $headings = getHeadings($worksheet);
    // Leemos los headings
    printTable($worksheet);
    $target_dir = "../../uploads/";
    $target_file = $target_dir . basename($tmpfname);
    $uploadOk = 1;
    $imageFileType = strtolower(pathinfo($target_file, PATHINFO_EXTENSION));
    // Check if image file is a actual image or fake image
    if (move_uploaded_file($tmpfname, $target_file)) {
        // Se subio

    } else {
        // No subio

    }
}
?>
  </div>


<!--Seccion de ícono-->

  <div class="container" style="width:90%!important;">
    <div class="section">
      <div class="row">
        <div class="col s12 m12">
          <div class="icon-block">
            <h2 class="center blue-text"><i class="material-icons"  style="color:#26A69A">build</i></h2>
            <h5 class="center">Definir grupos y variables</h5>
          </div>
        </div>
        <div class="col s12 m12">
        <?php
printForm($headings);
?>
        </div>
      </div>
    </div>
  </div>

  <footer class="page-footer lighten-5" style="margin-top:30px;background:white!important">

    <div class="footer-copyright cyan accent-4" style="background-color:#26A69A!important;">
      <div class="container" >
      <center>
      <h6>Desarrollado por ADITUM CR</h6>
      </center>
      </div>
    </div>
  </footer>
    <script type="text/javascript" src="../../js/jquery-3.1.1.min.js"></script>
    <script type="text/javascript" src="../../js/materialize.min.js"></script>
    <script src="../../script.js"></script>
    <script type="text/javascript" src="../../js/init.js"></script>
<script type="text/javascript">
$(document).ready(function() {
 $('select').formSelect();
$('#btnsubmit').prop('disabled',true);
$('#groupVariable').on('change', function () {
  if($('#groupVariable').val() && $('#selected').val().length>0){
    $('#btnsubmit').prop('disabled',false);
  }else{
    $('#btnsubmit').prop('disabled',true);
  }
}).trigger('change');
$('#selected').on('change', function () {
  if($('#groupVariable').val() && $('#selected').val().length>0){
    $('#btnsubmit').prop('disabled',false);
  }else{
    $('#btnsubmit').prop('disabled',true);

  }
}).trigger('change');
});

</script>
</body>
</html>
