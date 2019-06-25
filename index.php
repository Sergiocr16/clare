<?php
include './php/clare/globalMethods.php';
deleteUploads();
$_FILES['excelFile']['tmp_name'] = null;
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
    <link type="text/css" rel="stylesheet" href="css/materialize.min.css"/>
    <link rel="stylesheet" type="text/css" href="stylesheet.css">
    <link rel="shortcut icon" href="images/favicon.png" type="image/x-icon">
    <link rel="apple-touch-icon" href="images/favicon.png" sizes="192x192">
</head>

<body class="blue-grey lighten-5">
<!--Barra de navegación-->
   <nav class="white" role="navigation">
    <div class="nav-wrapper container">
      <a id="logo-container" href="index.html" class="brand-logo"><img src="res/logo1.jpg" style="margin-top:7px;" alt="Smiley face" height="50" > </a>
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
   <div id="index-banner" class="parallax-container" style="height:605px!important;">
    <div class="parallax"><img src="res/primera.jpg" alt="Unsplashed background img 1"></div>
    <div class="section no-pad-bot">
      <div class="container " style="margin-top:120px;">
        <br><br>
        <h1 class="header center white-text text-lighten-2">Veloz y sencillo</h1>
        <div class="row center">
          <h5 class="header col s12 light">Seleccione el archivo que desea procesar</h5>
        </div>
        <div class="row center">
          <form action="php/clare/readExcel.php" method="POST" enctype="multipart/form-data">
            <div>
              <h1 style="text-align:center">
              <input class="input-file" id="my-file" type="file" name="excelFile" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel">
              <label tabindex="0" for="my-file" class="btn-large waves-effect waves-light teal lighten-1 input-file-trig">Subir un archivo</label>
              <button id="process" class="btn-large waves-effect waves-light orange darken-4 lighten-1" type="submit" style="margin-left:10px;opacity:0;display:none">Procesar</label>
             </h1>
             </div>
             <h6 style="text-align:center">
            <span class="file-return"></span>
          </h6>
          </form>
        </div>
        <br><br>
      </div>
    </div>
    </div>
  </div>

<!--Seccion de ícono-->
  <div class="container" style="width:90%!important;">
    <div class="section">
      <div class="row">
        <div class="col s12 m4">
          <div class="icon-block">
            <h2 class="center blue-text"><i class="material-icons"  style="color:#26A69A">assignment_returned</i></h2>
            <h5 class="center">Importar tabla Excel</h5>

            <p class="light" style="text-align: center;">
            	Seleccione la tabla de excel de su dispositivo y suba el archivo.
            </p>
          </div>
        </div>

        <div class="col s12 m4">
          <div class="icon-block">
            <h2 class="center blue-text"><i class="material-icons"  style="color:#26A69A">build</i></h2>
            <h5 class="center">Definir grupos y variables</h5>

            <p class="light" style="text-align: center;">
            	Configure las variables de agrupación y la información que desea visualizar en el resultado.
            </p>
          </div>
        </div>

        <div class="col s12 m4">
          <div class="icon-block">
            <h2 class="center blue-text"><i class="material-icons" style="color:#26A69A">done_outline</i></h2>
            <h5 class="center">Visualizar resultados</h5>

            <p class="light" style="text-align: center;">
            	Visualize los resultados agrupados de una manera ágil y sencilla. Exportelos a un archivo Word si así desea.
            </p>
          </div>
        </div>
      </div>
    </div>
  </div>

<!--Segundo banner-->
  <!-- <div class="parallax-container valign-wrapper">
    <div class="section no-pad-bot">
      <div class="container">
        <div class="row center">
          <h5 class="header col s12 light">Eres público, date a conocer cuando quieras</h5>
        </div>
      </div>
    </div>
    <div class="parallax"><img src="res/primera.jpg" alt="Unsplashed background img 2"></div>
  </div> -->

<!--Segunda seccion-->
  <!-- <div class="container">
    <div class="section">

      <div class="row">
        <div class="col s12 center">
          <h3><i class="mdi-content-send brown-text"></i></h3>
          <h4>Contáctanos</h4>
          <p class="center-align light">
          		NoteBlog fue creado por un simple ser humano, no es perfecto, así que si encuentras algún error, o tienes alguna sugerencia, exprésala, eres una gran ayuda para ésta comunidad.
          		Espero que sea de tu agrado y lo disfrutes. <br>
          		Att. JD
          		<br><br></p>
          		<div class="row center">
		          <a href="php/correo/contacto.php" id="download-button" class="btn-large waves-effect waves-light teal lighten-1">Te escuchamos</a>
		        </div>
        </div>
      </div>
    </div>
  </div> -->

<!--Tercer banner-->
  <!-- <div class="parallax-container valign-wrapper">
    <div class="section no-pad-bot">
    <div class="section no-pad-bot">
      <div class="container">
        <div class="row center">
          <h5 class="header col s12 light">No te olvides: Alegra el día de los demás</h5>
        </div>
      </div>
    </div>
    <div class="parallax"><img src="res/Wall3.png" alt="Unsplashed background img 3"></div>
  </div> -->

<!-- <div class="fixed-action-btn horizontal click-to-toggle">
    <a class="btn-floating btn-large red">
      <i class="material-icons">menu</i>
    </a>
    <ul>
      <li><a href="crearCookie.php?idioma=en" class="btn-floating red"><i class="material-icons">g_translate</i></a></li>
    </ul>
 </div> -->

<!--Footer-->
  <footer class="page-footer blue-grey lighten-5">
    <!-- <div class="container">
      <div class="row">
        <div class="col l6 s12">
          <h5 class="blue-text">QuickAdminNoter</h5>
          <p class="blue-text">
          	No importa que tan simple, rara, o inalcanzable <br> parezca la idea, sólo hazlo.
          </p>
        </div>

        <div class="col l3 s12">
          <h5 class="blue-text">Cuenta conmigo</h5>
          <ul>
            <li><a class="blue-text" href="#!">Facebook</a></li>
            <li><a class="blue-text" href="#!">Twitter</a></li>
            <li><a class="blue-text" href="#!">Whatsapp</a></li>
          </ul>
        </div>
      </div>
    </div> -->
    <div class="footer-copyright cyan accent-4" style="background-color:#26A69A!important;">
      <div class="container" >
      <center>
      <h6>Desarrollado por ADITUM CR</h6>
      </center>
      </div>
    </div>
  </footer>

    <script type="text/javascript" src="js/jquery-3.1.1.min.js"></script>
    <script type="text/javascript" src="js/materialize.min.js"></script>
    <script src="script.js"></script>
    <script type="text/javascript" src="js/init.js"></script>

</body>
</html>
