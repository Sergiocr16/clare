<?php

    setcookie("idiomaSelec", $_GET['idioma'], time()+86400);

    header("location: en/index.html");