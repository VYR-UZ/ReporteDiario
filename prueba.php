<!doctype html>
<?php
  include 'Classes/PHPExcel.php';
  include 'Classes/PHPExcel/IOFactory.php';
  	//Variable con el nombre del archivo
  	$nombreArchivo = 'Libro1.xlsx';
  	// Cargo la hoja de cálculo
  	$objPHPExcel = PHPExcel_IOFactory::load($nombreArchivo);
    //Asigno la hoja de calculo activa
  	$objPHPExcel->setActiveSheetIndex(0);
  	//Obtengo el numero de filas del archivo
  	$numRows = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();
    //-------------------B
  		$b7 = $objPHPExcel->getActiveSheet()->getCell('B'.'7')->getCalculatedValue();
      $b8 = $objPHPExcel->getActiveSheet()->getCell('B'.'8')->getCalculatedValue();
      $b15 = $objPHPExcel->getActiveSheet()->getCell('B'.'15')->getCalculatedValue();
      $b16 = $objPHPExcel->getActiveSheet()->getCell('B'.'16')->getCalculatedValue();
      $b18 = $objPHPExcel->getActiveSheet()->getCell('B'.'18')->getCalculatedValue();
      $b19 = $objPHPExcel->getActiveSheet()->getCell('B'.'19')->getCalculatedValue();
      $b20 = $objPHPExcel->getActiveSheet()->getCell('B'.'20')->getCalculatedValue();
      $b21 = $objPHPExcel->getActiveSheet()->getCell('B'.'21')->getCalculatedValue();
      $b24 = $objPHPExcel->getActiveSheet()->getCell('B'.'24')->getCalculatedValue();
      $b25 = $objPHPExcel->getActiveSheet()->getCell('B'.'25')->getCalculatedValue();
      $b27 = $objPHPExcel->getActiveSheet()->getCell('B'.'27')->getCalculatedValue();
      $b28 = $objPHPExcel->getActiveSheet()->getCell('B'.'28')->getCalculatedValue();
      $b29 = $objPHPExcel->getActiveSheet()->getCell('B'.'29')->getCalculatedValue();
      $b30 = $objPHPExcel->getActiveSheet()->getCell('B'.'30')->getCalculatedValue();
      $b34 = $objPHPExcel->getActiveSheet()->getCell('B'.'34')->getCalculatedValue();

      //------------------C
      $c5 = $objPHPExcel->getActiveSheet()->getCell('C'.'5')->getCalculatedValue();
      $c10 = $objPHPExcel->getActiveSheet()->getCell('C'.'10')->getCalculatedValue();
      $c11 = $objPHPExcel->getActiveSheet()->getCell('C'.'11')->getCalculatedValue();
      $c12 = $objPHPExcel->getActiveSheet()->getCell('C'.'12')->getCalculatedValue();


      //-------------------------D
      $d34 = $objPHPExcel->getActiveSheet()->getCell('D'.'34')->getCalculatedValue();
      $d35 = $objPHPExcel->getActiveSheet()->getCell('D'.'35')->getCalculatedValue();
      $d36 = $objPHPExcel->getActiveSheet()->getCell('D'.'36')->getCalculatedValue();
      $d37 = $objPHPExcel->getActiveSheet()->getCell('D'.'37')->getCalculatedValue();
      $d38 = $objPHPExcel->getActiveSheet()->getCell('D'.'38')->getCalculatedValue();
      //----------------------------E
      $e7 = $objPHPExcel->getActiveSheet()->getCell('E'.'7')->getCalculatedValue();
      $e10 = $objPHPExcel->getActiveSheet()->getCell('E'.'10')->getCalculatedValue();
      $e11 = $objPHPExcel->getActiveSheet()->getCell('E'.'11')->getCalculatedValue();
      $e12 = $objPHPExcel->getActiveSheet()->getCell('E'.'12')->getCalculatedValue();
      $e33 = $objPHPExcel->getActiveSheet()->getCell('E'.'33')->getCalculatedValue();
      $e34 = $objPHPExcel->getActiveSheet()->getCell('E'.'34')->getCalculatedValue();
      $e35 = $objPHPExcel->getActiveSheet()->getCell('E'.'35')->getCalculatedValue();
      $e36 = $objPHPExcel->getActiveSheet()->getCell('E'.'36')->getCalculatedValue();
      $e37 = $objPHPExcel->getActiveSheet()->getCell('E'.'37')->getCalculatedValue();
      $e38 = $objPHPExcel->getActiveSheet()->getCell('E'.'38')->getCalculatedValue();

      //---------------------------F
      $f3 = $objPHPExcel->getActiveSheet()->getCell('F'.'3')->getCalculatedValue();
      $f15 = $objPHPExcel->getActiveSheet()->getCell('F'.'15')->getCalculatedValue();
      $f16 = $objPHPExcel->getActiveSheet()->getCell('F'.'16')->getCalculatedValue();
      $f18 = $objPHPExcel->getActiveSheet()->getCell('F'.'18')->getCalculatedValue();
      $f19 = $objPHPExcel->getActiveSheet()->getCell('F'.'19')->getCalculatedValue();
      $f20 = $objPHPExcel->getActiveSheet()->getCell('F'.'20')->getCalculatedValue();
      $f21 = $objPHPExcel->getActiveSheet()->getCell('F'.'21')->getCalculatedValue();
      $f24 = $objPHPExcel->getActiveSheet()->getCell('F'.'24')->getCalculatedValue();
      $f25 = $objPHPExcel->getActiveSheet()->getCell('F'.'25')->getCalculatedValue();
      $f27 = $objPHPExcel->getActiveSheet()->getCell('F'.'27')->getCalculatedValue();
      $f28 = $objPHPExcel->getActiveSheet()->getCell('F'.'28')->getCalculatedValue();
      $f29 = $objPHPExcel->getActiveSheet()->getCell('F'.'29')->getCalculatedValue();
      $f30 = $objPHPExcel->getActiveSheet()->getCell('F'.'30')->getCalculatedValue();
      $f34 = $objPHPExcel->getActiveSheet()->getCell('F'.'34')->getCalculatedValue();
      $f35 = $objPHPExcel->getActiveSheet()->getCell('F'.'35')->getCalculatedValue();
      $f36 = $objPHPExcel->getActiveSheet()->getCell('F'.'36')->getCalculatedValue();
      $f37 = $objPHPExcel->getActiveSheet()->getCell('F'.'37')->getCalculatedValue();
      $f38 = $objPHPExcel->getActiveSheet()->getCell('F'.'38')->getCalculatedValue();

      //---------------------G
      $g15 = $objPHPExcel->getActiveSheet()->getCell('G'.'15')->getCalculatedValue();
      $g16 = $objPHPExcel->getActiveSheet()->getCell('G'.'16')->getCalculatedValue();
      $g18 = $objPHPExcel->getActiveSheet()->getCell('G'.'18')->getCalculatedValue();
      $g19 = $objPHPExcel->getActiveSheet()->getCell('G'.'19')->getCalculatedValue();
      $g20 = $objPHPExcel->getActiveSheet()->getCell('G'.'20')->getCalculatedValue();
      $g21 = $objPHPExcel->getActiveSheet()->getCell('G'.'21')->getCalculatedValue();
      $g24 = $objPHPExcel->getActiveSheet()->getCell('G'.'24')->getCalculatedValue();
      $g25 = $objPHPExcel->getActiveSheet()->getCell('G'.'25')->getCalculatedValue();
      $g27 = $objPHPExcel->getActiveSheet()->getCell('G'.'27')->getCalculatedValue();
      $g28 = $objPHPExcel->getActiveSheet()->getCell('G'.'28')->getCalculatedValue();
      $g29 = $objPHPExcel->getActiveSheet()->getCell('G'.'29')->getCalculatedValue();
      $g30 = $objPHPExcel->getActiveSheet()->getCell('G'.'30')->getCalculatedValue();
      $g34 = $objPHPExcel->getActiveSheet()->getCell('G'.'34')->getCalculatedValue();

      //FECHA_DE_AYER
      $arrayMeses = array('Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
         'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre');

      $arrayDias = array( 'Domingo', 'Lunes', 'Martes',
           'Miércoles', 'Jueves', 'Viernes', 'Sábado');

      $arrayDiasPlural = array( 'domingos', 'lunes', 'martes',
                'miércoles', 'jueves', 'viernes', 'sábados');

        $fechaAyer =  $arrayDias[date('w',strtotime("-1 day") )].", ".date( 'd', strtotime("-1 day") ) ." de ".$arrayMeses[date('m')-1]." de ".date('Y');
      $diaAyer = $arrayDiasPlural[date('w',strtotime("-1 day"))];
      //FECHA_DE_AYER
  ?>

<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title></title>
    <link rel="stylesheet" href="style.css">
    <style media="screen">
      .odd th, .odd td {
        background: #eee;
      }
    </style>
  </head>
  <body>
    <center>
      <table>
        <tr>
          <td width="750"><img src="img/logo-milenio-Horiz.jpg" alt="Grupo Milenio" width="375" height="65"></td>
            <td><font color = "grey"><h2><?php echo $f3; ?></h2></font></td>
        </tr>
      </table>
    <font color = "grey" face = "Calibri">


      <section>
        <div class="contenido fecha">
          <p><center>
            <font color="blue" face="calibri" size ="5px"><?php echo $fechaAyer?></font>
          </p></div>
      </section>
    <section>
      <div class="contenido descripcion">
        <p><font size ="5px"><?php echo $b7 ;?></font>
          <?php
          if($c12 < 0 && $e12 < 0 || $c12< 0 && $e12 >=0 || $c12>= 0 && $e12 <0){?>
            <font color="red" size ="5px"><b><?php echo "por abajo del promedio" ;?></b></font>
          <?php
          }
          if($c12 >= 0 && $e12 >= 0){?>
            <font color="green" size ="5px"><b><?php echo "por arriba del promedio" ;?></b></font>
            <?php
          }
          ?>
          <br>
          <font size ="5px"><?php echo "en comparación de los últimos cuatro ".$diaAyer;?></p></font>
      </div>
    </section>
    <br>
    <br>
    <!-- TABLA 1-->

      <table class="resp">
        <thead class="thead">
        <tr id="tr" >
          <th scope="col" width="200"><?php echo $c10 ;?></th>
          <th scope="col"><?php echo $e10 ;?></th>
        </tr>
        </thead>
        <tbody id="tbody">
          <tr id="tr">
            <td id="td" align="center"> <font color="blue"><h1><?php echo number_format($c11) ;?></h1></font></td>
            <td id="td" align="center"> <font color="blue"><h1><?php echo number_format($e11) ;?></h1></font></td>
          </tr>
        <tr id="tr">
                <?php
          if($c12 < 0  ){?>
            <td id="td" align="center"><font color="red" size ="5px"><b><?php echo round($c12*100)."%";?></b></font></td>
          <?php
          }
          if($c12 >= 0){?>
            <td id="td" align="center"><font color="green" size ="5px"><b><?php echo round($c12*100)."%";?></b></font></td>
            <?php
          }
          ?>
          <?php
            if($e12 < 0){?>
              <td id="td" align="center"><font color="red" size ="5px"><b><?php echo round($e12*100)."%";?></b></font></td>
            <?php
            }
            if($e12 >= 0){?>
              <td id="td" align="center"><font color="green" size ="5px"><b><?php echo round($e12*100)."%";?></b></font></td>
              <?php
            }
          ?>
        </tr>
    </table>
    <br>
    <br>

    <!-- TABLA 2 -->
    <table class="resp" class="rwd_auto fontsize">
      <thead class="thead">
      <tr id="tr">
        <td scope="col" width="650"><font align = "left"><h3><b><?php echo $b15 ;?></b></h3></td>
        <td scope="col" width="100"><?php echo $f15 ;?></td>
        <td scope="col"><?php echo $g15 ;?></td>
      </tr>
      </thead>
      <tbody id="tbody">
        <tr id="tr" class="odd">
          <td id="td" ><font color="darkblue"><h3><b><?php echo $b16 ;?></h3></font></td>
          <td id="td"><font color="CornflowerBlue"><?php echo number_format($f16);?></td>
          <td id="td"><?php echo $g16 ;?></td>
        </tr>
      <tr id="tr">
        <td id="td"><font color="darkblue"><?php echo $b18 ;?></td>
        <td id="td"><font color="CornflowerBlue"><?php echo number_format($f18) ;?></td>
        <td id="td"><?php echo $g18 ;?></td>
      </tr>
      <tr id="tr" class="odd" >
        <td id="td"><font color="darkblue"><?php echo $b19 ;?></td>
        <td id="td"><font color="CornflowerBlue"><?php echo number_format($f19) ;?></td>
        <td id="td"><?php echo $g19 ;?></td>
      </tr>
      <tr id="tr">
        <td id="td"><font color="darkblue"><?php echo $b20 ;?> </td>
        <td id="td"><font color="CornflowerBlue"><?php echo number_format($f20) ;?></td>
        <td id="td"><?php echo $g20 ;?></td>
      </tr>
      <tr id="tr" class="odd">
        <td id="td"><font color="darkblue"><?php echo $b21 ;?></td>
        <td id="td"><font ><font color="CornflowerBlue"><?php echo number_format($f21) ;?></td>
        <td id="td"><?php echo $g21 ;?></td>
      </tr>
      </tbody>
    </table>
    <br>
    <br>
    <!-- TABLA 3-->
    <table class="resp">
    <thead class="thead">
      <tr id="tr">
        <td id="td" scope="col" width="650"><font align = "left"><h3><?php echo $b24 ;?></h3></td>
        <td id="td" scope="col" width="100"><?php echo $f24 ;?></td>
        <td id="td"><?php echo $g24 ;?></td>
      </tr>
    </thead>
      <tbody id="tbody">
      <tr id="tr" class="odd">
        <td id="td"><font color="darkblue" ><h3><?php echo $b25 ;?></h3></font></td>
        <td id="td"><font color="CornflowerBlue"><?php echo number_format($f25) ;?></td>
        <td id="td"><?php echo $g25 ;?></td>
      </tr>
      <tr id="tr">
        <td id ="td"><font color="darkblue"><?php echo $b27 ;?></td>
        <td id ="td"><font color="CornflowerBlue"><?php echo number_format($f27 );?></td>
        <td id ="td"><?php echo $g27 ;?></td>
      </tr>
      <tr id="tr" class="odd">
        <td id ="td"><font color="darkblue"><?php echo $b28 ;?> </td>
        <td id ="td"><font color="CornflowerBlue"><?php echo number_format($f28) ;?></td>
        <td id ="td"><?php echo $g28 ;?></td>
      </tr>
      <tr id="tr">
        <td id ="td"><font color="darkblue"><?php echo $b29 ;?></td>
        <td id ="td"><font color="CornflowerBlue"><?php echo number_format($f29) ;?></td>
        <td id ="td"><?php echo $g29 ;?></td>
      </tr>
      <tr id="tr" class="odd">
        <td id ="td"><font color="darkblue"><?php echo $b30 ;?> </td>
        <td id ="td"><font color="CornflowerBlue"><?php echo number_format($f30) ;?></td>
        <td id ="td"><?php echo $g30 ;?></td>
      </tr>
      </tbody>
    </table>
    <br>
    <br>
    <!-- TABLA 4-->
    <table class="resp">
      <tr id="tr">
        <td id ="td">
          <th rowspan="3">
            <th>
              <th>
              <th><?php echo $e33 ;?></th>
              </th>
            </th>
          </th>
        </td>
      </tr>
      <tr>
        <td id ="td">
          <td id ="td" scope="col" width="400" ><font align = "left"><h3><?php echo $b34;?></th>
          <td id ="td" scope="col" width="150" bgcolor="#eee"><font color="darkblue" ><h3> <?php echo $d34 ;?></h3></td>
          <td id ="td" scope="col" width="100" bgcolor="#eee"><font color="CornflowerBlue"><?php echo number_format($e34) ;?></td>
            <?php
            if($f34 < 0){?>
              <td id ="td" bgcolor="#eee"><font color="red"><?php echo round($f34*100)."%";?></td>
            <?php
            }
            if($f34 >= 0 ){?>
              <td id ="td" bgcolor="#eee"><font color="green"><?php echo round($f34*100)."%";?></td>
              <?php
            }
            ?>
          <td id ="td" scope="col" width="130" bgcolor="#eee"><?php echo $g34 ;?></td>
        </td>
      </tr>
      <tr>
        <th colspan="3">
        <td id ="td"><font color="darkblue"><h3> <?php echo $d35 ;?></h3></td>
        <td id ="td"><font color="CornflowerBlue"><?php echo number_format($e35) ;?></td>
          <?php
          if($f35 < 0){?>
            <td id="td"><font color="red"><?php echo round($f35*100)."%";?></td>
          <?php
          }
          if($f35 >= 0 ){?>
            <td id="td"><font color="green"><?php echo round($f35*100)."%";?></td>
            <?php
          }
          ?>
        </th>
      </tr>
      <tr>
        <th colspan="3">
        <td bgcolor="#eee"><font color="darkblue"><h3> <?php echo $d36 ;?></h3></td>
        <td bgcolor="#eee"><font color="CornflowerBlue"><?php echo number_format($e36) ;?></td>
          <?php
          if($f36 < 0){?>
            <td id="td" bgcolor="#eee"><font color="red"><?php echo round($f36*100)."%";?></td>
          <?php
          }
          if($f36 >= 0 ){?>
            <td id="td" bgcolor="#eee"><font color="green"><?php echo round($f36*100)."%";?></td>
            <?php
          }
          ?>
        </th>
      </tr>
      <tr>
        <th colspan="3">
        <td id="td"><font color="darkblue"><h3> <?php echo $d37 ;?></h3></td>
        <td id="td"><font color="CornflowerBlue"><?php echo number_format($e37) ;?></td>
          <?php
          if($f37 < 0){?>
            <td id="td"><font color="red"><?php echo round($f37*100)."%";?></td>
          <?php
          }
          if($f37 >= 0 ){?>
            <td id="td"><font color="green"><?php echo round($f37*100)."%";?></td>
            <?php
          }
          ?>
        </th>
      </tr>
      <tr>
        <th id="td" colspan="3">
        <td id="td" bgcolor="#eee"><font color="darkblue"><h3> <?php echo $d38 ;?></h3></td>
        <td id="td" bgcolor="#eee"><font color="CornflowerBlue"><?php echo number_format($e38) ;?></td>
          <?php
          if($f38 < 0){?>
            <td id="td" bgcolor="#eee"><font color="red"><?php echo round($f38*100)."%";?></td>
          <?php
          }
          if($f38 >= 0 ){?>
            <td id="td" bgcolor="#eee"><font color="green"><?php echo round($f38*100)."%";?></td>
            <?php
          }
          ?>
        </th>
      </tr>
    </table>
  </center>
  </body>
</html>
