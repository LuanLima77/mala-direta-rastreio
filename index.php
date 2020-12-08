<?php

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

require 'vendor/autoload.php';

Use Automattic\WooCommerce\Client;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\SMTP;
//RENOMEAR NACIONAIS


$inputFileName = './estadualnovembro2.xls';
$outputPlanilha = 'rastreios_estaduais_nao_localizados.ods';
$outputFileName = "./$outputPlanilha";
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
$reader->setReadDataOnly(true);

$readerOds = new \PhpOffice\PhpSpreadsheet\Reader\Ods();

/** Load $inputFileName to a Spreadsheet Object  **/
$spreadsheet = $reader->load($inputFileName);
$workSheet = $spreadsheet->getActiveSheet()->toArray();

//$workSheetToDo= new \PhpOffice\PhpSpreadsheet\Spreadsheet();
//$workSheetToDo->getActiveSheet()->setCellValue('A1', 'DESTINATARIO');
//$workSheetToDo->getActiveSheet()->setCellValue('B1', 'CEP');
//$workSheetToDo->getActiveSheet()->setCellValue('C1', 'UF');
//$workSheetToDo->getActiveSheet()->setCellValue('D1', 'STATUS');

$workSheetToDo = $readerOds->load($outputFileName);
//$teste = $workSheetToDo->getActiveSheet()->toArray();

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($workSheetToDo, "Ods");


//Verificando pedidos do woocommerce
$url = "https://www.literatour.com.br";
$consumer_key = "ck_9e9f6e07f48147b3c6c4cf4b66225e4414a11724";
$consumer_secret ="cs_d79c90ba06f745edafebc270a27d3934682b4014";

$woocommerce = new Client($url, $consumer_key, $consumer_secret);
//Comecei com 16 dias atras, no dia 1/09
$inicioQuinzena = gmdate("o-m-d",strtotime("-17 days")). "T00:00:00";
$fimQuinzena = gmdate("o-m-d",strtotime("-2 days")). "T00:00:00";

//USAR PAGINACAO
//$ontem = "2020-05-31T18:30:00";

$endpoint = "orders";

echo "Pedidos entre  $inicioQuinzena e $fimQuinzena" . "<br>\n";
$parameters = [
    "status" => "processing",
    "after" => $inicioQuinzena,
    "before" => $fimQuinzena,
    "per_page" => 100,
    "page" => 5,
    "order" => "asc"
   
];

$recentcustomers = $woocommerce->get($endpoint, $parameters);
$acertos = 0;


echo count($recentcustomers) . " pedidos recentes encontrados..." . "<br>";

foreach($recentcustomers as $customer)
{
    var_dump($customer->correios_tracking_code);
    exit();
    //echo "DESTINATARIO=>" . $customer->shipping->first_name . "<br>";
    //echo "CEP=>" . $customer->shipping->postcode . "<br>";
    //echo "UF=>" . $customer->shipping->state . "<br>";
  // echo "-------------------------" . "<br>";
  $userFound = 0 ;
  $userFoundName = "";
  $rastreioLocalizado = "";


   foreach($workSheet as $row)
  {
    $rastreio = $row[8];
    $cliente = $customer->shipping->first_name;
    $pedido = $customer->id;
    $data = new DateTime($customer->date_created);

    $plano = $customer->line_items[0]->name;
    $email = $customer->billing->email;
    $row[1] = 'teste';


    $endereco = "CEP: " . $customer->shipping->postcode . ", " . $customer->shipping->address_1 . " " . $customer->shipping->address_2
               . ", " . $customer->shipping->number . ", " . $customer->shipping->neighborhood . ", " . $customer->shipping->city
               . ", " . $customer->shipping->state;

              // echo "DESTINATARIO=>" . $row[2] . "<br>";
    // echo "CEP=>" . $row[3] . "<br>";
    // echo "UF=>" . $row[4] . "<br>";
    // echo "RASTREIO=>" . $row[8] . "<br>";
    if($customer->shipping->postcode == $row[3])
     {
       $userFoundName = $customer->billing->first_name . " " . $customer->billing->last_name  ; 
       $rastreioLocalizado = $rastreio;
       $userFound++;
     }
       
}

  if($userFound == 1)
  {   
    echo  "<b>$userFoundName encontrado(a)!</b>(match simples de CEP ) </br>";
    //enviando email...
     enviarRastreioPorEmail($data,$pedido, $plano,$email,$cliente,$rastreioLocalizado,$endereco);
     //post pedido atualizado com rastreio
     exit();
     $customer->correios_tracking_code = $rastreioLocalizado;
     $woocommerce->post($endpoint, $customer);
     $acertos++;


   }else if($userFound > 1) 
   {
    similar_text(strtolower($customer->billing->first_name), strtolower($row[2]), $percent);
    if($percent > 55)
    {
      echo  "<b>$userFoundName encontrado(a)!</b>(match duplo de CEP e nome por similaridade ) </br>";
      $acertos++;
      //enviando email...
      enviarRastreioPorEmail($data,$pedido, $plano,$cliente,$rastreioLocalizado);
      //post pedido atualizado com email
       $customer->correios_tracking_code = $rastreioLocalizado;
       $woocommerce->post($endpoint, $customer);



    }else
    {
      //echo "MATCH DUPLO PERDIDO ! </br>"; 
      //listagem manual
      echo "<b>NÃO LOCALIZADO -</b>DATA: " . $data->format('d/m/Y H:i:s') . " PEDIDO:  " . $customer->id  ."  NOME : " . 
               $customer->billing->first_name . " " . $customer->billing->last_name . 
               "CEP " . $customer->shipping->postcode ." - UF:  " . $customer->shipping->state . " <br>";

               //marcando na planilha como nao localizado
               $nomeCompleto = $customer->billing->first_name . " " . $customer->billing->last_name;
               $row = $workSheetToDo->getActiveSheet()->getHighestRow()+1;
               $workSheetToDo->getActiveSheet()->insertNewRowBefore($row);
               $workSheetToDo->getActiveSheet()->setCellValue('A'.$row,$nomeCompleto);
               $workSheetToDo->getActiveSheet()->setCellValue('B'.$row,$customer->shipping->postcode);
               $workSheetToDo->getActiveSheet()->setCellValue('C'.$row,$customer->shipping->state);



    // echo "DESTINATARIO=>" . $row[2] . "<br>";
    // echo "CEP=>" . $row[3] . "<br>";
    // echo "UF=>" . $row[4] . "<br>";
    // echo "RASTREIO=>" . $row[8] . "<br>";
    }
  }else
   {
    //listagem manual
    echo "<b>NÃO LOCALIZADO -</b>Data: <b>" . $data->format('d/m/Y H:i:s') . "</b> Pedido:  <b>" . $customer->id  ."</b>  Nome :<b> " . 
    $customer->billing->first_name . " " . $customer->billing->last_name . 
    "</b> CEP <b>" . $customer->shipping->postcode ."</b> - UF: <b>" . $customer->shipping->state . "</b> <br>";

    $nomeCompleto = $customer->billing->first_name . " " . $customer->billing->last_name;
    $row = $workSheetToDo->getActiveSheet()->getHighestRow()+1;
    $workSheetToDo->getActiveSheet()->insertNewRowBefore($row);
    $workSheetToDo->getActiveSheet()->setCellValue('A'.$row,$nomeCompleto);
    $workSheetToDo->getActiveSheet()->setCellValue('B'.$row,$customer->shipping->postcode);
    $workSheetToDo->getActiveSheet()->setCellValue('C'.$row,$customer->shipping->state);  
  }


   echo "-------------------------" . "<br>";
   $writer->save($outputPlanilha);
}


function enviarRastreioPorEmail($data,$pedido, $plano, $email,$cliente,$rastreio,$endereco)
{

 // Instantiation and passing `true` enables exceptions
$mail = new PHPMailer(true);

try {
    //Server settings
    $mail->SMTPDebug = 3;
    $mail->CharSet = 'UTF-8';
    $mail->isSMTP();                                            // Send using SMTP
    $mail->Host       = 'smtp.elasticemail.com';                    // Set the SMTP server to send through
    $mail->SMTPAuth   = true;
    $mail->SMTPSecure = 'tls';
    $mail->Username   = 'edumachion@gmail.com';                     // SMTP username
    $mail->Password   = '2A26BD6A15013BBB3BA1DBCDB4C7E8B4680B';                               // SMTP password
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;         // Enable TLS encryption; `PHPMailer::ENCRYPTION_SMTPS` encouraged
    $mail->Port       = 2525;                                    // TCP port to connect to, use 465 for `PHPMailer::ENCRYPTION_SMTPS` above

    //Recipients
    $mail->setFrom('assessoria@literatour.com.br', 'Literatour');
    $mail->addAddress($email,$cliente);     // Add a recipient
    //$mail->addAddress('ellen@example.com');               // Name is optional
    //$mail->addReplyTo('info@example.com', 'Information');
    //$mail->addCC('cc@example.com');
    $mail->addBCC('77luanlima@gmail.com', 'Luan Lima');

    // Attachments
    //$mail->addAttachment('/var/tmp/file.tar.gz');         // Add attachments
    //$mail->addAttachment('/tmp/image.jpg', 'new.jpg');    // Optional name
//https://rastreamentocorreios.info/consulta/JN420977697BR
    // Content
    $mail->isHTML(true);                                  // Set email format to HTML
    $mail->Subject = 'Literatour - Sua caixinha já está a caminho!';
    $mail->Body    = "
              <html>
              <head><title>Caixinha enviada pelos Correios :)</title></head>
              <body>
                  <div style='background-color: #ff8000; color:snow; font-family: Arial, Helvetica, sans-serif'>
                      <h1>Caixinha chegandooo! uhull  </h1>
                  </div>
                  <div id='content'>
              <p>Olá,$cliente! Sua caixinha já foi enviada!</p>
              <p>Para acompanhar a entrega, use o seguinte código de rastreio:</p>
              <ul> 
                <li><a href='https://rastreamentocorreios.info/consulta/$rastreio'>$rastreio</a>
                </li>
              </ul>
          </div>
              <div style='font-family: Arial, Helvetica, sans-serif'>
                  <h1> Seu Pedido:</h1>
                  <table style='border: 1px solid #ddd;'>
                      <tr>
                        <th>Pedido</th>
                        <th>Plano</th>
                      </tr>
                      <tr>
                        <td>$pedido</td>
                        <td>$plano</td>
                      </tr>
                    
                    </table>

              </div>

              <div style='font-family: Arial, Helvetica, sans-serif'>
              <h1> Sua Entrega:</h1>
              <table style='border: 1px solid #ddd;'>
                  <tr>
                    <th>Endereço</th>
                  </tr>
                  <tr>
                    <td><i>$endereco</i></td>
                  </tr>
                
                </table>

          </div>

              </body>
              </html>";
    $mail->AltBody = 'Aqui está seu rastreio!';
    //$mail->send();    
    
    echo 'Message has been sent';
   } catch (Exception $e) {
    echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
  }

}




