<?php

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

require 'vendor/autoload.php';

Use Automattic\WooCommerce\Client;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\SMTP;



$inputFileName = './mala-estadual.xls';
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
$reader->setReadDataOnly(true);

/** Load $inputFileName to a Spreadsheet Object  **/
$spreadsheet = $reader->load($inputFileName);
$workSheet = $spreadsheet->getActiveSheet()->toArray();


//var_dump($workSheet);

foreach($workSheet as $row)
{

    //echo "DESTINATARIO=>" . $row[2] . "<br>";
    //echo "CEP=>" . $row[3] . "<br>";
    //echo "UF=>" . $row[4] . "<br>";
//echo "-------------------------" . "<br>";

}


//Verificando pedidos do woocommerce
$url = "https://www.literatour.com.br";
$consumer_key = "ck_9e9f6e07f48147b3c6c4cf4b66225e4414a11724";
$consumer_secret ="cs_d79c90ba06f745edafebc270a27d3934682b4014";

$woocommerce = new Client($url, $consumer_key, $consumer_secret);

$ontem = gmdate("o-m-d",strtotime("-15 days")). "T00:00:00";
$hoje = gmdate("o-m-d"). "T00:00:00";

//$ontem = "2020-05-31T18:30:00";

$endpoint = "orders";

echo "Iniciando coletas do dia $ontem" . "<br>\n";
$parameters = [
    "status" => "processing",
    "after" => $ontem,
    "before" => $hoje,
    "per_page" => 100,
    "order" => "asc"
   
];

$recentcustomers = $woocommerce->get($endpoint, $parameters);

echo count($recentcustomers) . " USUÁRIOS ENCONTRADOS..." . "<br>";



foreach($recentcustomers as $customer)
{
    //var_dump($customer->shipping);
    //echo "DESTINATARIO=>" . $customer->shipping->first_name . "<br>";
    //echo "CEP=>" . $customer->shipping->postcode . "<br>";
    //echo "UF=>" . $customer->shipping->state . "<br>";
  // echo "-------------------------" . "<br>";
  $userFound = 0 ;
  $userFoundName = "";

   foreach($workSheet as $row)
{
  if($customer->shipping->postcode == $row[3])
  {
     $userFoundName = $customer->shipping->first_name; 
     $userFound++;
     echo  "MATCH DE CEP - ASSINANTE $userFoundName! </br>";

  }
    //echo "DESTINATARIO=>" . $row[2] . "<br>";
    //echo "CEP=>" . $row[3] . "<br>";
   // echo "UF=>" . $row[4] . "<br>";
//echo "-------------------------" . "<br>";

}

if($userFound == 1)
{
    echo "MATCH PERFEITO! <br><br>";
    //enviando email...
   

// Instantiation and passing `true` enables exceptions
$mail = new PHPMailer(true);

try {
    //Server settings
    $mail->SMTPDebug = 3;
    $mail->isSMTP();                                            // Send using SMTP
    $mail->Host       = 'smtp.zoho.com';                    // Set the SMTP server to send through
    $mail->SMTPAuth   = true;
    $mail->SMTPSecure = 'ssl';
    $mail->Username   = 'assessoria@literatour.com.br';                     // SMTP username
    $mail->Password   = 'desapegando#2020@';                               // SMTP password
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;         // Enable TLS encryption; `PHPMailer::ENCRYPTION_SMTPS` encouraged
    $mail->Port       = 587;                                    // TCP port to connect to, use 465 for `PHPMailer::ENCRYPTION_SMTPS` above

    //Recipients
    $mail->setFrom('assessoria@literatour.com.br', 'Literatour');
    $mail->addAddress('77luanlima@gmail.com', 'Luan Lima');     // Add a recipient
    //$mail->addAddress('ellen@example.com');               // Name is optional
    //$mail->addReplyTo('info@example.com', 'Information');
    //$mail->addCC('cc@example.com');
    //$mail->addBCC('bcc@example.com');

    // Attachments
    //$mail->addAttachment('/var/tmp/file.tar.gz');         // Add attachments
    //$mail->addAttachment('/tmp/image.jpg', 'new.jpg');    // Optional name

    // Content
    $mail->isHTML(true);                                  // Set email format to HTML
    $mail->Subject = 'Literatour - Sua caixinha foi enviada!';
    $mail->Body    = 'Aqui está seu rastreio!';
    $mail->AltBody = 'Aqui está seu rastreio!';

    $mail->send();
    echo 'Message has been sent';
} catch (Exception $e) {
    echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
}

exit();

}else if($userFound > 1)
{
 // $var_1 = "Luan";
//$var_2 = "LUSN";
//similar_text(strtolower($var_1), strtolower($var_2), $percent);
//echo "PERCENT É $percent <br>";
 
similar_text(strtolower($customer->shipping->first_name), strtolower($row[2]), $percent);
echo "SIMILARIDADE E $percent <br>";
  if($percent > 55)
  {
      echo "MATCH DUPLO ENCONTRADO ! </br>";
     //enviando email...

  }else
  {
    echo "MATCH DUPLO PERDIDO ! </br>";

  }
}



}







