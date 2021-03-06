<?php
ini_set('max_execution_time', '0'); // for infinite time of execution 
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);



require 'vendor/autoload.php';

Use Automattic\WooCommerce\Client;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\SMTP;


$inputFileName = './nacional1deabril.xls';
//$outputPlanilha = 'rastreios_estaduais_nao_localizados.ods';
$outputPlanilha = 'rastreios_nacionais_nao_localizados.ods';

$outputFileName = "./$outputPlanilha";
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
$reader->setReadDataOnly(true);

$readerOds = new \PhpOffice\PhpSpreadsheet\Reader\Ods();

/** Load $inputFileName to a Spreadsheet Object  **/
$spreadsheet = $reader->load($inputFileName);
$workSheet = $spreadsheet->getActiveSheet()->toArray();



$workSheetToDo = $readerOds->load($outputFileName);
$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($workSheetToDo, "Ods");


//Verificando pedidos do woocommerce
$url = "https://www.literatour.com.br";
$consumer_key = "key";
$consumer_secret ="secret";

$woocommerce = new Client($url, $consumer_key, $consumer_secret);
//Comecei com 16 dias atras, no dia 1/03
$inicioQuinzena = gmdate("Y-m-d",strtotime("-16 days")). "T00:00:00";
$fimQuinzena = gmdate("Y-m-d",strtotime("-1 days")). "T00:00:00";

//USAR PAGINACAO
//$ontem = "2020-05-31T18:30:00";

$endpoint = "orders";
$totaldados = 1;

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
    $mail->Username   = 'email';                     // SMTP username
    $mail->Password   = 'pass';                               // SMTP password
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;         // Enable TLS encryption; `PHPMailer::ENCRYPTION_SMTPS` encouraged
    $mail->Port       = 2525;                                    // TCP port to connect to, use 465 for `PHPMailer::ENCRYPTION_SMTPS` above

    //Recipients
    $mail->setFrom('assessoria@literatour.com.br', 'Literatour');
    $mail->addAddress($email,$cliente);     // Add a recipient
    //$mail->addAddress('ellen@example.com');               // Name is optional
    //$mail->addReplyTo('info@example.com', 'Information');
    //$mail->addCC('cc@example.com');
    $mail->addBCC('77luanlima@gmail.com', 'Luan Lima');
//https://rastreamentocorreios.info/consulta/JN420977697BR
    // Content
    $mail->isHTML(true);                                  // Set email format to HTML
    $mail->Subject = 'Literatour - Sua caixinha já foi enviada!!';
    $mail->Body    = "
              <html>
              <head><title>Caixinha postada! :) </title></head>
              <body>
                  <div style='background-color: #ff8000; color:snow; font-family: Arial, Helvetica, sans-serif'>
                      <h1>Caixinha chegandoo!! Uhull!  </h1>
                  </div>
                  <div id='content'>
              <p>Olá,$cliente! Seu kit já está a caminho via Correios :)</p>
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
    $mail->send();    
    echo 'Email enviado <br>';
	sleep(10);
   } catch (Exception $e) {
    echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
  }

}
$j = 1;
$totaldados = 100;
for($i = 1;$totaldados > 0;$i++){
		echo "Pedidos entre  $inicioQuinzena e $fimQuinzena" . "<br>\n";
$parameters = [
    "status" => "processing",
    "after" => $inicioQuinzena,
    "before" => $fimQuinzena,
    "per_page" => 100,
    "page" => $i,
    "order" => "asc"
   
];

$recentcustomers = $woocommerce->get($endpoint, $parameters);
$acertos = 0;
$totaldados  = count($recentcustomers);
if($totaldados > 0){
echo $totaldados." pedidos recentes encontrados..." . "<br>";

//var_dump($recentcustomers);
foreach($recentcustomers as $customer)
{
    
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
    //echo '<hr> RASTREIO: '.$row[8];
    $cliente = $customer->shipping->first_name;
    $pedido = $customer->id;
    $data = new DateTime($customer->date_created);

    $plano = $customer->line_items[0]->name;
    
    $email = $customer->billing->email;


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
    
	 //$rastreio_final = (string) $rastreioLocalizado;
     $customerId = $customer->id;
     $dados = [
	   'correios_tracking_code' => "$rastreioLocalizado"
	   ];
	   
       $woocommerce->put("$endpoint/$customerId", $dados);
     echo  "<b>$userFoundName adicionado rastreio no painel do usuário</b> </br>";

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
	   $rastreio_final = (string) $rastreioLocalizado;
       $customerId = $customer->id;
	   $dados = [
	   'correios_tracking_code' => "$rastreioLocalizado"
	   ];
	   
       $woocommerce->put("$endpoint/$customerId", $dados);
       echo  "<\br><b>$userFoundName adicionado rastreio no painel do usuário</b> </br>";

    }else
    {
      //echo "MATCH DUPLO PERDIDO ! </br>"; 
      //listagem manual
      echo "<b>Página: $j - NÃO LOCALIZADO -</b>DATA: " . $data->format('d/m/Y H:i:s') . " PEDIDO:  " . $customer->id  ."  NOME : " . 
               $customer->billing->first_name . " " . $customer->billing->last_name . 
               "CEP " . $customer->shipping->postcode ." - UF:  " . $customer->shipping->state . " <br>";

               //marcando na planilha como nao localizado
               $nomeCompleto = $customer->billing->first_name . " " . $customer->billing->last_name;
               $row = $workSheetToDo->getActiveSheet()->getHighestRow()+1;
               $workSheetToDo->getActiveSheet()->insertNewRowBefore($row);
               $workSheetToDo->getActiveSheet()->setCellValue('A'.$row,$nomeCompleto);
               $workSheetToDo->getActiveSheet()->setCellValue('B'.$row,$customer->shipping->postcode);
               $workSheetToDo->getActiveSheet()->setCellValue('C'.$row,$customer->shipping->state);
    }
  }else
   {
    //listagem manual
    echo "<b>Página: $j - NÃO LOCALIZADO -</b>Data: <b>" . $data->format('d/m/Y H:i:s') . "</b> Pedido:  <b>" . $customer->id  ."</b>  Nome :<b> " . 
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

}
sleep(60);
$j++;	
}










