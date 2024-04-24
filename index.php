<?php
//ini_set('max_execution_time', '0'); // for infinite time of execution 
ini_set('max_execution_time', 0); //0=NOLIMIT
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);
$start = microtime(true);
require 'vendor/autoload.php';
require 'email-sender.php';



use Automattic\WooCommerce\Client;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\SMTP;

$errors_data = array();

$inputFileName = './nacional1abril24.xls';


$reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
$reader->setReadDataOnly(true);

$readerOds = new \PhpOffice\PhpSpreadsheet\Reader\Ods();

$emailSender = new EmailSender();

$localizados_msgs = [];
$nao_localizados_msgs = [];

/* Load $inputFileName to a Spreadsheet Object  */
$spreadsheet = $reader->load($inputFileName);
$workSheet = $spreadsheet->getActiveSheet()->toArray();

define("CONSUMER_KEY", "ck_a6c8529cde896bd878ed42bb99fc757134ba16ab");
define("CONSUMER_SECRET", "cs_c5b1f053c045dbcbd895ee5ce23d93d62a7441a4");



//Verificando pedidos do woocommerce
$url = "https://www.literatour.com.br";
$consumer_key = CONSUMER_KEY;
$consumer_secret = CONSUMER_SECRET;

//TIRAR 32 DIAS
$woocommerce = new Client($url, $consumer_key, $consumer_secret);
$inicio = gmdate("Y-m-d", strtotime("-1 days")) . "T00:00:00";
$fim = gmdate("Y-m-d", strtotime("today")) . "T00:00:00";


$endpoint = "orders";
$totaldados = 1;

function addTrackingOnWP($endpoint, $customerId, $dados)
{
  global $woocommerce;
  global $errors_data;
  try {
    $woocommerce->put("$endpoint/$customerId", $dados);

  } catch (Throwable $e) {
    echo '<hr><p style="color:red;">Fatal error ao cadastrar tracking code para usuário com id: ' . $customerId . ' </p><hr>';
    $errors_data[] = array("customerId" => $customerId, "dados" => $dados);
    sleep(120);

  }

}

$j = 1;
$totaldados = 100;
$acertos = 0;

for ($i = 1; $totaldados > 0; $i++) {
  echo "Pedidos entre  $inicio e $fim" . "<br>\n";
  $parameters = [
    "status" => "processing",
    "after" => $inicioQuinzena,
    "before" => $fimQuinzena,
    "per_page" => 100,
    "page" => $i,
    "order" => "asc",

  ];


  try {
    $recentcustomers = $woocommerce->get($endpoint, $parameters);

  } catch (Throwable $e) {
    echo '<hr><p style="color:red;">Fatal error não foi possível obter os dados dos clientes</p><hr>';
    sleep(120);
    $i--;

  }


  if (isset ($recentcustomers)) {
    $totaldados = count($recentcustomers);
  } else {
    $totaldados = 0;
  }

  if ($totaldados > 0) {
    echo $totaldados . " pedidos recentes encontrados..." . "<br>";

    foreach ($recentcustomers as $customer) {

      $userFound = 0;
      $userFoundName = "";
      $rastreioLocalizado = "";
      $cliente_primeiro_nome_pedido = explode(" ", $customer->shipping->first_name)[0] ?? $customer->shipping->first_name;


      foreach ($workSheet as $row) {
        $rastreio = $row[8];
        $cep_planilha = $row[3];
        $cliente_primeiro_nome_planilha = explode(" ", $row[2])[0] ?? $row[2];

        $pedido = $customer->id;
        $data = new DateTime($customer->date_created);

        $plano = $customer->line_items[0]->name;

        $email = $customer->billing->email;

        $endereco = "CEP: " . $customer->shipping->postcode . ", " . $customer->shipping->address_1 . " " . $customer->shipping->address_2
          . ", " . $customer->shipping->number . ", " . $customer->shipping->neighborhood . ", " . $customer->shipping->city
          . ", " . $customer->shipping->state;

        if ($customer->shipping->postcode == $cep_planilha) {
          $userFoundName = $customer->billing->first_name . " " . $customer->billing->last_name;
          $rastreioLocalizado = $rastreio;
          $userFound++;
        }

        if ($userFound == 1) {


          similar_text(strtolower($cliente_primeiro_nome_pedido), strtolower($cliente_primeiro_nome_planilha), $percent);

          similar_text(strtolower($cliente_primeiro_nome_pedido), strtolower($cliente_primeiro_nome_planilha), $percent);
          if ($percent > 50) {
            //enviando email...
            $emailSender->enviarRastreioPorEmail($data,$pedido, $plano,$email,$cliente_primeiro_nome_pedido,$rastreioLocalizado,$endereco);
            //post pedido atualizado com rastreio
            //$rastreio_final = (string) $rastreioLocalizado;
            $customerId = $customer->id;
            $dados = [
              'correios_tracking_code' => "$rastreioLocalizado"
            ];

            addTrackingOnWP($endpoint,$customerId,$dados);

            array_push($localizados_msgs, "<b>$userFoundName encontrado(a)!</b>(match simples de CEP) </br>");
            $acertos++;
            unset($row);

            break;
          }

        } 

      }

      if($userFound == 0) {

        array_push($nao_localizados_msgs, "<b>Página: $j - NÃO LOCALIZADO -</b>DATA: " . $data->format('d/m/Y H:i:s') . " PEDIDO:  " . $customer->id . "  NOME : " .
          $customer->billing->first_name . " " . $customer->billing->last_name .
          " CEP " . $customer->shipping->postcode . " - UF:  " . $customer->shipping->state . " <br>");

      }


    }

  }
  sleep(25);
  $j++;
}

$time_elapsed_secs = microtime(true) - $start;

echo '<hr><p style="color:red;"> Tempo total de execução: ' . ($time_elapsed_secs / 60) . ' Minutos. </p><hr>';
$tentativas = 0;
while (count($errors_data) > 0 and $tentativas <= 3) {
  $tentativas++;
  foreach ($errors_data as $ed) {


    try {
      $woocommerce->put("$endpoint/" . $ed["customerId"], $ed["dados"]);
      if (($key = array_search($ed["customerId"], $errors_data)) !== false) {
        unset($errors_data[$key]);
      }

    } catch (Throwable $e) {
      echo '<hr><p style="color:red;">Fatal error pela ' . $tentativas . 'ª vez ao cadastrar tracking code para usuário com id: ' . $ed["customerId"] . ' </p><hr>';
      sleep(120);
      //echo $e->getMessage(); 

    }


  }


}

if (isset ($errors_data)) {
  if (count($errors_data) > 0) {

    echo 'Os códigos de rastreio dos seguintes usuários não foram cadastrados: ';

    print_r($errors_data);

    foreach ($errors_data as $edf) {

      $fp = fopen('errosrastreio.txt', 'a+');
      fwrite($fp, '');
      fwrite($fp, PHP_EOL . 'CustomerId: ' . $edf["customerId"] . PHP_EOL);
      fwrite($fp, 'Rastreio: ' . $edf["dados"]["correios_tracking_code"] . PHP_EOL);
      fwrite($fp, '');
      fwrite($fp, '############################' . PHP_EOL);
      fwrite($fp, '');
      fclose($fp);

      echo 'CustomerId: ' . $edf["customerId"] . ' - ' . 'Rastreio: ' . $edf["dados"]["correios_tracking_code"] . '<br /><br />';

    }

  }
}
echo "<br>RASTREIOS LOCALIZADOS <br>";
print_r($localizados_msgs);

echo "<br>RASTREIOS NÃO LOCALIZADOS <br>";
print_r($nao_localizados_msgs);

function calculaPercentualAcerto($localizadasCount, $naoLocalizadasCount)
{
  return $localizadasCount / ($localizadasCount + $naoLocalizadasCount) * 100;
}

echo "<br> PERCENTUAL DE ACERTO <br> de " . calculaPercentualAcerto(count($localizados_msgs), count($nao_localizados_msgs)) . " % <br>";
echo "<br>TOTAL DE ACERTOS  de " . $acertos . "<br>";