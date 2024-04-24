<?php

require 'vendor/autoload.php';
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PHPMailer\PHPMailer\SMTP;

class EmailSender
{

    public function enviarRastreioPorEmail($data, $pedido, $plano, $email, $cliente, $rastreio, $endereco)
    {

        // Instantiation and passing `true` enables exceptions
        $mail = new PHPMailer(true);

        try {
            //Server settings
            $mail->SMTPDebug = 3;
            $mail->CharSet = 'UTF-8';
            $mail->isSMTP();                                            // Send using SMTP
            $mail->Host = 'smtp.elasticemail.com';                    // Set the SMTP server to send through
            $mail->SMTPAuth = true;
            $mail->SMTPSecure = 'tls';
            $mail->Username = 'suporte@literatour.com.br';                     // SMTP username
            $mail->Password = 'A1C58D1C35B5D28469407A95024B54F3F95B';                              // SMTP password
            $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;         // Enable TLS encryption; `PHPMailer::ENCRYPTION_SMTPS` encouraged
            $mail->Port = 2525;                                    // TCP port to connect to, use 465 for `PHPMailer::ENCRYPTION_SMTPS` above

            //Recipients
            $mail->setFrom('assessoria@literatour.com.br', 'Literatour');
            $mail->addAddress($email, $cliente);     // Add a recipient
            $mail->addBCC('77luanlima@gmail.com', 'Luan Lima');
            // Content
            $mail->isHTML(true);                                  // Set email format to HTML
            $mail->Subject = 'Literatour - Sua caixinha já foi enviada  pelos Correios';
            $mail->Body = "
              <html>
              <head><title>Caixinha a caminhoo :) </title></head>
              <body>
                  <div style='background-color: #ff8000; color:snow; font-family: Arial, Helvetica, sans-serif'>
                      <h1>Kit enviado!! Uhull!  </h1>
                  </div>
                  <div id='content'>
              <p>Olá,$cliente! Sua caixinha já foi enviada pelos Correios :)</p>
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
}
