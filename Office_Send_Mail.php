<?php
    
    namespace Office_Send_Mail;
    /*
     @Author Michael McIntee
     @Version 0.1
     @Date 14/10/2020
     */
    class Office_Send_Mail {
        
        public $scope = 'Mail.Send%20offline_access%20SMTP.Send';
        public $tenantId;
        public $clientId;
        private $clientSecret;
        public $code;
        public $refreshToken;
        public $accessToken;
        private $redirect_uri;
        private $base_uri;
        private $response;
        private $filename = 'refresh.ini';
        
        /*
         The constructor accepts the tenant ID of the azure application, the client ID and the client secret. It negotiates to recieve the authorization code once the app developer logs in. After obtaining the code, the constructor then negotiates to either refresh the client token, or obtains the token without a refresh, and this allows access to the mail functionality of 365.
         */
        function __construct($tenantId, $clientId, $clientSecret) {
            
            $this->tenantId = $tenantId;
            $this->clientId = $clientId;
            $this->clientSecret = $clientSecret;
            $this->redirect_uri = (isset($_SERVER['HTTPS'])) ? "https://" : "http://" . $_SERVER['HTTP_HOST'] . $_SERVER['PHP_SELF'];
            $this->base_uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/";
            
            if($this->check_refresh_date()) {
                
                $this->get_refreshed_token();
                
            } else {
                
                if(!isset($_GET['code'])) {
                    
                    $req = curl_init();
                    curl_setopt($req, CURLOPT_URL, $this->base_uri. "authorize?client_id=$this->clientId&response_type=code&redirect_uri=$this->redirect_uri&response_mode=query&scope=$this->scope");
                    $response = curl_exec($req);
                    $this->code = $_GET['code'];
                    
                } else {
                    
                    $this->get_access_token($_GET['code']);
                    
                }
            }
        }
        
        /*
         Using the obtained access code, this function will negotiate with office 365 to obtain the access token and the refresh token. These tokens will be used to either refresh the period of access, or utilise the token to send an email using the appropriate function.
         */
        function get_access_token($code) {
            
            $body_data =  "client_id=$this->clientId&scope=$this->scope&code=$code&redirect_uri=$this->redirect_uri&grant_type=authorization_code&client_secret=$this->clientSecret";
            $locate = "{$this->base_uri}token";
            
            $req = curl_init($locate);
            curl_setopt($req, CURLOPT_RETURNTRANSFER, true);
            curl_setopt($req, CURLOPT_CUSTOMREQUEST,"POST");
            curl_setopt($req, CURLOPT_POSTFIELDS, $body_data);
            curl_setopt($req, CURLOPT_HTTPHEADER, array('Content-Type: application/x-www-form-urlencoded', 'Content-Length: ' .strlen($body_data)));
            $response = curl_exec($req);
            
            $response_data = json_decode($response, true);
            $this->refreshToken = $response_data['refresh_token'];
            $this->accessToken = $response_data['access_token'];
            
            if(!isset($response_data['refresh_token'])) {
                
                die("Refresh token not created!");
                
            } else {
                
                if(!$this->check_refresh_ini()) {
                    
                    $this->update_create_refresh_ini();
                    
                }
            }
        }
        
        /*
         This function gets a new access token by using the refresh token to obtain it. It writes the new refresh token to a file, and proceeds to obtain the access file. This is called to obtain an access token for use in sending emails to ensure the access token is obtained successfully.
         */
        function get_refreshed_token() {
            
            if(!$this->check_refresh_exists()) {
                
                $this->update_create_refresh_ini();
                
            }
            
            if($this->check_refresh_date()) {
                
                $this->load_refresh_token();
                $this->update_create_refresh_ini();
                
            } else {
                
                $this->update_create_refresh_ini();
                
            }
            
            $body_data =  "client_id=$this->clientId&scope=$this->scope&refresh_token=$this->refreshToken&grant_type=refresh_token&client_secret=$this->clientSecret";
            
            $locate = "{$this->base_uri}token";
            $req = curl_init($locate);
            curl_setopt($req, CURLOPT_RETURNTRANSFER, true);
            curl_setopt($req, CURLOPT_CUSTOMREQUEST,"POST");
            curl_setopt($req, CURLOPT_POSTFIELDS, $body_data);
            curl_setopt($req, CURLOPT_HTTPHEADER, array('Content-Type: application/x-www-form-urlencoded', 'Content-Length: ' .strlen($body_data)));
            
            $response = curl_exec($req);
            $response_data = json_decode($response, true);
            $this->accessToken = $response_data['access_token'];
            
        }
        
        /*
         This function checks for the refresh_ini date, which means a new refresh token needs to be produced and stored. If the refresh token is still in date, then no new refresh token needs to be produced, and only the date needs to be updated.
         */
        function check_refresh_date() {
            
            if($this->check_refresh_exists()) {
                
                $ini_data = parse_ini_file($this->filename);
                $date_create = $ini_data['date_created'];
                
                if(date('Y-m-d') < date('Y-m-d', strtotime($date_create. ' + 14 days')))  {
                    return true;
                } else {
                    return false;
                }
                
            } else {
                return false;
            }
        }
        
        /*
         This function checks if the refresh_ini exists, and can be used to test it exists before the loading of the file is attempted.
         */
        function check_refresh_exists() {
            
            return file_exists($this->filename);
            
        }
        
        /*
         The update create function will create or update the ini file with a new refresh token, following an email being sent or in the event that reauthentication was required.
         */
        function update_create_refresh_ini() {
            
            $refresh_file = fopen($this->filename, "w") or die("Unable to open file!");
            $date_created = date('Y-m-d');
            $ini_data = "refresh_token={$this->refreshToken}\ndate_created={$date_created}";
            fwrite($refresh_file, $ini_data);
            fclose($refresh_file);
            
        }
        
        /**
         This function loads the value of the access token from the file as long as it's still in date. This allows the same refresh token to be written to file in the event that a send mail function is called, as the tokens expiration period resets if it's used again to refresh.
         */
        function load_refresh_token() {
            
            $ini_data = parse_ini_file($this->filename);
            $this->refreshToken=$ini_data['refresh_token'];
            
        }
        
        /*
         Once the relevant access token retrieval has taken place, this method will make a request to office to send an email to the email provided in the to field. It will set the email subject to the value of the subject field, and it will set the email body to the value of the html body.*/
        function send_mail($to, $subject, $html_body) {
            
            $this->get_refreshed_token();
            $body_data = '{"message":{"subject":"'.$subject.'","body":{"contentType":"HTML","content":"'.$html_body.'"},"toRecipients":[{"emailAddress":{"address":"'.$to.'"}}]}}';
            
            $locate = "https://graph.microsoft.com/v1.0/me/sendMail";
            
            $req = curl_init($locate);
            curl_setopt($req, CURLOPT_RETURNTRANSFER, true);
            curl_setopt($req, CURLOPT_CUSTOMREQUEST,"POST");
            curl_setopt($req, CURLOPT_POSTFIELDS, $body_data);
            curl_setopt($req, CURLOPT_HTTPHEADER, array("Authorization: {$this->accessToken}",'Content-Type: application/json', 'Content-Length: ' .strlen($body_data)));
            
            $response = curl_exec($req);
            curl_close($req);
            
        }
    }
    
    $my_class = new Office_Send_Mail('tenant_id','client_id', 'client_secret');
    
    $my_class->send_mail('email','subject','<b>html body</b>');
    ?>
