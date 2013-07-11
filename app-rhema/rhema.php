<html>
<head>
    <title>
        Insert Rhema words
    </title>
</head>
<body>
    <div id="loader">
        <img src="ajax-loader.gif" />
    </div>
        <?php
require_once('Zend/Loader.php');
//Gmail 用戶名和密碼
$user = $_POST['username'];
$pass = $_POST['password'];


//SpreadSheet key
$spreadsheetKey = "0AnGdul1OQkSfdEp1TGdzX21TZG9Kcnd4NU1QRE12aXc";
$worksheetId = "od6";

Zend_Loader::loadClass('Zend_Gdata');
Zend_Loader::loadClass('Zend_Gdata_ClientLogin');
Zend_Loader::loadClass('Zend_Gdata_Spreadsheets');
Zend_Loader::loadClass('Zend_Http_Client');

$service = Zend_Gdata_Spreadsheets::AUTH_SERVICE_NAME;
$client = Zend_Gdata_ClientLogin::getHttpClient($user,$pass,$service);
$spreadsheetService = new Zend_Gdata_Spreadsheets($client);

$query = new Zend_Gdata_Spreadsheets_CellQuery();
$query->setSpreadsheetKey($spreadsheetKey);
$query->setWorksheetId($worksheetId);
$cellFeed = $spreadsheetService->getCellFeed($query);

$id_arr = array();
$title = array();
foreach($cellFeed as $cellEntry) {
    $row = $cellEntry->cell->getRow();
    $col = $cellEntry->cell->getColumn();
    //$val = $cellEntry->cell->getText();
    //echo "$row, $col = $val <br />";
    if($row %2 == 1 && $col >2) {
        $val = explode("/",$cellEntry->cell->getText());
        $raw_id = explode(".",$val[5]);
        $id = $raw_id[0];

        array_push($id_arr,$id);
//        echo $id;
    }
    if($row % 2 == 0 && $col >2) {
        $val = $cellEntry->cell->getText();

        array_push($title,$val);
//        echo $val . "\t" . "<br />";
    }
}

//將兩個數組合併為關聯數組
$res = array_combine($id_arr,$title);

//ksort($res, SORT_NUMERIC);
//鏈接mysql資料庫
try {
    $db = new PDO("mysql:host=localhost;dbname=luke54org","luke54org","luke54org");
    $db->exec("SET NAMES 'utf8';");
} catch (PDOException $e) {
    printf("Connection error: %s", $e->getMessage());
}


foreach($res as $k => $v) {
    $db->exec("INSERT IGNORE INTO jos_rhema_articles(aid,atitle) VALUES ('$k','$v')");
}

foreach($cellFeed as $cellEntry) {
    $row = $cellEntry->cell->getRow();
    $col = $cellEntry->cell->getColumn();

    if($row % 2 == 1 && $col > 2) {

        if($row >=1 && $row <= 20) {

            $val = explode("/",$cellEntry->cell->getText());
            $raw_id = explode(".",$val[5]);
            $id = $raw_id[0];

            $iid = ($row + 1) /2;
            $db->exec("INSERT IGNORE INTO jos_rhema_result(pid, iid, aid) VALUES (1,'$iid','$id')");
//        echo $id;
        }

        if($row >=21 && $row <= 40) {

            $val = explode("/",$cellEntry->cell->getText());
            $raw_id = explode(".",$val[5]);
            $id = $raw_id[0];

            $iid = ($row - 20 + 1) /2;
            $db->exec("INSERT IGNORE INTO jos_rhema_result(pid, iid, aid) VALUES (2,'$iid','$id')");
//        echo $id;
        }

        if($row >=41 && $row <= 60) {

            $val = explode("/",$cellEntry->cell->getText());
            $raw_id = explode(".",$val[5]);
            $id = $raw_id[0];

            $iid = ($row - 40 + 1) /2;
            $db->exec("INSERT IGNORE INTO jos_rhema_result(pid, iid, aid) VALUES (3,'$iid','$id')");
//        echo $id;
        }

        if($row >=61 && $row <= 80) {

            $val = explode("/",$cellEntry->cell->getText());
            $raw_id = explode(".",$val[5]);
            $id = $raw_id[0];

            $iid = ($row - 60 + 1) /2;
            $db->exec("INSERT IGNORE INTO jos_rhema_result(pid, iid, aid) VALUES (4,'$iid','$id')");
//        echo $id;
        }
    }
}
?>
    <script type="text/javascript">
    document.getElementById("loader").innerHTML="Insert Data complete!";
    </script>
</body>
</html>





