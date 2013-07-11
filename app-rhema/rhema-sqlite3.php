<?php
require_once('Zend/Loader.php');
//Gmail 用戶名和密碼
$user = "hitobias@gmail.com";
$pass = "ZionofChrist";
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
//創建資料庫

$db = new SQLite3("rhema.db");
if(!$db) die($error);

//創建表
$sql_create_table_people='CREATE TABLE IF NOT EXISTS people (id  INTEGER PRIMARY KEY, pid INTEGER NOT NULL, pname TEXT NOT NULL, UNIQUE(pid,pname));';
$sql_create_table_issue='CREATE TABLE IF NOT EXISTS issue (id  INTEGER PRIMARY KEY, iid INTEGER NOT NULL, iname TEXT NOT NULL, UNIQUE(iid,iname));';
$sql_create_table_articles='CREATE TABLE IF NOT EXISTS articles (id  INTEGER PRIMARY KEY, aid INTEGER NOT NULL, atitle TEXT NOT NULL, UNIQUE(aid,atitle));';
$sql_create_table_result='CREATE TABLE IF NOT EXISTS result (id  INTEGER PRIMARY KEY, pid INTEGER NOT NULL, iid INTEGER NOT NULL, ritemid INTEGER NOT NULL);';
//插入people數據
$sql_insert_into_people_worker='INSERT INTO people(pid, pname) VALUES (1, "上班族");';
$sql_insert_into_people_student='INSERT INTO people(pid, pname) VALUES (2, "學生");';
$sql_insert_into_people_love='INSERT INTO people(pid, pname) VALUES (3, "感情 & 婚姻");';
$sql_insert_into_people_family='INSERT INTO people(pid, pname) VALUES (4, "家庭 & 親子");';

//插入情況數據
$sql_insert_into_issue_press='INSERT INTO issue(iid, iname) VALUES (1, "被壓扁了");';
$sql_insert_into_issue_money='INSERT INTO issue(iid, iname) VALUES (2, "錢不夠用");';
$sql_insert_into_issue_lonely='INSERT INTO issue(iid, iname) VALUES (3, "總是一個人");';
$sql_insert_into_issue_staff='INSERT INTO issue(iid, iname) VALUES (4, "諸事不順");';
$sql_insert_into_issue_worry='INSERT INTO issue(iid, iname) VALUES (5, "想太多");';
$sql_insert_into_issue_belive='INSERT INTO issue(iid, iname) VALUES (6, "還能信任誰？");';
$sql_insert_into_issue_no='INSERT INTO issue(iid, iname) VALUES (7, "我也不想這樣");';
$sql_insert_into_issue_boring='INSERT INTO issue(iid, iname) VALUES (8, "好無聊喔！");';
$sql_insert_into_issue_unhappy='INSERT INTO issue(iid, iname) VALUES (9, "不開心");';
$sql_insert_into_issue_guide='INSERT INTO issue(iid, iname) VALUES (10,"需要方向");';

//執行創建表
$db->exec($sql_create_table_people);
$db->exec($sql_create_table_issue);
$db->exec($sql_create_table_articles);
$db->exec($sql_create_table_result);
//插入people表數據
$db->exec($sql_insert_into_people_worker);
$db->exec($sql_insert_into_people_student);
$db->exec($sql_insert_into_people_love);
$db->exec($sql_insert_into_people_family);

$db->exec($sql_insert_into_issue_press);
$db->exec($sql_insert_into_issue_money);
$db->exec($sql_insert_into_issue_lonely);
$db->exec($sql_insert_into_issue_staff);
$db->exec($sql_insert_into_issue_worry);
$db->exec($sql_insert_into_issue_belive);
$db->exec($sql_insert_into_issue_no);
$db->exec($sql_insert_into_issue_boring);
$db->exec($sql_insert_into_issue_unhappy);
$db->exec($sql_insert_into_issue_guide);

foreach($res as $k => $v) {
    $db->beginTransaction();
    $db->exec("INSERT INTO articles(aid,atitle) VALUES ('$k','$v')");
    $db->commit();
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
            $db->beginTransaction();
            $db->exec("INSERT INTO result(pid, iid, ritemid) VALUES (1,'$iid','$id')");
            $db->commit();
//        echo $id;
        }

        if($row >=21 && $row <= 40) {

            $val = explode("/",$cellEntry->cell->getText());
            $raw_id = explode(".",$val[5]);
            $id = $raw_id[0];

            $iid = ($row - 20 + 1) /2;
            $db->exec("INSERT INTO result(pid, iid, ritemid) VALUES (2,'$iid','$id')");
//        echo $id;
        }

        if($row >=41 && $row <= 60) {

            $val = explode("/",$cellEntry->cell->getText());
            $raw_id = explode(".",$val[5]);
            $id = $raw_id[0];

            $iid = ($row - 40 + 1) /2;
            $db->exec("INSERT INTO result(pid, iid, ritemid) VALUES (3,'$iid','$id')");
//        echo $id;
        }

        if($row >=61 && $row <= 80) {

            $val = explode("/",$cellEntry->cell->getText());
            $raw_id = explode(".",$val[5]);
            $id = $raw_id[0];

            $iid = ($row - 60 + 1) /2;
            $db->exec("INSERT INTO result(pid, iid, ritemid) VALUES (4,'$iid','$id')");
//        echo $id;
        }
    }
}

?>



