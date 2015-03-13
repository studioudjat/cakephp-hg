<?php

// Excel出力用ライブラリ
App::import('Vendor', 'PHPExcel', array('file' => 'phpexcel' . DS . 'PHPExcel.php'));
App::import('Vendor', 'PHPExcel_IOFactory', array('file'=>'phpexcel'. DS .'PHPExcel'. DS .'IOFactory.php'));
App::import('Vendor', 'PHPExcel_Cell_AdvancedValueBinder', array('file'=>'phpexcel'. DS .'PHPExcel'. DS .'Cell'. DS .'AdvancedValueBinder.php'));

// Excel95用ライブラリ
App::import('Vendor', 'PHPExcel_Writer_Excel5', array('file'=>'phpexcel' . DS . 'PHPExcel' . DS . 'Writer' . DS . 'Excel5.php'));
App::import('Vendor', 'PHPExcel_Reader_Excel5', array('file'=>'phpexcel' . DS . 'PHPExcel' . DS . 'Reader' . DS . 'Excel5.php'));
  
class ImportProductsController extends AppController {

  public $name = 'ImportProducts';
  public $uses = array('Collection', 'CategoriesProduct', 'Product', 'Datamap', 'ImportProduct', 'DeldateDatamap');
  public $components = array('Session');
  
  public function index() {
  
    if ($this->request->is('post')) {
      
      $this->ImportProduct->set($this->request->data);
      if ($this->ImportProduct->validates()) {
      
        // 台帳ファイルの定義
        $inventory = realpath(TMP).'/phpexcel/' . $this->request->data['ImportProduct']['filename'];
      
        // データマッピング情報を取得
        $dataMap = $this->_readDataMap();
      
        // オブジェクトの作成
        $objPHPExcel = new PHPExcel();
        $objReader = new PHPExcel_Reader_Excel2007();
  	
  	    // 台帳ファイルの読み込み
  	    $objPHPExcel = $objReader->load($inventory);
  	
  	    // 台帳ファイルからデータを取得
  	    $data = $this->_getData($objPHPExcel, $dataMap, $this->request->data['ImportProduct']);
  	    
  	    // 形式チェックにエラーがある場合は、処理を終了
  	    if ($data['error']) {
  	      $this->Session->setFlash($data['error']);
  	    }
  	    else {
  	    
  	      if (isset($this->request->data['import'])) {
  
            // 取得したデータをDBに挿入
            if ($this->Product->saveAll($data['Product'])) {
              $this->Session->setFlash('Import successful');
            }
            else {
              //debug($this->Product->validationErrors);
              $this->Session->setFlash('Failed to save to database');
            }
            
          }
          else {
            $this->set('data', $data);
          }
        }
      }
      else {
        $this->Session->setFlash('Import failed');
      }
    }
  
    $this->set('collections', $this->Collection->find('all'));
    $this->set('categories', $this->CategoriesProduct->find('all'));
    
    // アップロードフォルダーにあるファイル情報を取得
    $this->set('files', scandir(realpath(TMP).'/phpexcel/'));
    
  }
  
  public function _readDataMap () {
  
    $results = $this->Datamap->find('all', array(
      'conditions' => array('status' => 'Enable'),
      'order' => array('style_no' => 'asc')));
    
    for ($i = 0; $i < count($results); $i++) {
    
      $dataMap[$i] = array(
        'style_no' => $results[$i]['Datamap']['style_no'],
        'style_name' => $results[$i]['Datamap']['style_name'],
        'del_date' => $results[$i]['Datamap']['del_date'],
        'price' => $results[$i]['Datamap']['price'],
        'material' => $results[$i]['Datamap']['material'],
        'size' => $results[$i]['Datamap']['size'],
        'color' => $results[$i]['Datamap']['color']
      );
      
    }
          
    return $dataMap;
    
  }
  
  public function _getData ($objPHPExcel, $dataMap, $params) {
  
    $row = 0;
    $data['error'] = '';
    
    for ($i = 0; $i < $objPHPExcel->getSheetCount(); $i++) {
  
      // 0番目のシートをアクティブにする（シートは左から順に、0、1，2・・・）
      $objPHPExcel->setActiveSheetIndex($i);
      // アクティブにしたシートの情報を取得
      $objSheet = $objPHPExcel->getActiveSheet();
      
      // シートにあるユニークの商品番号を取得
      $style_nos = array();
      $products = array();
      for ($j = 0; $j < count($dataMap); $j++) {
      
        // スタイル番号を取得し、最後の文字列を削除
        $style_no = substr($objSheet->getCell($dataMap[$j]["style_no"])->getValue(), 0, -1);
        
        if (!empty($style_no)) {
        
          // スタイル番号の形式をチェック。マッチしない場合は処理を終了
          if (!preg_match('/^\w{8}$/', $style_no)) {
            $data['error'] = 'Invalid style number ' . $style_no . ' in sheet ' . $objSheet->getTitle() .
              ' cell ' . $dataMap[$j]["style_no"] . '. Please review data mapping';
            break;
          }
          
          // ユニークな商品だけ配列に追加
          if (array_search($style_no, $style_nos) === FALSE) {
            array_push($style_nos, $style_no);
            array_push($products, array('index' => $j, 'style_no' => $style_no));
          }
            
        }
      
      }
      
      // データが空の場合は、次のシートを読み込む
      if (empty($products)) continue;
      
      // データ形式にエラーがある場合はループ処理を終了
      if ($data['error']) break;
      
      // シートにある商品を全部取得
      for ($j = 0; $j < count($products); $j++) {
      
        // ユニークな商品があったインデックスを取得
        $index = $products[$j]['index'];
        
        // 発売日の形式をチェック。マッチしない場合は処理を終了
        $del_date = $objSheet->getCell($dataMap[$index]["del_date"])->getValue();
        if(!preg_match('/^[Beginning|Mid|End]/i', $del_date)) {
          $data['error'] = 'Invalid delivery date ' . $del_date . ' in sheet ' . $objSheet->getTitle() .
            ' cell ' . $dataMap[$index]["del_date"] . '.  Please review data mapping';
          break;
        }
        // 一ヶ月半時期をずらす
        $del_date = $this->_convertDelDate($del_date);
        
        // 価格の形式をチェック。マッチしない場合は処理を終了
        $yen = $objSheet->getCell($dataMap[$index]["price"])->getValue();
        if (!preg_match('/^\d+$/', $yen)) {
          $data['error'] = 'Invalid price ' . $yen . ' in sheet ' . $objSheet->getTitle() .
          ' cell ' . $dataMap[$index]["price"] . '.  Please review data mapping';
          break;
        }
        // (日本上代 x 40%) x 1.5 / 100 = US Wholesale Price
        //$dollar = $yen * 0.4 * 1.5 / 100;
        $dollar = round($yen * 0.4 * 1.5 / 100);
        $price = sprintf("%.2f", $dollar);
    	
    	$size = $this->_mergeCells($objSheet, $dataMap[$index]["size"], ',');
    	// 余計の文字列を取り除く 例：(S/SIZE) -> S
    	$size = preg_replace(array('/\(/', '/\/\w+\)/'), '', $size);
    	// サイズではない文字列は削除
    	$size = preg_replace('/[^XXS|XS|S|M|L|S|\d+],/i', '', $size);
    	$size = preg_replace('/[^XXS|XS|S|M|L|S|\d+|,]/i', '', $size);
    	$size = preg_replace('/,$/', '', $size);
    	if (empty($size)) $size = 'ONE SIZE';
    	
    	// 素材情報を取得し、翻訳
    	$material = $this->_translateJP($this->_mergeCells($objSheet, $dataMap[$index]["material"], ' '));
    	//$material = $this->_mergeCells($objSheet, $dataMap[$index]["material"], ' ');
    	
    	// 各色の前に、アルファベット文字列を追加　例：White -> A: White
    	$color = $this->_mergeCells($objSheet, $dataMap[$index]["color"], ',');
    	$colors = explode(',', $color);
    	for ($k = 0; $k < count($colors); $k++) {
    	  $colors[$k] = chr(65+$k) . ": " . $colors[$k];
    	}
    	$color = implode(',', $colors);
    	
    	$data['Product'][$row] = array (
    	  // 最後の一文字を削除
    	  'style_no' => $products[$j]['style_no'],
    	  'style_name' => $this->_translateJP($this->_mergeCells($objSheet, $dataMap[$index]["style_name"], ' ')),
    	  //'style_name' => $this->_mergeCells($objSheet, $dataMap[$index]["style_name"], ' '),
    	  'del_date' => "{$del_date[0]} {$del_date[1]}", 
    	  'price' => $price,
    	  'material' => $material,
    	  'size' => $size,
    	  'color' => $color,
    	  'status' => 'Enable',
    	  'sex' => $params['sex'],
    	  'collection_id' => $params['collection_id'],
    	  'category_id' => $params['category_id']
    	  );
    	  
    	$row++;

  	  }
    }
    
    return $data;
    
  }
  
  public function _translateJP ($jp_val) {
  
  	$appid = 'etb8CdZV42kvrQTkWNlPdlfpTIjHvTBPPhC1swOCSt4';
  	$from = 'ja';
  	$to = 'en';
  	$text = mb_convert_kana($jp_val,"K");
  	
  	$ch = curl_init('https://api.datamarket.azure.com/Bing/MicrosoftTranslator/v1/Translate?Text=%27'.urlencode($text).'%27&From=%27'.$from.'%27&To=%27'.$to.'%27');
  	curl_setopt($ch, CURLOPT_USERPWD, $appid.':'.$appid);
  	curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
  	
  	$result = curl_exec($ch);
  	$result = explode('<d:Text m:type="Edm.String">', $result);
  	$result = explode('</d:Text>', $result[1]);
  	$result = $result[0];
  
  	return $result;
  }
  
  public function _mergeCells ($objSheet, $val, $separator) {
  
    // カンマ区切りのセル情報を取得し、配列に挿入
  	$vals = explode(',',$val);
  	
  	$results = "";
  	
  	// ループで取得した複数のセル情報をカンマ区切りで文字列として結合
  	for ($i = 0; $i < count($vals); $i++) {
  	
  	  // セル情報が空の場合はループ処理を終了
  	  if ($objSheet->getCell($vals[$i])->getValue() == "") continue;
  	  
  	  $results = $results . $objSheet->getCell($vals[$i])->getValue() . $separator;
  	  
  	}
  	
  	// 最後の文字列を削除
  	$results = substr($results, 0, -1);
  	
  	return $results;
  }
  
  public function _convertDelDate ($val) {
    
    $vals = explode(' ', $val);
    
    // 時期を取得 例：Beginning => 1
    $conditions = array('term like' => $vals[0] . '%');
    $day = $this->DeldateDatamap->find('first', array('fields' => 'num_value', 'conditions' => $conditions));
    
    // 数字の月を取得 例：January => 1
    $conditions = array('term like' => $vals[1] . '%');
    $month = $this->DeldateDatamap->find('first', array('fields' => 'num_value', 'conditions' => $conditions));
    
    $year = date('Y');
    
    // 日本の出荷日を取得
    $jp_del_date = "{$year}-{$month['DeldateDatamap']['num_value']}-{$day['DeldateDatamap']['num_value']}";
    
    // 一ヶ月半後の日付を取得
    $day = date('d', strtotime("{$jp_del_date} +45 days"));
    $month = date('m', strtotime("{$jp_del_date} +45 days"));
    
    if ($day <= 10) {
      $term = "Beginning";
    }
    elseif ($day > 10 and $day < 20) {
      $term = "Mid";
    }
    else {
      $term = "End";
    }
    
    // 月を取得 例： 1 => January
    $conditions = array('num_value' => $month);
    $month = $this->DeldateDatamap->find('first', array('fields' => 'term', 'conditions' => $conditions));
    
    return array($term, $month['DeldateDatamap']['term']);
  }
    
}
?>