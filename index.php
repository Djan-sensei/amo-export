<?php

$start = microtime(true);

if(!isset($_GET['from']) and !isset($_GET['to'])) {
	echo 'Ошибка';
	exit;
}

// Задачи по дате
function tasks_curl() {
	
	global $_GET;
	
	$date_from = strtotime($_GET['from'].' 00:00:00');
	$date_to = strtotime($_GET['to'].' 23:59:59');
	
	$link = 'https://gklad.amocrm.ru/api/v2/tasks?login=логин&api_key=ключ&type=lead&filter[date_create][from]='.$date_from.'&filter[date_create][to]='.$date_to.'&filter[pipe][ID_воронки][]=ID_этапов';
	$curl = curl_init();
	curl_setopt($curl, CURLOPT_RETURNTRANSFER,true);
	curl_setopt($curl, CURLOPT_USERAGENT, "amoCRM-API-client/1.0");
	curl_setopt($curl, CURLOPT_HTTPHEADER, "Accept: application/json");
	curl_setopt($curl, CURLOPT_URL, $link);
	curl_setopt($curl, CURLOPT_HEADER,false);
	$out = curl_exec($curl);
	curl_close($curl);
	$result = json_decode($out,TRUE);
	
	// Если ошибка
	if(isset($result['response']['error_code'])) {
		mail('*****', "AmoCRM Tasks export (export_tasks - tasks_curl): ".$result['response']['error'], print_r($result, true));
		exit; // останавливаем весь процесс
	}
	
	return $result;
}

// Инфо по аккаунту
function account() {
	
	$link = 'https://gklad.amocrm.ru/api/v2/account?login=логин&api_key=ключ&with=users';
	$curl = curl_init();
	curl_setopt($curl, CURLOPT_RETURNTRANSFER,true);
	curl_setopt($curl, CURLOPT_USERAGENT, "amoCRM-API-client/1.0");
	curl_setopt($curl, CURLOPT_HTTPHEADER, "Accept: application/json");
	curl_setopt($curl, CURLOPT_URL, $link);
	curl_setopt($curl, CURLOPT_HEADER,false);
	$out = curl_exec($curl);
	curl_close($curl);
	$result = json_decode($out,TRUE);
	
	// Если ошибка
	if(isset($result['response']['error_code'])) {
		mail('*****', "AmoCRM Tasks export (export_tasks - account): ".$result['response']['error'], print_r($result, true));
		exit; // останавливаем весь процесс
	}
	
	return $result;
}

	$array = array();

	$task = tasks_curl();
	$mass = $task['_embedded']['items'];
	
	$acc = account();
	$account = $acc['_embedded']['users'];
	
	// Пользователи
	foreach ($mass as $key => $value) {
		
		$sozd = ''; $otv = '';
		foreach ($account as $k => $val) {
			if($val['id'] == $value['responsible_user_id']) {
				$otv = $val['name'];
			}
		}
		
		if($value['is_completed'] == 1) { $fin = 'завершена'; }
		else { $fin = 'НЕ завершена'; }
		
		$array[$key] = array(
			'otv' => $otv,
			'text' => $value['text'],
			'end_date' => date('j.m.Y H:i:s', $value['complete_till_at']),
			'completed' => $fin,
			'element_id' => $value['element_id'],
			'created_at' => date('j.m.Y H:i:s', $value['created_at']),
			'updated_at' => date('j.m.Y H:i:s', $value['updated_at'])
		);
		
	}

/****************** EXEL ******************/

	require_once 'PHPExcel-1.8/Classes/PHPExcel.php';

	$document = new \PHPExcel();		 
	// Выбираем первый лист в документе
	$sheet = $document->setActiveSheetIndex(0);
	// Начальная координата x
	$columnPosition = 0;
	// Начальная координата y
	$startLine = 1;
	// Массив с названиями столбцов
	$columns = ['Ответственный', 'Информация', 'До какого выполнить', 'Статус задачи', 'ID сделки', 'Дата создания', 'Дата последнего обновления'];
	// Указатель на первый столбец
	$currentColumn = $columnPosition;		 
	
	// Формируем шапку
	foreach ($columns as $column) {
		$sheet->setCellValueByColumnAndRow($currentColumn, $startLine, $column);
		// Смещаемся вправо
		$currentColumn++;
	}
	 
	// Формируем список
	foreach ($array as $key=>$catItem) {
		// Перекидываем указатель на следующую строку
		$startLine++;
		// Указатель на первый столбец
		$currentColumn = $columnPosition;
		// Ставляем информацию об имени и цвете
		foreach ($catItem as $value) {				
			$sheet->setCellValueByColumnAndRow($currentColumn, $startLine, $value);
			// Смещаемся вправо
			$currentColumn++;
		}
	}
	 
	$objWriter = \PHPExcel_IOFactory::createWriter($document, 'Excel5');
	
	$file_name = date('dmY_His').".xls";
	
	$objWriter->save('export_all/'.$file_name);

	echo '<script>
		window.location.href="/export_all/'.$file_name.'";
	</script>';
	
/****************** [end]EXEL ******************/

?>