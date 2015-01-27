<?php
ini_set('max_execution_time', 300);
require_once 'excel_reader2.php';
$xls = new Spreadsheet_Excel_Reader("ZFC.xls");
//Display the xls data with cell info
//echo "<pre/>";print_r($xls->sheets[0]['cells']);exit;

$table = "<table border=1>";
$temp_array = array();

for($i=1; $i<=count($xls->sheets[0]['cells']); $i++){
	$temp_array[] = $xls->sheets[0]['cells'][$i][1];
}
//echo "<pre>";print_r($temp_array);exit;

// Loop every element in xls
for($i=1; $i<=count($xls->sheets[0]['cells']); $i++){
	// Split number to get whole number and decimal number. Save it an array
	$num_split[] = explode('.', $xls->sheets[0]['cells'][$i][1]);
	// New array starts from index 0 but we assigned i = 1. So make it i-1
	$j = $i-1;
	// Get 1st digit from decimal part and push to an array
    $decimal[] = substr($num_split[$j][1],0,1);
    // Get whole number and push to array
    $whole_number[] = $num_split[$j][0];
    // To check first iteration or not
	if($i > 1){
		$whole_number_count = array_count_values($whole_number);
		$current_whole_num = $num_split[$j][0];
		$number = $num_split[$j][0].'.5';
		$result_key = get_nearest_number($temp_array,$number);	
		if($whole_number_count[$current_whole_num] == 1){
			$temp = $xls->sheets[0]['cells'][$i][1];
			$res = $xls->sheets[0]['cells'][$i][2];					  
		}else if($result_key == $i){				
			$temp = $xls->sheets[0]['cells'][$i][1];
			$res = $xls->sheets[0]['cells'][$i][2];	
		}else{
			$temp = null;
			$res = null;	
		}
	}else{ 
		$temp = $xls->sheets[0]['cells'][$i][1];
		$res = $xls->sheets[0]['cells'][$i][2];		
	}
	if(!empty($temp) && !empty($res)){
		$table.= "<tr>
						<td>
							<p><span>$temp</span><o:p></o:p></p>
						</td>
						<td colspan=\"2\">
							<p><span>$res</span><o:p></o:p></p>
						</td>
					  </tr>";
	}
}

$table.="</table>";
$file="ZFC(Mid_Points).xls";
header("Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header("Content-Disposition: attachment; filename=$file");
echo $table;

function get_nearest_number($arr = array(), $search = null){
	$closest = null;
	   foreach($arr as $key => $item) {
	      if($closest == null || abs($search - $closest) > abs($item - $search)) {
	         $closest = $item;
			$result_key = $key;
	      }
	   }
	return $result_key+1 ;
}
?>
