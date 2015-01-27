<?php 
require_once 'excel_reader2.php';
$xls = new Spreadsheet_Excel_Reader("FC.xls");
//Display the xls data with cell info
//echo "<pre/>";print_r($xls->sheets);exit;

$table = "<table border=1>";
$temparature = $decimal = $num_split = array();
// Loop every element in xls
for($i=1; $i<=count($xls->sheets[0]['cells']); $i++){
	// Split number to get whole number and decimal number. Save it an array
	$num_split[] = explode('.', $xls->sheets[0]['cells'][$i][1]);
	// New array starts from index 0 but we assigned i = 1. So make it i-1
	$j = $i-1;
	// Get 1st digit from decimal part and push to an array
    $decimal[] = substr($num_split[$j][1],0,1);
    // To check first iteration or not
	if($i > 1){
		// Eg: 2.00452
		// Check current whole number(Eg : 2) equal to previour and current decimal number(0) equal to previous or not
		// If equal, do not insert dat data to xls. Else insert it to xls
		if($num_split[$j][0] == $num_split[$j-1][0] && $decimal[$j] == $decimal[$j-1]){         
		}else{ 	          	
			$temp = $xls->sheets[0]['cells'][$i][1];
			$res = $xls->sheets[0]['cells'][$i][2];					
			$table.= "<tr>
						<td>
							<p><span>$temp</span><o:p></o:p></p>
						</td>
						<td colspan=\"2\">
							<p><span>$res</span><o:p></o:p></p>
						</td>
					  </tr>";
		}
	}else{ 
		$temp = $xls->sheets[0]['cells'][$i][1];
		$res = $xls->sheets[0]['cells'][$i][2];			
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
$file="FC(Decimal).xls";
header("Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header("Content-Disposition: attachment; filename=$file");
echo $table;
?>
