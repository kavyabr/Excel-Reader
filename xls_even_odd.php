<?php 
require_once 'excel_reader2.php';
$xls = new Spreadsheet_Excel_Reader("BAO.20_ZFC.xls");
//Display the xls data with cell info
//echo "<pre/>";print_r($xls->sheets);exit;

$temparature = array();
$table = "<table border=1>";
// Loop every element in xls
for($i=1;$i<=count($xls->sheets[0]['cells']);$i++)
{
	// To check number is even or not
	// If u want only even series, uncomment the below if block
	// If u wnat only odd series, replace 0 by 1 and uncomment the below if block
	
	//if(($xls->sheets[0]['cells'][$i][1] % 2) != 0) {
		// To check first reading or not. First time it will be empty and goes to else part & insert the data into xls
		if(!empty($temparature))
		{ 
			// Convert 2.0345 to 2 and check 2 is already exist in xls or not.
			// If exist do not insert 2 series again
			$check_reading = check_reading_exist($temparature,floor($xls->sheets[0]['cells'][$i][1]));
			// If number is not exist, then insert current temparature to array and display in xls
			if($check_reading == 0){
				$temparature[] = floor($xls->sheets[0]['cells'][$i][1]);
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
			$temparature[] = floor($xls->sheets[0]['cells'][$i][1]);
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
	//}
}
$table.="</table>";
$file="BAO.20_FC(EvenOdd).xls";
header("Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header("Content-Disposition: attachment; filename=$file");
echo $table;

// To check reading exist in xls or not
function check_reading_exist($temparature,$current_temp)
{
	if(in_array($current_temp,$temparature)){
			return 1;
	}else
			return 0;	
}

?>
