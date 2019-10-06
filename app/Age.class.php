<?php

	namespace App;

	use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
	use PhpOffice\PhpSpreadsheet\IOFactory;

	class Age {

		private $apiKeySim;
		private $file;
		private $ageSheet;
		private $row;
		private $sheetName;
		private $writer;
		private $reader;
		private $resultsFile;
		private $spreadsheet;
		private $alphabet;

		public function __construct($apiKey = "CmAHq1F6wW7UQ", $file = 'sample.csv', $resultsFile = "results.xlsx", $sheetName = false) {
			if(!$sheetName) {
				throw new \Exception("no sheetName");
			}
			$this->sheetName = $sheetName;
			$this->alphabet = range( 'A', 'Z' );
			$this->resultsFile = $resultsFile;
			$this->row = 1;
			$this->apiKeySim  = $apiKey;
			$this->file = $file;
			$this->ageSheet = new Worksheet($this->spreadsheet, $this->sheetName);

			try {
				$this->reader = IOFactory::createReader("Xlsx");
				$this->spreadsheet = $this->reader->load($this->resultsFile);
				if($this->spreadsheet->sheetNameExists($this->sheetName)) {
					$sheetIndex = $this->spreadsheet->getIndex(
						$this->spreadsheet->getSheetByName($this->sheetName)
					);
					$this->spreadsheet->removeSheetByIndex($sheetIndex);
					$this->spreadsheet->addSheet($this->ageSheet);
				} else {
					$this->spreadsheet->addSheet($this->ageSheet);
				}
				$this->writer = IOFactory::createWriter($this->spreadsheet, "Xlsx");
			} catch(\Exception $e) {
				throw new \Exception($e);
			}

			$this->readCsv();
		}

		public function readCsv() {
			if (($handle = fopen($this->file, "r")) !== false) {
				$currentId = 1;
				while (($data = fgetcsv($handle, 1000, ",")) !== false) {
					$num = count($data);

					$arr = [];
					for ($c=0; $c < $num; $c++) {
						if($data[$c] === 'Date' || $data[$c] === 'Link') {
							continue;
						} else {
							$arr[] = $data[$c];
						}
					}
					if(!$arr) {
						continue;
					}
					if($arr && count($arr)) {
						$currentId++;
						$this->process($arr, $currentId);
					}
				}

				try {
					$this->writer->save($this->resultsFile);
				} catch(\Exception $e) {
					throw new \Exception($e->getMessage());
				}

			}
		}

		public function process($arr, $currentId)
		{

			if(!$arr[0] || !$arr[1]) {
				return;
			}

			$parse = (object)parse_url($arr[1]);
			if(!property_exists($parse, 'host') || $parse->host === "") {
				return;
			}
			$date = strtotime($arr[0]);
			$startDate = date("Y-m", strtotime("-3 day", $date));
			$endDate = date("Y-m", strtotime("+3 day", $date));

			$url = "https://api.similarweb.com/v1/website/***/demographics/age?" .
				"format=json&".
				"api_key=$this->apiKeySim&".
				"country=us&".
				"start_date=$startDate&".
				"end_date=$endDate&".
				"main_domain_only=false&".
				"granularity=monthly&".
				"limit=0"
			;

			$url = str_replace('***', $parse->host, $url);

			$json = $this->get_a_curl($url);

			if ($json && is_object($json) && !empty($json) && !property_exists($json->meta,"error_code")) {

				$needed = [
					'id'     => $currentId,
					'link'   => $arr[1],
					'domain' => $parse->host,
					'18-24'  => $json->age_18_to_24,
					'25-34'  => $json->age_25_to_34,
					'35-44'  => $json->age_35_to_44,
					'45-54'  => $json->age_45_to_54,
					'55-64'  => $json->age_55_to_64,
					'+65'    => $json->age_65_plus,
				];

				$this->setSpreadSheetValues($needed);

			} else {

				$needed = [
					'id' => $currentId,
					'link'   => $arr[1],
					'domain' => $parse->host,
					'18-24' => "N/A",
					'25-34' => "N/A",
					'35-44' => "N/A",
					'45-54' => "N/A",
					'55-64' => "N/A",
					'+65' => "N/A",
				];

				$this->setSpreadSheetValues($needed);

			}
		}

		public function setSpreadSheetValues($needed) {
			/* headings */
			if($this->row === 1) {
				$colCount = 0;
				foreach($needed as $needKey => $needVal) {
					$cell = $this->alphabet[$colCount] . $this->row;
					$this->ageSheet->setCellValue($cell, $needKey);
					$colCount++;
				}
				$this->row += 1;
			}

			/* body */
			$colCount = 0;
			foreach($needed as $need) {
				$cell = $this->alphabet[$colCount] . $this->row;

				try{
					$this->ageSheet->setCellValue($cell, $need);
				} catch(\Exception $e) {
					echo new \Exception($e->getMessage());
				}
				$colCount++;
			}

			$this->row += 1;
			return;
		}

		public function get_a_curl($url) {
			$curl = curl_init();
			curl_setopt_array($curl, array(
				CURLOPT_URL => $url,
				CURLOPT_RETURNTRANSFER => true,
				CURLOPT_ENCODING => "",
				CURLOPT_MAXREDIRS => 10,
				CURLOPT_TIMEOUT => 0,
				CURLOPT_FOLLOWLOCATION => false,
				CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
				CURLOPT_CUSTOMREQUEST => "GET",
			));

			$response = curl_exec($curl);
			$err = curl_error($curl);

			curl_close($curl);

			if ($err) {
				var_dump("cURL Error #:" . $err); exit('problem');
			} else {
				return json_decode($response);
			}
		}
	}