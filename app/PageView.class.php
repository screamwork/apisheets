<?php

	namespace App;

	use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
	use PhpOffice\PhpSpreadsheet\IOFactory;

	class PageView {

		private $apiKeySimilarWeb;
		private $file;
		private $pageViewSheet;
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
			$this->resultsFile = $resultsFile;
			$this->apiKeySimilarWeb  = $apiKey;
			$this->file = $file;
			$this->alphabet = range( 'A', 'Z' );
			$this->row = 1;
			$this->pageViewSheet = new Worksheet($this->spreadsheet, $this->sheetName);

			try {
				$this->reader = IOFactory::createReader("Xlsx");
				$this->spreadsheet = $this->reader->load($this->resultsFile);
				if($this->spreadsheet->sheetNameExists($this->sheetName)) {
					$sheetIndex = $this->spreadsheet->getIndex(
						$this->spreadsheet->getSheetByName($this->sheetName)
					);
					$this->spreadsheet->removeSheetByIndex($sheetIndex);
					$this->spreadsheet->addSheet($this->pageViewSheet);
				} else {
					$this->spreadsheet->addSheet($this->pageViewSheet);
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

		public function process($arr, $currentId)	{

			if(!$arr[0] || !$arr[1]) {
				return;
			}

			$parse = (object) parse_url($arr[1]);
			if(!property_exists($parse, 'host') || $parse->host === "") {
				return;
			}
			$date = strtotime($arr[0]);
			$startDate = date("Y-m", strtotime("-3 day", $date));
			$endDate = date("Y-m", strtotime("+3 day", $date));

			$url = "https://api.similarweb.com/v1/website/***/traffic-and-engagement/pages-per-visit?" .
				"format=json&".
				"api_key=$this->apiKeySimilarWeb&".
				"country=us&".
				"start_date=$startDate&".
				"end_date=$endDate&".
				"main_domain_only=false&".
				"granularity=monthly"
			;
			$url = str_replace('***', $parse->host, $url);
			$json = $this->get_a_curl($url);

			if ($json && is_object($json) && property_exists($json->meta, 'status') && $json->meta->status === "Success") {

				foreach($json->pages_per_visit as $visit) {
					$needed = [
						'id' => $currentId,
						'date' => $arr[0],
						'link' => $arr[1],
						'page_views' => $visit->pages_per_visit,
					];
					$this->setSpreadSheetValues($needed);
				}

			} else {

				$needed = [
					'id'     => $currentId,
					'date'   => $arr[0],
					'link'   => $arr[1],
					'page_views' => "N/A",
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
					$this->pageViewSheet->setCellValue($cell, $needKey);
					$colCount++;
				}
				$this->row += 1;
			}

			/* body */
			$colCount = 0;
			foreach($needed as $need) {
				$cell = $this->alphabet[$colCount] . $this->row;

				try{
					$this->pageViewSheet->setCellValue($cell, $need);
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