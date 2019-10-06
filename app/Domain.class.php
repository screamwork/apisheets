<?php

	namespace App;

	use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
	use PhpOffice\PhpSpreadsheet\IOFactory;

	class Domain {

		private $spreadsheet;
		private $domainSheet;
		private $alphabet;
		private $apiKey;
		private $file;
		private $resultsFile;
		private $url_articles;
		private $row;
		private $sheetName;
		private $writer;
		private $reader;
		private $querySheet;
		private $queryRow;

		public function __construct($apiKey = "CmAHq1F6wW7UQ", $file = 'sample.csv', $resultsFile = "results.xlsx", $sheetName = false) {
			if(!$sheetName) {
				throw new \Exception("no sheetName");
			}
			$this->sheetName = $sheetName;
			$this->alphabet = range( 'A', 'Z' );
			$this->apiKey = $apiKey;
			$this->file = $file;
			$this->resultsFile = $resultsFile;
			$this->url_articles = "https://api.newswhip.com/v1/articles?key=" . $this->apiKey;
			$this->row = 1;
			$this->queryRow = 1;
			$this->domainSheet = new Worksheet($this->spreadsheet, $this->sheetName);
			$this->querySheet = new Worksheet($this->spreadsheet, "Query");

			try {
				$this->reader = IOFactory::createReader("Xlsx");
				$this->spreadsheet = $this->reader->load($this->resultsFile);
				if($this->spreadsheet->sheetNameExists($this->sheetName)) {
					$sheetIndex = $this->spreadsheet->getIndex(
						$this->spreadsheet->getSheetByName($this->sheetName)
					);
					$this->spreadsheet->removeSheetByIndex($sheetIndex);
					$this->spreadsheet->addSheet($this->domainSheet);
				} else {
					$this->spreadsheet->addSheet($this->domainSheet);
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
					throw new \Exception($e);
				}
			}
		}

		public function process($arr, $currentId) {

			if(!$arr[0] || !$arr[1]) {
				return;
			}

			$parse = (object) parse_url($arr[1]);
			if(!property_exists($parse, 'host') || $parse->host === "") {
				return;
			}
			$exp = explode(".", $parse->host);

			if(count($exp) === 3) {
				$host = $exp[1] . '.' . $exp[2];
			} elseif(count($exp) === 2) {
				$host = $exp[0] . '.' . $exp[1];
			} else {
				$host = $exp[0];
			}

			$date = strtotime($arr[0]) * 1000;
			$startDate = strtotime("-3 day", $date);
			$endDate = strtotime("+3 day", $date);
			$filter = json_encode("domain:\"$host\"");
			$json = $this->get_a_curl($filter, $startDate, $endDate);

			sleep(1);

			if($json && is_object($json) && !empty($json->articles)) {

				$needed = [
					'id'            => $currentId,
					'date'          => $arr[0], // date('Y-m-d', strtotime($arr[0])),
					'link'          => $arr[1],
					'domain'        => $parse->host,
					'article_count' => count($json->articles)
				];

				/* query tab */
				$queryArr = [
					'domain_href' => $host,
				];

				$this->doQueryTab($queryArr);
				$this->setSpreadSheetValues($needed);

			} else {
				$now = strtotime(date("m/d/y"));
				$weekAgo = date("m/d/y", strtotime("-7 day", $now));

				$needed = [
					'id'            => $currentId,
					'date'          => $weekAgo,
					'link'          => $arr[1],
					'domain'        => $parse->host,
					'article_count' => "N/A"
				];

				/* query tab */
				$queryArr = [
					'domain_href' => $host,
				];

				$this->doQueryTab($queryArr);
				$this->setSpreadSheetValues($needed);

			}
		}

		public function setSpreadSheetValues($needed) {
			/* headings */
			if($this->row === 1) {
				$colCount = 0;
				foreach($needed as $needKey => $needVal) {
					$cell = $this->alphabet[$colCount] . $this->row;
					$this->domainSheet->setCellValue($cell, $needKey);
					$colCount++;
				}
				$this->row += 1;
			}

			/* body */
			$colCount = 0;
			foreach($needed as $need) {
				$cell = $this->alphabet[$colCount] . $this->row;

				try{
					$this->domainSheet->setCellValue($cell, $need);
				} catch(\Exception $e) {
					echo new \Exception($e);
				}
				$colCount++;
			}
			$this->row += 1;
			return;
		}

		public function doQueryTab($queryData) {

			try {
				if (!$this->spreadsheet->sheetNameExists("Query")) {
					$this->spreadsheet->addSheet($this->querySheet);
				}
				else {
					$this->querySheet = $this->spreadsheet->getSheetByName("Query");
				}
			} catch(\Exception $e) {
				throw new \Exception($e->getMessage());
			}

			/* query body */
			foreach($queryData as $need) {
				if($this->queryRow === 1) {
					$this->queryRow++;
				}
				$cell = $this->alphabet[2] . $this->queryRow;

				try{
					$this->querySheet->setCellValue($cell, $need);
				} catch(\Exception $e) {
					echo new \Exception($e->getMessage());
				}
			}
			$this->queryRow += 1;
			return;
		}

		public function get_a_curl($filter, $from, $to) {

			$curl_articles = "
				curl -H \"Content-Type: application/json\" -X POST -d '{
					\"filters\": [$filter],
					\"from\": $from,
					\"to\": $to,
					\"size\": 5000
				}' \"https://api.newswhip.com/v1/articles?key=$this->apiKey\"
			";

			$exec = shell_exec($curl_articles);
			$json = json_decode($exec);

			if($json && is_object($json) && !empty($json->articles)) {
				return $json;
			} else {
				$curl_articles_no_date = "
					curl -H \"Content-Type: application/json\" -X POST -d '{
						\"filters\": [$filter],
						\"size\": 5000
					}' \"https://api.newswhip.com/v1/articles?key=$this->apiKey\"
				";

				$exec1 = shell_exec($curl_articles_no_date);
				return json_decode($exec1);
			}
		}
	}
