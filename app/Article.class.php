<?php

	require 'vendor/autoload.php';

	use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
	use PhpOffice\PhpSpreadsheet\IOFactory;

	class Article {

		private $spreadsheet;
		private $articleSheet;
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

		public function __construct() {
			$this->sheetName = 'Article';
			$this->alphabet = range('A', 'Z');
			$this->apiKey = 'CmAHq1F6wW7UQ';
			$this->file = 'sample.csv';
			$this->resultsFile = 'results.xlsx';
			$this->url_articles = "https://api.newswhip.com/v1/articles?key=" . $this->apiKey;
			$this->row = 1;
			$this->queryRow = 1;
			$this->articleSheet = new Worksheet($this->spreadsheet, $this->sheetName);
			$this->querySheet = new Worksheet($this->spreadsheet, "Query");

			try {
				$this->reader = IOFactory::createReader("Xlsx");
				$this->spreadsheet = $this->reader->load($this->resultsFile);
				if($this->spreadsheet->sheetNameExists($this->sheetName)) {
					$sheetIndex = $this->spreadsheet->getIndex(
						$this->spreadsheet->getSheetByName($this->sheetName)
					);
					$this->spreadsheet->removeSheetByIndex($sheetIndex);
					$this->spreadsheet->addSheet($this->articleSheet);
				} else {
					$this->spreadsheet->addSheet($this->articleSheet);
				}

				$this->writer = IOFactory::createWriter($this->spreadsheet, "Xlsx");
			} catch(\Exception $e) {
				throw new \Exception($e->getMessage());
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

		public function process($arr, $currentId) {
			if(!$arr[0] || !$arr[1]) {
				return;
			}

			$date = strtotime($arr[0]) * 1000;
			$startDate = strtotime("-3 day", $date);
			$endDate = strtotime("+3 day", $date);
			$parse = (object) parse_url($arr[1]);
			if(!property_exists($parse, 'host') || $parse->host === "") {
				return;
			}
			$filter = json_encode(["href:\"$arr[1]\"", "domain:\"$parse->host\""]);
			$json = $this->get_a_curl($filter, $startDate, $endDate);

			if(isset($_GET['debug'])) {
				print '<pre>'.print_r($json, true).'</pre>';
			}

			sleep(1);

			if($json && is_object($json) && !empty($json->articles)) {

				foreach($json->articles as $d) {

					$needed = [
						'id'           => $currentId,
						'date'         => gmdate("m-d-y", $d->publication_timestamp / 1000),
						'link'         => $d->link, //$arr[1]
						'facebook'     => !empty($d->fb_data->total_engagement_count) ? $d->fb_data->total_engagement_count : "N/A",
						'twitter'      => !empty($d->tw_data->tw_count) ? $d->tw_data->tw_count: "N/A",
						'linkedIn'     => !empty($d->li_data->li_count) ? $d->li_data->li_count : "N/A",
						'publisher'    => !empty($d->authors) ? implode(",", $d->authors) : "N/A",
						'nw_max'       => !empty($d->max_nw_score) ? $d->max_nw_score : "N/A",
					];

					$this->setSpreadSheetValues($needed);
				}

				/* query tab */
				$queryArr = [
					'id'           => $currentId,
					'link_href'    => $arr[1],
					'domain_href'  => "",
					'date'         => $arr[0],
				];

				$this->doQueryTab($queryArr);

			} else {

				$now = strtotime(date("m/d/y"));
				$weekAgo = date("m/d/y", strtotime("-7 day", $now));

				$needed = [
					'id'        => $currentId,
					'date'      => $weekAgo,
					'link'      => $arr[1],
					'facebook'  => 'N/A',
					'twitter'   => 'N/A',
					'linkedIn'  => 'N/A',
					'publisher' => 'N/A',
					'nw_max'    => 'N/A',
				];

				/* query tab */
				$queryArr = [
					'id'           => $currentId,
					'link_href'    => $arr[1],
					'domain_href'  => "",
					'date' => $arr[0],
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
					$this->articleSheet->setCellValue($cell, $needKey);
					$colCount++;
				}
				$this->row += 1;
			}

			/* body */
			$colCount = 0;
			foreach($needed as $need) {
				$cell = $this->alphabet[$colCount] . $this->row;

				try{
					$this->articleSheet->setCellValue($cell, $need);
				} catch(\Exception $e) {
					echo new \Exception($e->getMessage());
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

			/* query headings */
			if($this->queryRow === 1) {
				$colCount = 0;
				foreach($queryData as $needKey => $needVal) {
					$cell = $this->alphabet[$colCount] . $this->queryRow;
					try{
						$this->querySheet->setCellValue($cell, $needKey);
					} catch(\Exception $e) {
						echo new \Exception($e->getMessage());
					}
					$colCount++;
				}
				$this->queryRow += 1;
			}

			/* query body */
			$colCount = 0;
			foreach($queryData as $need) {
				if($colCount === 4) {
					$colCount++;
					continue;
				} else {
					$cell = $this->alphabet[$colCount] . $this->queryRow;
				}

				try{
					$this->querySheet->setCellValue($cell, $need);
				} catch(\Exception $e) {
					echo new \Exception($e->getMessage());
				}
				$colCount++;
			}
			$this->queryRow += 1;
			return;
		}

		public function get_a_curl($filter, $from, $to) {

			$curl_articles = "
				curl -H \"Content-Type: application/json\" -X POST -d '{
					\"filters\": $filter,
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
						\"filters\": $filter,
						\"size\": 5000
					}' \"https://api.newswhip.com/v1/articles?key=$this->apiKey\"
				";

				$exec1 = shell_exec($curl_articles_no_date);
				return json_decode($exec1);
			}
		}
	}

	new Article();
