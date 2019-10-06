<?php

	namespace App;

	use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
	use PhpOffice\PhpSpreadsheet\IOFactory;

	class Summary
	{
		private $file;
		private $summarySheet;
		private $row;
		private $sheetName;
		private $writer;
		private $reader;
		private $resultsFile;
		private $spreadsheet;
		private $alphabet;

		public function __construct($file = 'sample.csv', $resultsFile = "results.xlsx", $sheetName = false) {
			if(!$sheetName) {
				throw new \Exception("no sheetName");
			}
			$this->sheetName = $sheetName;
			$this->resultsFile = $resultsFile;
			$this->file = $file;
			$this->alphabet = range('A', 'Z');
			$this->row = 1;
			$this->summarySheet = new Worksheet($this->spreadsheet, $this->sheetName);

			try {
				$this->reader = IOFactory::createReader("Xlsx");
				$this->spreadsheet = $this->reader->load($this->resultsFile);
				if ($this->spreadsheet->sheetNameExists($this->sheetName)) {
					$sheetIndex = $this->spreadsheet->getIndex(
						$this->spreadsheet->getSheetByName($this->sheetName)
					);
					$this->spreadsheet->removeSheetByIndex($sheetIndex);
					$this->spreadsheet->addSheet($this->summarySheet);
				} else {
					$this->spreadsheet->addSheet($this->summarySheet);
				}
				$this->writer = IOFactory::createWriter($this->spreadsheet, "Xlsx");
			} catch (\Exception $e) {
				throw new \Exception($e);
			}

			$this->readCsv();
		}

		public function readCsv()
		{
			if (($handle = fopen($this->file, "r")) !== false) {
				while (($data = fgetcsv($handle, 1000, ",")) !== false) {
					$num = count($data);

					$arr = [];
					for ($c = 0; $c < $num; $c++) {
						if ($data[$c] === 'Date' || $data[$c] === 'Link') {
							continue;
						} else {
							$arr[] = $data[$c];
						}
					}
					if (!$arr) {
						continue;
					}
					if ($arr && count($arr)) {
						$needed = [
							'date'   => $arr[0],
							'link'   => $arr[1],
						];

						/* headings */
						if($this->row === 1) {
							$colCount = 0;
							foreach($needed as $needKey => $needVal) {
								$cell = $this->alphabet[$colCount] . $this->row;
								$this->summarySheet->setCellValue($cell, $needKey);
								$colCount++;
							}
							$this->row += 1;
						}

						/* body */
						$colCount = 0;
						foreach($needed as $need) {
							$cell = $this->alphabet[$colCount] . $this->row;

							try{
								$this->summarySheet->setCellValue($cell, $need);
							} catch(\Exception $e) {
								echo new \Exception($e->getMessage());
							}
							$colCount++;
						}
						$this->row += 1;
					}

					try {
						$this->writer->save($this->resultsFile);
					} catch (\Exception $e) {
						throw new \Exception($e->getMessage());
					}
				}
			}
		}
	}