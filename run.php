<?php

	require 'vendor/autoload.php';

	use App\Article1;
	use App\Domain;
	use App\Gender;
	use App\Age;
	use App\PageView;
	use App\Summary;
	use App\Input;

	$start = date('H:i:s');
	$inputFile = dirname(__FILE__) . DIRECTORY_SEPARATOR . "input" . DIRECTORY_SEPARATOR . "sample.csv";
	$outputFile = dirname(__FILE__) . DIRECTORY_SEPARATOR . "output" . DIRECTORY_SEPARATOR . "results.xlsx";

	if(!file_exists($outputFile)) {
		shell_exec("touch " . $outputFile);
	}

	$apiKeyNewsWhip = "CmAHq1F6wW7UQ";
	$apiKeySimilarWeb = "3dd3e5ad256144738582bb3776f1acc3";

	try {
		new Article1($apiKeyNewsWhip, $inputFile, $outputFile, "Article");
		new Domain($apiKeyNewsWhip, $inputFile, $outputFile, "Domain");
		new Gender($apiKeySimilarWeb, $inputFile, $outputFile, "Gender");
		new Age($apiKeySimilarWeb, $inputFile, $outputFile, "Age");
		new PageView($apiKeySimilarWeb, $inputFile, $outputFile, "PageView");
		new Input($inputFile, $outputFile, "InputData");
		new Summary($inputFile, $outputFile, "Summary");

	} catch(\Exception $e) {
		throw new \Exception($e->getMessage());
	}

	$end = date('H:i:s');
	echo abs(strtotime($end) - strtotime($start)) . 'sec';
	shell_exec("cp " . $outputFile . " output/results-" . date("d-m-Y-H-i-s") . ".xlsx");
	shell_exec("rm " . $outputFile);