Postman Excel for Yii 2
=======================
Base On scotthuangzl (ptrnov update)

Postman Excel Export for view or console cronjobs

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist ptrnov/yii2-postman4excel "*"
```

or add

```
"ptrnov/yii2-postman4excel": "*"
```

to the require section of your `composer.json` file.


Usage
-----

Once the extension is installed, simply use it in your code by  :

```php
	use ptrnov\postman4excel\Postman4ExcelBehavior;
	public function behaviors(){
			'export4excel' => [
				'class' => Postman4ExcelBehavior::className(),
				'downloadPath'=>'your path'		//defult "/vendor/ptrnov/yii2-postman4excel/tmp/", Ecample Windows path "d:/folder/"
				'widgetType'=>'download' 		//download web browser, delete before download, tmp_download
				//'widgetType'=>'cronjob' 	 	//console method, folder tmp_cronjob
				//'widgetType'=>'mail'		 	//posman mail, folder tmp_mail
				//'widgetType'=>''	 			//Empty same download, file  stay on folder "tmp_mix"
				//'prefixStr' => yii::$app->user->identity->username,
				//'suffixStr' => date('Ymd-His'),
			],
	}
	// localhost/yourController/test-export
	public function actionTestExport()
    {
		//get data from database
		$sqlDataProvider= new ArrayDataProvider([
			'allModels'=>\Yii::$app->db->createCommand("	
				SELECT id, username FROM user
			")->queryAll(), 
		]);	
		$arySqlDataProvider=$sqlDataProvider->allModels;	
		
		/*Array Model Data*/
		$excel_dataAll = Postman4ExcelBehavior::excelDataFormat($arySqlDataProvider);
		//$excel_title = $excel_dataAll['excel_title'];
		$excel_ceilsAll = $excel_dataAll['excel_ceils'];
		
		
		//Path for Export The Data
		//default path is empty:(Yii::getAlias('@vendor').'/ptrnov/yii2-postman4excel/tmp');
		//path of windows, example "d:/folder/"; folder nama should be exist in path
		$this->downloadPath = ''; //'d:/tools/
		
		$excel_content = [
			[
				'sheet_name' => 'TEST EXPORT 1',
                'sheet_title' => ['ID','USERNAME'],
			    'ceils' => $excel_ceilsAll,
				'ceils_start_rows'=>2, // header 1 or 2
                //'freezePane' => 'E2',
                'headerColor' => Postman4ExcelBehavior::getCssClass("header"),
                'headerColumnCssClass' => [
					 'id' => Postman4ExcelBehavior::getCssClass('header'),
                     'username' => Postman4ExcelBehavior::getCssClass('header'),                   
                ], 
               'oddCssClass' => Postman4ExcelBehavior::getCssClass("odd"),
               'evenCssClass' => Postman4ExcelBehavior::getCssClass("even"),
			],
		];
		
		$excelFile = "TestExport";
		$this->export4excel($excel_content, $excelFile); 	
    }
```

