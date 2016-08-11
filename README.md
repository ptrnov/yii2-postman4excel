Postman Excel for Yii 2
=======================
Base On scotthuangzl (ptrnov update)

Postman Excel Export for view or console cronjobs, mail postman

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
		$excel_ceilsAll = $excel_dataAll['excel_ceils'];
		$excel_title1 = $excel_dataAll['excel_title'];
		$excel_title2 = ['ID','USERNAME'];
		
		//Tite Header From Table Name
		//$excel_title = $excel_dataAll['excel_title'];
		
		//Tite Header From Modify byself
		//$excel_title = ['ID','USERNAME']
				
		//Note 'sheet_title'
		//old version : 'sheet_title' => $excel_title,
		//new Version : 'sheet_title' => [$excel_title1,$excel_title2,$excel_title3, dst],
		
		
		$excel_content = [
			[
				'sheet_name' => 'TEST EXPORT 1',
                'sheet_title' => $excel_title1, 					//old version
                //'sheet_title' => [$excel_title1], 				//new version
                //'sheet_title' => [$excel_title1,$excel_title2], 	//new version				
			    'ceils' => $excel_ceilsAll,
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

