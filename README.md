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
				columnAutoSize'=>'true', //false/true; default True
				
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
                //'sheet_title' => $excel_title1, 					//old version
                //'sheet_title' => [$excel_title1], 				//new version
                'sheet_title' => [$excel_title1,$excel_title2], 	//new version				
			    'ceils' => $excel_ceilsAll,
				//'freezePane' => 'E2',
                'headerColor' => Postman4ExcelBehavior::getCssClass("header"),	//All Header Color font and Backgroud
                'headerColumnCssClass' => [
					  [
						'id' => Postman4ExcelBehavior::getCssClass('yellow'),
						'username' => Postman4ExcelBehavior::getCssClass('green'),
					 ],
					 [
						 'ID' => Postman4ExcelBehavior::getCssClass('red'),  
						 'USERNAME' => Postman4ExcelBehavior::getCssClass('green'), 						 
					 ]             
                             
                ],
               'oddCssClass' => Postman4ExcelBehavior::getCssClass("odd"),
               'evenCssClass' => Postman4ExcelBehavior::getCssClass("even"),
			],
			[
				'sheet_name' => 'TEST EXPORT 2',
                //'sheet_title' => $excel_title1, 					//old version
                //'sheet_title' => [$excel_title1], 				//new version
                'sheet_title' => [$excel_title1,$excel_title2], 	//new version				
			    'ceils' => $excel_ceilsAll,
				//'freezePane' => 'E2',
                'headerColor' => Postman4ExcelBehavior::getCssClass("header"),	//All Header Color font and Backgroud
               
			    //Deafault Header properties
				'headerColumnCssClass' => [
					  [
						'id' => Postman4ExcelBehavior::getCssClass('red'),
						'username' => Postman4ExcelBehavior::getCssClass('green'),
					 ],
					 [
						 'ID' => Postman4ExcelBehavior::getCssClass('red'),  
						 'USERNAME' => Postman4ExcelBehavior::getCssClass('green'), 						 
					 ]             
                             
                ],
				//customize Header properties
				// Color Ref: http://dmcritchie.mvps.org/excel/colors.htm
				/*
				'headerStyle' => [
					[
						'id' => ['align'=>'CENTER','color-font'=>'0000FF','color-background'=>'FFCCCC','merge'=>'0,2','width'=>'20'], //columnAutoSize=false, width is Active
						'username' => ['align'=>'left','color-font'=>'FF0000','color-background'=>'CCFF99','merge'=>'2,1'], //'merge'=>'col,row'
					],
					[
						 'ID' =>  ['align'=>'right','color-font'=>'0000FF','color-background'=>'CCFFCC'],
						 'USERNAME' =>  ['align'=>'right','color-font'=>'000000','color-background'=>'FFFF99'],				 
					]                              
                ],  
				*/
				//customize Content properties
				// Content properties, validate on use last column title
				'contentStyle' => [
					  [
						'id' => ['align'=>'left'],
						'username' => ['align'=>'right'],
					 ],
					 [
						 'ID' =>  ['align'=>'center','color-font'=>'0000FF'],
						 'USERNAME' =>  ['align'=>'center','color-font'=>'0000FF'],				 
					 ], 
					// [
					//	 'id' =>  ['align'=>'right'],
					//	 'username' =>  ['align'=>'right'],				 
					// ] 					 
                             
                ],  
               'oddCssClass' => Postman4ExcelBehavior::getCssClass("odd"),
               'evenCssClass' => Postman4ExcelBehavior::getCssClass("even"),
			],
			
		];
		
		$excelFile = "TestExport";
		$this->export4excel($excel_content, $excelFile); 	
    }
```

