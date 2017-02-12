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
		$excel_dataAll = Postman4ExcelBehavior::excelDataFormat($arySqlDataProvider);	(not used for columnGroup just put 'ceils' => $arySqlDataProvider)
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
				'sheet_name' => 'TEST EXPORT 1 old version',
                'sheet_title' => $excel_title1, 					//old version				
			    'ceils' => $excel_ceilsAll,// atau langsung $arySqlDataProvider;
					//for use column Group.
					//noted: 'ceils' => "Array Source", not need difinition "Postman4ExcelBehavior::excelDataFormat($array source);"
					//'columnGroup'=>Name of Field//column group difinition.
				'columnGroup'=>'CUST_NM',//column for grouping.
				'autoSize'=>true,	//true/false.
				//'freezePane' => 'A2',
                'headerColor' => Postman4ExcelBehavior::getCssClass("header"),	//All Header Color font and Backgroud
                'headerColumnCssClass' => [							//old version
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
				'sheet_name' => 'TEST EXPORT 2 new version ',
				//'sheet_title' => [$excel_title1], 				//new version | one Header
                'sheet_title' => [									//new version | two or more Header
					$excel_title1,
					$excel_title2
				], 	//new version				
			    'ceils' => $excel_ceilsAll,
				//'freezePane' => 'E2',
                'headerColor' => Postman4ExcelBehavior::getCssClass("header"),	//old Version | All Header Color font and Backgroud            
				/* --------------------------------------------------------
				 * customize Header properties
				 * Color Ref: http://dmcritchie.mvps.org/excel/colors.htm
				 * columnAutoSize=false, width is Active
				 * 'merge'=>'col,row'
				 * Content properties, validate on use last column title
				 * --------------------------------------------------------
				 */
				'headerStyle' => [									//new version
					[	//the first Header						
						'id' => ['align'=>'CENTER','color-font'=>'0000FF','color-background'=>'FFCCCC','merge'=>'0,2','width'=>'2,0'],
						'username' => ['align'=>'left','color-font'=>'FF0000','color-background'=>'CCFF99','merge'=>'2,1','width'=>'32.29','valign'=>'center','wrap'=>true], 
					],
					[	//The second Header
						 'ID' =>  ['align'=>'right','color-font'=>'0000FF','color-background'=>'CCFFCC'],
						 'USERNAME' =>  ['align'=>'right','color-font'=>'000000','color-background'=>'FFFF99'],				 
					]                              
                ],  
				'contentStyle' => [									//new version
					[//the first Content
						//'id' => ['align'=>'left'],
						//'username' => ['align'=>'right'],
					],
					[//The second Content
						'ID' =>  ['align'=>'center','color-font'=>'0000FF'],
						'USERNAME' =>  ['align'=>'center','color-font'=>'0000FF'],				 
					],                              
                ],  
               'oddCssClass' => Postman4ExcelBehavior::getCssClass("odd"),
               'evenCssClass' => Postman4ExcelBehavior::getCssClass("even"),
			],
			
		];
		
		$excelFile = "TestExport";
		$this->export4excel($excel_content, $excelFile); 	
    }
```

