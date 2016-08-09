Postman Excel for Yii 2
=======================

Postman Excel Export for view or console cronjobs

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist ptrnov/yii2-postman4excel "dev-master"
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
```

class YourControllerController extends Controller
{
	public function behaviors()
    {
        return [
			//Postman4ExcelBehavior::className(),
			'export4excel' => [
				'class' => Postman4ExcelBehavior::className(),
			], 
			'verbs' => [
                'class' => VerbFilter::className(),
                'actions' => [
                    'delete' => ['post'],
                ],
            ]
        ];
    }
	
	public function actionTestExport1()
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
		//$excel_title = $excel_dataNKA['excel_title'];
		$excel_ceilsAll = $excel_dataAll['excel_ceils'];
		
		
		//Path for Export The Data
		//default path is empty:(Yii::getAlias('@vendor').'/ptrnov/yii2-postman4excel/tmp');
		//path of windows, example "d:/folder/"; folder nama should be exist in path
		$this->downloadPath = ''; //'d:/tools/
		
		//Widget Type set for download or cronjob or mail else mix
		//widgetType=download
		//widgetType=cronjob
		//widgetType=mail
		$this->widgetType='cronjob';
		
		$excel_content = [
			[
				'sheet_name' => 'TEST EXPORT 1',
                'sheet_title' => ['ID','USERNAME'],
			    'ceils' => $excel_ceilsAll,
				'ceils_start_rows'=>1,
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
		
		$excel_file = "TestExport";
		$this->export2excel($excel_content, $excel_file,0); 	
    }
}