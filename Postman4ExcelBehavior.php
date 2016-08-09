<?php
/**
 * User: Scott_Huang
 * Date: 6/16/2015
 * Time: 5:16 PM
 */
namespace  ptrnov\postman4excel;

use yii\base\Behavior;
use yii\helpers\Url;
use Yii;
use \PHPExcel;
use \PHPExcel_IOFactory;
use \PHPExcel_Settings;
use \PHPExcel_Style_Fill;
use \PHPExcel_Writer_IWriter;
use \PHPExcel_Worksheet;
use \PHPExcel_Style;
use \PHPExcel_Style_Border;

/**
 * @var string
 * @author ptrnov [ptr.nov@gmail.com]
 * @since 1.0.0
 * @state Indonesia
 */
class Postman4ExcelBehavior extends Behavior
{	
	 /**
     * @var string  you can set use logged username , it will add in the file as prefix
     * usually you can set as yii::$app->user->identity->username
     */
    public $prefixStr = '';

    /**
     * @var string
     */
    public $suffixStr = ''; //default will be date('Ymd-His')
	public $startRowsValue='';

	
	const TYPE_DEFAULT = 'download';
    const TYPE_CRONJOB = 'cronjob';
    const TYPE_MAIL = 'mail';
	
	/**
	* @var string
	* path destination download file
	*/
	public $downloadPath = '';	
	
	/**
	* @var string
	* widgetType: download|cronjob|mail
	* $downloadPath.Folder (folder inside downloadPath )
	* Normal Folder Download | Cronjob Folder Download | Mail Folder Download to Send	
	*/
	public $widgetType = '';	
		
	/**
	* @var string
	* Setting path
	* @author ptrnov [ptr.nov@gmail.com]
	* @since 1.0.0
	*/
	private function getPath(){
		//$defaultPath=dirname(__DIR__).'vendor';
		$defaultPath=Yii::getAlias('@vendor').'/ptrnov/yii2-postman4excel/tmp/';
		$newPath=$this->downloadPath;
		return  $newPath!=''?$newPath:$defaultPath;
	}
	
	/**
	* @var string
	* WidgetType: download|cronjob|mail
	* Normal Folder tmp_download or tmp_cronjob or tmp_mail else tmp_mix
	*
	* @author ptrnov [ptr.nov@gmail.com]
	* @since 1.0.0
	*/
	private function getFolder(){
		$folder=strtoupper($this->widgetType);
		if ($folder=='DOWNLOAD'){
			$folder_='tmp_download';
		}elseif($folder=='CRONJOB'){
			$folder_='tmp_cronjob';
		}elseif($folder=='MAIL'){
			$folder_='tmp_mail';
		}else{
			$folder_='tmp_mix';
		};
		
		$tempDir=$this->getPath().$folder_;
		
		if (!is_dir($tempDir)) {
			if (!is_dir($folder_)) {
				mkdir($folder_);
			}
			mkdir($tempDir);
			chmod($tempDir, 0755);
			// the default implementation makes it under 777 permission, 
			// which you could possibly change recursively before deployment, but here's less of a headache in case you don't
			return $tempDir.'/';
		}else{
			throw new Exception("Invalid path");
		}
	}
		
    /**
     * Return query contents to predefined sheet format
     *
     * @param $data
     * @return array
     */
    public static function excelDataFormat($data)
    {
        for ($i = 0; $i < count($data); $i++) {
            $each_arr = $data[$i];
            $new_arr[] = array_values($each_arr); 
        }
        $new_key[] = array_keys($data[0]); 
        return array('excel_title' => $new_key[0], 'excel_ceils' => $new_arr);
    }


    /**
     * Returns the coresponding excel column.(Abdul Rehman from yii forum)
     *
     * @param int $index
     * @return string
     * @throws Exception
     */
    public static function excelColumnName($index)
    {
        --$index;
        if ($index >= 0 && $index < 26)
            return chr(ord('A') + $index);
        else if ($index > 25)
            return (self::excelColumnName($index / 26)) . (self::excelColumnName($index % 26 + 1));
        else
            throw new Exception("Invalid Column # " . ($index + 1));
    }


    /**
     * save predefined sheet contents to excel
     *
     * @param $excel_content
     * @param $excel_file
     * @param array $excel_props
     * @return bool|string
     * @throws Exception
     */
    public function save2Excel($excel_content, $excel_file
        , $excel_props = array('creator' => 'WWSP Tool'
        , 'title' => 'WWSP_Tracking EXPORT EXCEL'
        , 'subject' => 'WWSP_Tracking EXPORT EXCEL'
        , 'desc' => 'WWSP_Tracking EXPORT EXCEL'
        , 'keywords' => 'WWSP Tool Generated Excel, Author: Scott Huang'
        , 'category' => 'WWSP_Tracking EXPORT EXCEL'))
    {
        if (!is_array($excel_content)) {
            return FALSE;
        }
        if (empty($excel_file)) {
            return FALSE;
        }
        $objPHPExcel = new PHPExcel();
        $objProps = $objPHPExcel->getProperties();
        $objProps->setCreator($excel_props['creator']);
        $objProps->setLastModifiedBy($excel_props['creator']);
        $objProps->setTitle($excel_props['title']);
        $objProps->setSubject($excel_props['subject']);
        $objProps->setDescription($excel_props['desc']);
        $objProps->setKeywords($excel_props['keywords']);
        $objProps->setCategory($excel_props['category']);

        $style_obj = new PHPExcel_Style();
        $style_array = array(
		  'borders' => array(
               'top' => array('style' => PHPExcel_Style_Border::BORDER_THICK),
               'left' => array('style' => PHPExcel_Style_Border::BORDER_THICK),
               'bottom' => array('style' => PHPExcel_Style_Border::BORDER_THICK),
               'right' => array('style' => PHPExcel_Style_Border::BORDER_THICK)
           ),
            'alignment' => array(
                'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                'vertical' => \PHPExcel_Style_Alignment::VERTICAL_CENTER,
                'wrap' => true
            ),
        );
        $style_obj->applyFromArray($style_array);

//开始执行EXCEL数据导出
        //start export excel
        for ($i = 0; $i < count($excel_content); $i++) {
            $each_sheet_content = $excel_content[$i];
            if ($i == 0) {
//默认会创建一个sheet页，故不需在创建
                //There will be a default sheet, so no need create
                $objPHPExcel->setActiveSheetIndex(intval(0));
                $current_sheet = $objPHPExcel->getActiveSheet();
            } else {
//创建sheet
                //create sheet
                $objPHPExcel->createSheet();
                $current_sheet = $objPHPExcel->getSheet($i);
            }
//设置sheet title
            //set title
            $current_sheet->setTitle(str_replace(array('/', '*', '?', '\\', ':', '[', ']'), array('_', '_', '_', '_', '_', '_', '_'), substr($each_sheet_content['sheet_name'], 0, 30))); //add by Scott
            $current_sheet->getColumnDimension()->setAutoSize(true); //Scott, set column autosize
//设置sheet当前页的标题
            //set sheet's current title
            $_columnIndex = 'A';

            $lineRange = "A1:" . self::excelColumnName(count($each_sheet_content['sheet_title'])) . "1";
            $current_sheet->setSharedStyle($style_obj, $lineRange);

            if (array_key_exists('sheet_title', $each_sheet_content) && !empty($each_sheet_content['sheet_title'])) {
                //header color
                if (array_key_exists('headerColor', $each_sheet_content) && is_array($each_sheet_content['headerColor']) and !empty($each_sheet_content['headerColor'])) {
                    if (isset($each_sheet_content['headerColor']["color"]) and $each_sheet_content['headerColor']['color'])
                        $current_sheet->getStyle($lineRange)->getFont()->getColor()->setARGB($each_sheet_content['headerColor']['color']);
                    //background
                    if (isset($each_sheet_content['headerColor']["background"]) and $each_sheet_content['headerColor']['background']) {
                        $current_sheet->getStyle($lineRange)->getFill()->getStartColor()->setRGB($each_sheet_content['headerColor']["background"]);
                        $current_sheet->getStyle($lineRange)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    }
                }

                for ($j = 0; $j < count($each_sheet_content['sheet_title']); $j++) {
                    $current_sheet->setCellValueByColumnAndRow($j, 1, $each_sheet_content['sheet_title'][$j]);
                    //start handle hearder column css
                    if (array_key_exists('headerColumnCssClass', $each_sheet_content)) {
                        if (isset($each_sheet_content["headerColumnCssClass"][$each_sheet_content['sheet_title'][$j]])) {
                            $tempStyle = $each_sheet_content["headerColumnCssClass"][$each_sheet_content['sheet_title'][$j]];
                            $tempColumn = self::excelColumnName($j + 1) . "1";
                            if (isset($tempStyle["color"]) and $tempStyle['color'])
                                $current_sheet->getStyle($tempColumn)->getFont()->getColor()->setARGB($tempStyle['color']);
                            //background
                            if (isset($tempStyle["background"]) and $tempStyle['background']) {
                                $current_sheet->getStyle($tempColumn)->getFill()->getStartColor()->setRGB($tempStyle["background"]);
                                $current_sheet->getStyle($tempColumn)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                            }
                        }
                    }
                    $current_sheet->getColumnDimension($_columnIndex)->setAutoSize(true); //
                    $_columnIndex++;
                }
            }
            if (array_key_exists('freezePane', $each_sheet_content) && !empty($each_sheet_content['freezePane'])) {
                $current_sheet->freezePane($each_sheet_content['freezePane']);
            }
//写入sheet页面内容
            //write sheet content
			$StartRowValContent=$each_sheet_content['ceils_start_rows'];
			$StartRowValueContent=$StartRowValContent!=''?$StartRowValContent:2; //author@ptr.nov
            if (array_key_exists('ceils', $each_sheet_content) && !empty($each_sheet_content['ceils'])) {
                for ($row = 0; $row < count($each_sheet_content['ceils']); $row++) {
                    //setting row css
                    $lineRange = "A" . ($row + $StartRowValueContent) . ":" . self::excelColumnName(count($each_sheet_content['ceils'][$row])) . ($row + $StartRowValueContent); //@ptr.nov - $StartRowValueContent -> mulai rows nilai data warna
                    if (($row + 1) % 2 == 1 and isset($each_sheet_content["oddCssClass"])) {
                        if ($each_sheet_content["oddCssClass"]["color"])
                            $current_sheet->getStyle($lineRange)->getFont()->getColor()->setARGB($each_sheet_content["oddCssClass"]["color"]);
                        //background
                        if ($each_sheet_content["oddCssClass"]["background"]) {
                            $current_sheet->getStyle($lineRange)->getFill()->getStartColor()->setRGB($each_sheet_content["oddCssClass"]["background"]);
                            $current_sheet->getStyle($lineRange)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                        }
                    } else if (($row + 1) % 2 == 0 and isset($each_sheet_content["evenCssClass"])) {
//                        echo "even",$row,"<BR>";
                        if ($each_sheet_content["evenCssClass"]["color"])
                            $current_sheet->getStyle($lineRange)->getFont()->getColor()->setARGB($each_sheet_content["evenCssClass"]["color"]);
                        //background
                        if ($each_sheet_content["evenCssClass"]["background"]) {
                            $current_sheet->getStyle($lineRange)->getFill()->getStartColor()->setRGB($each_sheet_content["evenCssClass"]["background"]);
                            $current_sheet->getStyle($lineRange)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                        }
                    }
                    //write content
                    for ($l = 0; $l < count($each_sheet_content['ceils'][$row]); $l++) {
                        $current_sheet->setCellValueByColumnAndRow($l, $row + $StartRowValueContent, $each_sheet_content['ceils'][$row][$l]); //@ptr.nov - $StartRowValueContent -> mulai rows nilai data 
                    }
                }
            }
        }
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		
		$file_name = self::getFolder(). $excel_file. '.xlsx';
			
			/* 
			$file_name = $tempDir .
			//            yii::$app->user->identity->username . '-' .  //hide this line , you can turn on by yourself.
            ($this->prefixStr ? $this->prefixStr . '-' : '') .
            str_replace(array('/', '*', '?', '\\', ':', '[', ']'), array('_', '_', '_', '_', '_', '_', '_'), $excel_file) .
            ($this->suffixStr ? '-' . $this->suffixStr : '-' . date('Ymd-His')) .
			//            '-' . date('Ymd-His').
            '.xlsx'; 
			*/

        $objWriter->save($file_name);
        return $file_name;
//        if (!file_exists($file_name)) {
//            self::setMsg("Error", "File not exist");
//            return false;
//        }
    }


    /**
     * define some class for header/even/odd row's style
     *
     * @param string $code
     * @return array
     */
    public
    static function getCssClass($code = '')
    {
        $cssClass =
            array(
                'red' => array('color' => 'FFFFFF', 'background' => 'FF0000'),
                'pink' => array('color' => '', 'background' => 'FFCCCC'),
                'green' => array('color' => '', 'background' => 'CCFF99'),
                'lightgreen' => array('color' => '', 'background' => 'CCFFCC'),
                'yellow' => array('color' => '', 'background' => 'FFFF99'),
                'white' => array('color' => '', 'background' => 'FFFFFF'),
                'grey' => array('color' => '000000', 'background' => '808080'),
                'greywhite' => array('color' => 'FFFFFF', 'background' => '808080'),
                'blue' => array('color' => 'FFFFFF', 'background' => 'blue'),
                'blueblack' => array('color' => '000000', 'background' => 'blue'),
                'lightblue' => array('color' => 'FFFFFF', 'background' => '6666FF'),
                'notice' => array('color' => '514721', 'background' => 'FFF6BF'),
                'header' => array('color' => 'FFFFFF', 'background' => '519CC6'),
                'headerblack' => array('color' => '000000', 'background' => '519CC6'),
                'odd' => array('color' => '', 'background' => 'E5F1F4'),
                'even' => array('color' => '', 'background' => 'F8F8F8'),
            );

        if (empty($code)) return $cssClass;
        elseif (isset($cssClass[$code])) return $cssClass[$code];
        else return [];
    }


    /**
     * Will invoke DownloadAction
     *
     * @param $excel_content
     * @param $excel_file
     * @param array $excel_props
     * @return bool
     */
    public function export2excel($excel_content, $excel_file,$pilih
        , $excel_props = array('creator' => 'WWSP Tool'
        , 'title' => 'WWSP_Tracking EXPORT EXCEL'
        , 'subject' => 'WWSP_Tracking EXPORT EXCEL'
        , 'desc' => 'WWSP_Tracking EXPORT EXCEL'
        , 'keywords' => 'WWSP Tool Generated Excel, Author: Scott Huang'
        , 'category' => 'WWSP_Tracking EXPORT EXCEL'))
    {
        if (!is_array($excel_content)) {
            return FALSE;
        }
        if (empty($excel_file)) {
            return FALSE;
        }
        $excelName = self::save2Excel($excel_content, $excel_file, $excel_props);
		/*
		 * Penambahan parameter variable $pilih
		 * 1= untuk View download
		 * 0= untuk console command line linux
		 */
		if ($pilih==1){		
			/*
			 * Link View download 
			 * 
			*/
			if ($excelName) {
				return $this->owner->redirect([Url::to('download'), "file_name" => 'temp/' . basename($excelName)
					, "file_type" => 'excel'
					, 'deleteAfterDownload' => true]);
			}
		}else{		
			$deleteAfterDownload=true;
			$file_name='temp/' . basename($excelName);
			if (empty($file_name)) {
				//return $this->goBack();
				return 0;
			}
			//$baseRoot=Yii::getAlias('@vendor').'/ptrnov/yii2-postman4excel/';
			$baseRoot=self::getPath();
			//$baseRoot=dirname(dirname(__DIR__)).'/cronjob/';
			//$baseRoot = Yii::getAlias('@webroot') . "/uploads/";
			$file_name = $baseRoot . $file_name;
			//echo $file_name,"<BR/>";
			if (!file_exists($file_name)) {
				//HzlUtil::setMsg("Error", "File not exist");
				return 0;
			}
			$fp = fopen($file_name, "r");
			$file_size = filesize($file_name);
			if ($file_type == 'excel') {
				header('Pragma: public');
				header('Expires: 0');
				header('Content-Encoding: none');
				header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
				header('Cache-Control: public');
				header('Content-Type: application/vnd.ms-excel');
				header('Content-Description: File Transfer');
				Header("Content-Disposition: attachment; filename=" . basename($file_name));
				header('Content-Transfer-Encoding: binary');
				Header("Content-Length:" . $file_size);
			} else if ($file_type == 'picture') { //pictures
				Header("Content-Type:image/jpeg");
				Header("Accept-Ranges: bytes");
				Header("Content-Disposition: attachment; filename=" . basename($file_name));
				Header("Content-Length:" . $file_size);
			} else { //other files
				Header("Content-type: application/octet-stream");
				Header("Accept-Ranges: bytes");
				Header("Content-Disposition: attachment; filename=" . basename($file_name));
				Header("Content-Length:" . $file_size);
			}

			$buffer = 1024;
			$file_count = 0;

			while (!feof($fp) && $file_count < $file_size) {
				$file_con = fread($fp, $buffer);
				$file_count+=$buffer;
				echo $file_con;
			}
			//echo fread($fp, $file_size);
			fclose($fp);
			if ($deleteAfterDownload) {
				
				//unlink($file_name);
			}
			return 1;
		
		}
    }

}
