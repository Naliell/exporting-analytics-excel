<?php
namespace Vjweb;

require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");
require_once($_SERVER["DOCUMENT_ROOT"].'/local/vendor/autoload.php');

use \Bitrix\Main\Loader;
use \Bitrix\Highloadblock\HighloadBlockTable;
use \Bitrix\Main\Type\DateTime;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\{Font, Border, Alignment, FILL};
use \Bitrix\Main\Data\Cache;

\CModule::IncludeModule("statistic");
Loader::includeModule('vjweb.manage');

class ExportStatistic
{
    private const MODULE_ID = 'vjweb.manage';

    function __construct($firstDate, $secDate)
    {
        $time = new DateTime($secDate);
        $this->date_from = $firstDate;
        $this->date_before = $time->add('+1439 minutes');
    }
    public function exportStatistic()
    {
        $data = array();
        $data['ANALYTICS'] = $this->getAnalytics();
        $data['BASKET'] = $this->GetBaskets();
        $data['REGISTRATIONS'] = $this->getRegistrations();
        $data['REGISTRATIONS_ONCLICK'] = $this->getRegistrationsOncklick();
        $data['SALES'] = $this->getSales();
        $data['SALES_PRICE'] = $this->getSalesPrice();
        $data['ORDER_ONCLICK'] = $this->getOrderOnlcik();
        $yandexMetric = new YandexMetric($this->date_from,$this->date_before);
        $yandexData = $yandexMetric->getDataTable();
        $data['YANDEX_METRIC'] = $yandexData;
        $file = array('0' => $this->createXlFile($data));
        $this->sendEmail($file[0]);
        if (!empty($file[0]))
        {
            return $file[0];
        }
        else
        {
            return 'error';
        }

    }
    private function sendEmail($file)
    {
        $mail = \COption::GetOptionString(self::MODULE_ID, "EXPORT_MAIL1", 0).','.
                \COption::GetOptionString(self::MODULE_ID, "EXPORT_MAIL2", 0).','.
                \COption::GetOptionString(self::MODULE_ID, "EXPORT_MAIL3", 0).','.
                \COption::GetOptionString(self::MODULE_ID, "EXPORT_MAIL4", 0).','.
                \COption::GetOptionString(self::MODULE_ID, "EXPORT_MAIL5", 0).',';
        $to      = $mail;
        $subject = 'Статистика';
        $message = '';
        return $this->sendMailAttachment($to, $subject, $message, $file);

    }

    private function sendMailAttachment($mailTo, $subject, $message, $file)
    {
        $separator = "---";

        $headers = "MIME-Version: 1.0\r\n";

        $headers .= "Content-Type: multipart/mixed; boundary=\"$separator\""; // в заголовке указываем разделитель

        if($file){
            $bodyMail = "--$separator\n";
            $bodyMail .= "Content-type: text/html; charset='utf-8'\n";
            $bodyMail .= "Content-Transfer-Encoding: quoted-printable";
            $bodyMail .= "Content-Disposition: attachment; filename==?utf-8?B?".base64_encode(basename($file))."?=\n\n";
            $bodyMail .= $message."\n";
            $bodyMail .= "--$separator\n";
            $fileRead = fopen($file, "r");
            $contentFile = fread($fileRead, filesize($file));
            fclose($fileRead);
            $bodyMail .= "Content-Type: application/octet-stream; name==?utf-8?B?".base64_encode(basename($file))."?=\n";
            $bodyMail .= "Content-Transfer-Encoding: base64\n";
            $bodyMail .= "Content-Disposition: attachment; filename==?utf-8?B?".base64_encode(basename($file))."?=\n\n";
            $bodyMail .= chunk_split(base64_encode($contentFile))."\n";
            $bodyMail .= "--".$separator ."--\n";
        }else{
            $bodyMail = $message;
        }
        return mail($mailTo, $subject, $bodyMail,$headers);
    }
    private function getAnalytics()
    {
        $time = new DateTime();
        $time_basket = $time->add('+300 minutes');

            $arFilter = array(
            "DATE1" => $this->date_from,
            "DATE2" => $this->date_before
            );

        $rsDays = \CTraffic::GetDailyList(
            ($by="s_date"),
            ($order="desc"),
            $arMaxMin,
            $arFilter,
            $is_filtered
            );

        while ($arDay = $rsDays->Fetch())
        {
            $analytics[$arDay['ID']]['DATE_STAT'] = $arDay["DATE_STAT"];
            $analytics[$arDay['ID']]['SESSIONS'] = $arDay["SESSIONS"];
            $analytics[$arDay['ID']]['GUESTS'] = $arDay["GUESTS"];
        }
        return $analytics;
    }

    private function GetBaskets()
    {
        $arBasketItems = array();

        $dbBasketItems = \CSaleBasket::GetList(
                array(

                        "ID" => "DESC"
                    ),
                array(
                        "><DATE_UPDATE" => array($this->date_from,$this->date_before),
                        "LID" => SITE_ID,
                        "ORDER_ID" => "NULL"
                    ),
                false,
                false,
                array()
            );

        while ($arItems = $dbBasketItems->Fetch())
        {
            if ($date_basket != $arItems['DATE_UPDATE']->format('d.m.Y')) {
                $abandoned_basket = 0;
                $abandoned_basket_price = 0;
            }
            $date_basket = $arItems['DATE_UPDATE']->format('d.m.Y');

                if (empty($fuser_id) || $fuser_id != $arItems['FUSER_ID']) {
                    $abandoned_basket++;
                    $basket[$date_basket]['COUNT'] = $abandoned_basket;
                }
                $abandoned_basket_price += $arItems['BASE_PRICE'];
                $basket[$date_basket]['BASE_PRICE'] = $abandoned_basket_price;
                $fuser_id = $arItems['FUSER_ID'];

        }
        return $basket;
    }

    private function getRegistrations()
    {
        $data = \CUser::GetList(($by="ID"), ($order="ASC"),
                    array(
                        "DATE_REGISTER_1" => $this->date_from,
                        "DATE_REGISTER_2" => $this->date_before,
                        'ACTIVE' => 'Y'
                    ),
                    array(
                        'SELECT' => array(
                            'UF_*'
                        )
                    ),

                );

        while($arUser = $data->Fetch()) {
            $date_register = new DateTime($arUser['DATE_REGISTER']);
            $register[$date_register->format('d.m.Y')]++;
        }
        return $register;
    }

    private function getRegistrationsOncklick()
    {
        $data = \CUser::GetList(($by="ID"), ($order="ASC"),
                    array(
                        "DATE_REGISTER_1" => $this->date_from,
                        "DATE_REGISTER_2" => $this->date_before,
                        'ACTIVE' => 'Y'
                    ),
                    array(
                        'SELECT' => array(
                            'UF_*'
                        )
                    ),

                );

        while($arUser = $data->Fetch()) {
            $date_register = new DateTime($arUser['DATE_REGISTER']);

            if ($arUser['UF_ONECLICK']) {
                $reg_onclick[$date_register->format('d.m.Y')]++;
            }
        }
        return $reg_onclick;
    }

    private function getSales()
    {
        $arFilter = Array(
           "><DATE_UPDATE" => array($this->date_from,$this->date_before)
           );

        $db_sales = \CSaleOrder::GetList(array("DATE_INSERT" => "ASC"), $arFilter);

        while ($ar_sales = $db_sales->Fetch())
        {
            $date_sale = new DateTime($ar_sales['DATE_UPDATE']);
            $sales[$date_sale->format('d.m.Y')][] = $ar_sales;
        }
        return $sales;
    }

    private function getSalesPrice()
    {
        $arFilter = Array(
           "><DATE_UPDATE" => array($this->date_from,$this->date_before)
           );

        $db_sales = \CSaleOrder::GetList(array("DATE_INSERT" => "ASC"), $arFilter);

        while ($ar_sales = $db_sales->Fetch())
        {
            $date_sale = new DateTime($ar_sales['DATE_UPDATE']);
            $sales_price[$date_sale->format('d.m.Y')]['PRICE'] += $ar_sales['PRICE'];
        }
        return $sales_price;
    }

    private function getOrderOnlcik()
    {
        $arFilter = Array(
           "><DATE_UPDATE" => array($this->date_from,$this->date_before)
           );

        $db_sales = \CSaleOrder::GetList(array("DATE_INSERT" => "ASC"), $arFilter);

        while ($ar_sales = $db_sales->Fetch())
        {
            $date_sale = new DateTime($ar_sales['DATE_UPDATE']);
            $arOrderProps = \CSaleOrderProps::GetByID($ar_sales['ID']);
            $db_vals = \CSaleOrderPropsValue::GetList(
                    array(),
                    array(
                            "ORDER_ID" => $ar_sales['ID'],
                            "CODE" => 'ONCLICK'
                        )
                );
                $arVals = $db_vals->Fetch();
              if ($arVals['VALUE'] == 'Y'){
                  $order_onclick[$date_sale->format('d.m.Y')]++;
              } elseif ($ar_sales['COMMENTS'] == 'Заказ оформлен в 1 клик') {
                  $order_onclick[$date_sale->format('d.m.Y')]++;
              }
        }
        return $order_onclick;
    }

    private function saveGrafic()
    {
        $site_url = 'https://'.$_SERVER["HTTP_HOST"].'/bitrix/admin';
        $post_var = 'AUTH_FORM=Y&TYPE=AUTH&USER_LOGIN=yan555&USER_PASSWORD=siteYAN.123lkP&Login=&USER_REMEMBER=Y&captcha_sid=&captcha_word=bca897af2d782ba894633b851ak35ff3&sessid=bda897af5d782ba194643b851af353f3'; //эти данные собираются путём парсинга формы авторизации, для простоты поместил их в одну переменную

        $ch = curl_init();
        curl_setopt($ch, CURLOPT_COOKIEJAR, $_SERVER['DOCUMENT_ROOT'].'/cookie.txt');
        curl_setopt($ch, CURLOPT_COOKIEFILE, $_SERVER['DOCUMENT_ROOT'].'/cookie.txt');
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
        curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
        curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);
        curl_setopt($ch, CURLOPT_HEADER, false);
        curl_setopt($ch, CURLOPT_URL, $site_url);
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $post_var);

        $text = curl_exec($ch);

        $new_date = new DateTime($this->date_before);

        curl_setopt($ch, CURLOPT_URL, 'https://'.$_SERVER["HTTP_HOST"].'/bitrix/admin/event_graph.php?&find_events[]=8&find_events[]=9&find_events[]=14&find_date1='.$this->date_from.'&find_date2='.$new_date->format('d.m.Y').'&filter=Y&set_filter=Y&width=576&height=400&lang=ru');
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        $graf_event = curl_exec($ch);

        $ReadFile = fopen ($_SERVER['DOCUMENT_ROOT'].'/local/export_statistic/graf_event.png', "wb");
        $res = fwrite($ReadFile, $graf_event);
        curl_close($ch);
        if ($res){
            return true;
        } else {
            return 'не удалось загрузить файл';
        }
    }

    private function createXlFile($data){
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $spreadsheet->setActiveSheetIndex(0);
        // Шапка таблицы
        if($expload_date[0]<10)
          $expload_date[0] = '0'.$expload_date[0];
        $str = 3;
        $sheet->setTitle('Таблица');
        $sheet->setCellValue('A'.$str, 'Дата');
        $sheet->setCellValue('B'.$str, 'Посетители (По данным Яндкес метрики)');
        $sheet->setCellValue('C'.$str, 'Визиты (По данным Яндкес метрики)');
        $sheet->setCellValue('D'.$str, 'Регистрации (Количество регистраций)');
        $sheet->setCellValue('E'.$str, 'Корзины кол-во (Количество пользователей, которые добавили товар в корзину, но не оформили)');
        $sheet->setCellValue('F'.$str, 'Сумма корзин (Общая сумма всех товаров в корзинах - не оформленные заказы)');
        $sheet->setCellValue('G'.$str, 'Заказы кол-во (Количество совершенных заказов)');
        $sheet->setCellValue('H'.$str, 'Сумма заказов (Сумма совершенных заказов)');
        $sheet->setCellValue('I'.$str, 'Средний чек (Сумма заказов/количество заказов)');
        $sheet->setCellValue('J'.$str, 'В один клик (Регистрации в 1 клик)');
        $sheet->setCellValue('K'.$str, 'В один клик (Заказы в 1 клик)');
        // //Установка ширины столбцов
        $sheet->getColumnDimension('A')->setWidth(25);
        $sheet->getColumnDimension('B')->setWidth(25);
        $sheet->getColumnDimension('C')->setWidth(25);
        $sheet->getColumnDimension('D')->setWidth(25);
        $sheet->getColumnDimension('E')->setWidth(25);
        $sheet->getColumnDimension('F')->setWidth(25);
        $sheet->getColumnDimension('G')->setWidth(25);
        $sheet->getColumnDimension('H')->setWidth(25);
        $sheet->getColumnDimension('I')->setWidth(25);
        $sheet->getColumnDimension('J')->setWidth(25);
        $sheet->getColumnDimension('K')->setWidth(25);


        //Установка высоты для строки
        $sheet->getRowDimension($str)->setRowHeight(100);
        $style_header = array(
          'font' => array('bold' => true),
          'alignment' => array(
              'horizontal' => Alignment::HORIZONTAL_CENTER,
              'vertical' => Alignment::VERTICAL_CENTER,
              'wrapText' => true,
          ),
          'borders' => array(
            'allBorders' => array(
                'borderStyle' => Border::BORDER_THIN,
                'color' => array('rgb' => '808080')
            ),
          ),
          'fill' => array(
            'fillType' => Fill::FILL_SOLID,
            'startColor' => array('rgb' => 'F0FBFE')
          )
        );

        $sheet->getStyle('A'.$str)->applyFromArray($style_header);
        $sheet->getStyle('B'.$str)->applyFromArray($style_header);
        $sheet->getStyle('C'.$str)->applyFromArray($style_header);
        $sheet->getStyle('D'.$str)->applyFromArray($style_header);
        $sheet->getStyle('E'.$str)->applyFromArray($style_header);
        $sheet->getStyle('F'.$str)->applyFromArray($style_header);
        $sheet->getStyle('G'.$str)->applyFromArray($style_header);
        $sheet->getStyle('H'.$str)->applyFromArray($style_header);
        $sheet->getStyle('I'.$str)->applyFromArray($style_header);
        $sheet->getStyle('J'.$str)->applyFromArray($style_header);
        $sheet->getStyle('K'.$str)->applyFromArray($style_header);
        // Тело таблицы
        $count = 1;
        foreach($data['ANALYTICS'] as $key => $val){
          $str++;
          $sheet->setCellValue('A'.$str, $val['DATE_STAT']);
          $sheet->setCellValue('B'.$str, $data['YANDEX_METRIC'][$val['DATE_STAT']]['users']);
          $sheet->setCellValue('C'.$str, $data['YANDEX_METRIC'][$val['DATE_STAT']]['visits']);

          unset($time);
          foreach($user['schedule']['workhours'] as $shld){
              $time .= $shld.' ';
          }
          if ($data['SALES'][$val['DATE_STAT']] > 0 && $data['SALES_PRICE'][$val['DATE_STAT']]['PRICE'] > 0){
              $sred_chek= $data['SALES_PRICE'][$val['DATE_STAT']]['PRICE']/count($data['SALES'][$val['DATE_STAT']]);
          } else {
              $sred_chek = 0;
          }

          $sheet->setCellValue('D'.$str, $data['REGISTRATIONS'][$val['DATE_STAT']] ? round($data['REGISTRATIONS'][$val['DATE_STAT']],2) : '0');
          $sheet->setCellValue('E'.$str, $data['BASKET'][$val['DATE_STAT']]['COUNT'] ? round($data['BASKET'][$val['DATE_STAT']]['COUNT'],2) : '0');
          $sheet->setCellValue('F'.$str, $data['BASKET'][$val['DATE_STAT']]['BASE_PRICE'] ? round($data['BASKET'][$val['DATE_STAT']]['BASE_PRICE'],2) : '0');
          $sheet->setCellValue('G'.$str, count($data['SALES'][$val['DATE_STAT']]) ? count($data['SALES'][$val['DATE_STAT']]) : '0');
          $sheet->setCellValue('H'.$str, $data['SALES_PRICE'][$val['DATE_STAT']]['PRICE'] ? round($data['SALES_PRICE'][$val['DATE_STAT']]['PRICE'],2): '0');
          $sheet->setCellValue('I'.$str, round($sred_chek,2));
          $sheet->setCellValue('J'.$str, $data['REGISTRATIONS_ONCLICK'][$val['DATE_STAT']] ? round($data['REGISTRATIONS_ONCLICK'][$val['DATE_STAT']],2) : '0');
          $sheet->setCellValue('K'.$str, $data['ORDER_ONCLICK'][$val['DATE_STAT']] ? round($data['ORDER_ONCLICK'][$val['DATE_STAT']],2) : '0');

          //Установка высоты для строки
          $sheet->getRowDimension($str)->setRowHeight(20);



          $count++;
        }
        $spreadsheet->setActiveSheetIndex(0);
        $activeSheet = $spreadsheet->getActiveSheet();
        $str_end = $str + 1;
        $sheet->setCellValue('A'.$str_end, 'ИТОГО');
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue(
                'B'.$str_end,
                '=SUM(B3:B'.$str.')'
            );
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue(
                'C'.$str_end,
                '=SUM(C3:C'.$str.')'
            );
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue(
                'D'.$str_end,
                '=SUM(D3:D'.$str.')'
            );
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue(
                'E'.$str_end,
                '=SUM(E3:E'.$str.')'
            );
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue(
                'F'.$str_end,
                '=SUM(F3:F'.$str.')'
            );
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue(
                'G'.$str_end,
                '=SUM(G3:G'.$str.')'
            );
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue(
                'H'.$str_end,
                '=SUM(H3:H'.$str.')'
            );
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue(
                'I'.$str_end,
                '=SUM(I3:I'.$str.')'
            );

        // Подвал таблицы
        $str++;
        $spreadsheet->setActiveSheetIndex(0);
        $writer = new Xlsx($spreadsheet);
        //Сохраняем файл
        $writer->save(__DIR__.'../../../../export_statistic/export.xlsx');
        return $_SERVER['DOCUMENT_ROOT'].'/local/export_statistic/export.xlsx';
    }
}
