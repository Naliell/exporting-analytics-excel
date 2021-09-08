<?php
namespace Vjweb;

require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_before.php");
require_once($_SERVER["DOCUMENT_ROOT"].'/local/vendor/autoload.php');

use \Bitrix\Main\Loader;
use \Bitrix\Main\Type\DateTime;

Loader::includeModule('vjweb.manage');

Class YandexMetric
{
    private const MODULE_ID = 'vjweb.manage';

    public function __construct($firstDate, $secDate)
    {
        $this->token = \COption::GetOptionString(self::MODULE_ID, "YANDEX_TOKEN", 0);
        $this->ids = array(
            '0' => \COption::GetOptionString(self::MODULE_ID, "YANDEX_COUNTER_ID", 0),
        );
        $time = new DateTime($secDate);
        $this->date_from = new DateTime($firstDate);
        $this->date_before = $time->add('+1439 minutes');



    }
    public function getDataTable()
    {
        $dates = $this->getDate($this->date_from->add('-1 day'), $this->date_before);
        $res = array();
        foreach ($dates as $date)
        {
            $date = new DateTime($date);

            $resApi = json_decode($this->requestAPI($date));

            $res[$date->format('d.m.Y')]['visits'] = $resApi->totals[0];
            $res[$date->format('d.m.Y')]['users'] = $resApi->totals[1];
        }

        return $res;
    }
    private function getDate($startTime, $endTime) {
        $day = 86400;
        $format = 'd.m.Y';
        $startTime = strtotime($startTime);
        $endTime = strtotime($endTime);
        $numDays = round(($endTime - $startTime) / $day);

        $days = array();

        for ($i = 1; $i < $numDays; $i++) {
            $days[] = date($format, ($startTime + ($i * $day)));
        }

        return $days;
    }
    private function requestAPI($date)
    {

        $data = array(
            'metrics' => 'ym:s:visits,ym:s:users',
            'dimensions' => 'ym:s:referer,ym:s:startURLDomain',
            'date1' => $date->format('Y-m-d'),
            'date2' => $date->format('Y-m-d'),
            'limit' => 10000,
            'offset' => 1,
            'ids' => $this->ids[0],
            'oauth_token' => $this->token,
            'pretty' => true,
        );
        $parameters = http_build_query($data);
        $headers = [
            'Content-type: application/x-yametrika+json',
            "Authorization:OAuth".$this->token,
        ];
        $url = 'https://api-metrika.yandex.net/stat/v1/data?'.$parameters;
        $curl = curl_init();
        curl_setopt($curl, CURLOPT_HTTPHEADER, $headers);
        curl_setopt($curl, CURLOPT_GET, true);
        curl_setopt($curl, CURLOPT_FOLLOWLOCATION, true);
        curl_setopt($curl, CURLOPT_URL, $url);
        curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($curl, CURLINFO_HEADER_OUT, true);

        $res_gen = curl_exec($curl);
        curl_close($curl);
        return $res_gen;
    }
}
