<?php

if (!defined("B_PROLOG_INCLUDED") || B_PROLOG_INCLUDED !== true) {
    die();
}

use Bitrix\Main\Localization\Loc;

$arActivityDescription = [
    "NAME" => Loc::getMessage("ASPOSE_AD_NAME"),
    "DESCRIPTION" => Loc::getMessage("ASPOSE_AD_DESCRIPTION"),
    "TYPE" => "activity",
    "CLASS" => "AsposeAppendDocumentsActivity",
    "JSCLASS" => "BizProcActivity",
    "CATEGORY" => [
        "OWN_ID" => "aspose",
        "OWN_NAME" => "Aspose",
    ]
];
