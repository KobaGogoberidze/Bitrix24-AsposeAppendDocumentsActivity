<?php

if (!defined("B_PROLOG_INCLUDED") || B_PROLOG_INCLUDED !== true) {
    die();
}

use Bitrix\Disk\Driver;
use Aspose\Words\WordsApi;
use Bitrix\Disk\SystemUser;
use Bitrix\Bizproc\FieldType;
use Bitrix\Main\Localization\Loc;
use Bitrix\Disk\Internals\FolderTable;
use Bitrix\Bizproc\Activity\PropertiesDialog;
use Aspose\Words\Model\Requests\{AppendDocumentOnlineRequest};
use Aspose\Words\Model\{DocumentEntry, DocumentEntryList, FileReference};

class CBPAsposeAppendDocumentsActivity extends CBPActivity
{

    /**
     * Initialize activity
     * 
     * @param string $name
     */
    public function __construct($name)
    {
        parent::__construct($name);

        $this->arProperties = [
            "Title" => "",
            "OriginalDocument" => null,
            "DocumentsToAppend" => null,
            "MergedDocumentId" => null
        ];

        $this->SetPropertiesTypes([
            "MergedDocumentId" => [
                "Type" => FieldType::INT
            ]
        ]);
    }

    /** 
     * Load modules
     * 
     * @return bool
     */
    protected function loadModules()
    {
        CModule::IncludeModule("disk");
    }

    /**
     * Start the execution of activity
     * 
     * @return CBPActivityExecutionStatus
     */
    public function Execute()
    {
        $validationErrors = self::ValidateProperties(array_map(
            fn ($property) => $this->{$property["FieldName"]},
            self::getPropertiesDialogMap()
        ));

        if (!empty($validationErrors)) {
            foreach ($validationErrors as $error) {
                $this->WriteToTrackingService($error["message"], 0, CBPTrackingType::Error);
            }
            return CBPActivityExecutionStatus::Closed;
        }

        try {
            $this->loadModules();
            $appSid = "29d5a0ea-8b04-407f-aaba-90238ce4735d";
            $appKey = "146091c59cd93e1cc40b5716572edb03";
            $wordsApi = new WordsApi($appSid, $appKey);

            $originalDocument = $this->prepareRawDocument($this->OriginalDocument);
            $documentsToAppend = $this->prepareDocumentEntryList($this->DocumentsToAppend);
            $appendRequest = new AppendDocumentOnlineRequest($originalDocument["tmp_name"], $documentsToAppend);
            $appendResponse = $wordsApi->appendDocumentOnline($appendRequest);

            $mergedDocument = array_pop($appendResponse->getDocument());
            if ($mergedDocument) {
                rename($mergedDocument->getPathName(), "{$mergedDocument->getPathName()}.docx");
                $this->WriteToTrackingService("{$mergedDocument->getPathName()}.docx", 0, CBPTrackingType::Error);

                $storage = Driver::getInstance()->getStorageByCommonId('administrative_s1');

                if ($storage) {
                    $documentTemplatesFolder = $storage->getChild([
                        "=NAME" => "Document Templates",
                        "TYPE" => FolderTable::TYPE_FOLDER
                    ]);
                    $mergedDocumentDiskObject = $documentTemplatesFolder->uploadFile(
                        $this->prepareRawDocument("{$mergedDocument->getPathName()}.docx"),
                        [
                            "CREATED_BY" => SystemUser::SYSTEM_USER_ID
                        ]
                    );
                    if ($mergedDocumentDiskObject) {
                        $this->WriteToTrackingService("Documents uploaded", 0, CBPTrackingType::Error);

                        $this->MergedDocumentId = $mergedDocumentDiskObject->getId();
                    } else {
                        throw new Exception(Loc::getMessage("ASPOSE_AD_UNABLE_TO_UPLOAD_DOCUMENT"));
                    }
                } else {
                    throw new Exception(Loc::getMessage("ASPOSE_AD_UNABLE_TO_LOAD_STORAGE"));
                }
                unlink("{$mergedDocument->getPathName()}.docx");

                return CBPActivityExecutionStatus::Closed;
            } else {
                throw new Exception(Loc::getMessage("ASPOSE_AD_UNABLE_TO_APPEND_DOCUMENTS"));
            }
        } catch (\Exception $ex) {
            $this->WriteToTrackingService($ex->getMessage(), 0, CBPTrackingType::Error);
        }

        return CBPActivityExecutionStatus::Closed;
    }

    /** 
     * Prepare raw document
     * 
     * @param string $document
     * @return array
     * @throws \Exception
     */
    protected function prepareRawDocument(string $document)
    {
        $rawDocument = CFile::makeFileArray($document);

        if ($rawDocument) {
            if (strtolower($rawDocument["type"]) == "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
                return $rawDocument;
            } else {
                throw new \Exception(Loc::getMessage("ASPOSE_AC_UNSUPPORTED_DOCUMENT_TYPE", ["#FILE_EXT#" => $rawDocument["type"]]));
            }
        } else {
            throw new \Exception(Loc::getMessage("ASPOSE_AC_DOCUMENT_NOT_FOUND", ["#DOCUMENT_ID#" => $document]));
        }
    }

    /** 
     * Prepare document entry list
     * 
     * @param array $documents
     * @return DocumentEntryList
     * @throws \Exception
     */
    protected function prepareDocumentEntryList(array $documents)
    {
        $documentEntries = [];

        foreach ($documents as $document) {
            if ($rawDocument = $this->prepareRawDocument($document)) {
                $documentEntries[] = new DocumentEntry([
                    "file_reference" => FileReference::fromLocalFileContent($rawDocument["tmp_name"]),
                    "import_format_mode" => "KeepSourceFormatting",
                ]);
            }
        }

        return new DocumentEntryList([
            "document_entries" => $documentEntries
        ]);
    }

    /**
     * Generate setting form
     * 
     * @param array $documentType
     * @param string $activityName
     * @param array $workflowTemplate
     * @param array $workflowParameters
     * @param array $workflowVariables
     * @param array $currentValues
     * @param string $formName
     * @return string
     */
    public static function GetPropertiesDialog($documentType, $activityName, $workflowTemplate, $workflowParameters, $workflowVariables, $currentValues = null, $formName = "", $popupWindow = null, $siteId = "")
    {
        $dialog = new PropertiesDialog(__FILE__, [
            "documentType" => $documentType,
            "activityName" => $activityName,
            "workflowTemplate" => $workflowTemplate,
            "workflowParameters" => $workflowParameters,
            "workflowVariables" => $workflowVariables,
            "currentValues" => $currentValues,
            "formName" => $formName,
            "siteId" => $siteId
        ]);
        $dialog->setMap(static::getPropertiesDialogMap($documentType));

        return $dialog;
    }

    /**
     * Process form submition
     * 
     * @param array $documentType
     * @param string $activityName
     * @param array &$workflowTemplate
     * @param array &$workflowParameters
     * @param array &$workflowVariables
     * @param array &$currentValues
     * @param array &$errors
     * @return bool
     */
    public static function GetPropertiesDialogValues($documentType, $activityName, &$workflowTemplate, &$workflowParameters, &$workflowVariables, $currentValues, &$errors)
    {
        $documentService = CBPRuntime::GetRuntime(true)->getDocumentService();
        $dialog = new PropertiesDialog(__FILE__, [
            "documentType" => $documentType,
            "activityName" => $activityName,
            "workflowTemplate" => $workflowTemplate,
            "workflowParameters" => $workflowParameters,
            "workflowVariables" => $workflowVariables,
            "currentValues" => $currentValues,
        ]);

        $properties = [];
        foreach (static::getPropertiesDialogMap($documentType) as $propertyKey => $propertyAttributes) {
            $field = $documentService->getFieldTypeObject($dialog->getDocumentType(), $propertyAttributes);
            if (!$field) {
                continue;
            }

            $properties[$propertyKey] = $field->extractValue(
                ["Field" => $propertyAttributes["FieldName"]],
                $currentValues,
                $errors
            );
        }

        $errors = static::ValidateProperties($properties, new CBPWorkflowTemplateUser(CBPWorkflowTemplateUser::CurrentUser));

        if (count($errors) > 0) {
            return false;
        }

        $currentActivity = &CBPWorkflowTemplateLoader::FindActivityByName($workflowTemplate, $activityName);
        $currentActivity["Properties"] = $properties;

        return true;
    }

    /**
     * Validate user provided properties
     * 
     * @param array $testProperties
     * @param CBPWorkflowTemplateUser $user
     * @return array
     */
    public static function ValidateProperties($testProperties = [], CBPWorkflowTemplateUser $user = null)
    {
        $errors = [];

        foreach (static::getPropertiesDialogMap() as $propertyKey => $propertyAttributes) {
            if (CBPHelper::getBool($propertyAttributes['Required']) && CBPHelper::isEmptyValue($testProperties[$propertyKey])) {
                $errors[] = [
                    "code" => "emptyText",
                    "parameter" => $propertyKey,
                    "message" => Loc::getMessage("ASPOSE_AD_FIELD_NOT_SPECIFIED", ["#FIELD_NAME#" => $propertyAttributes["Name"]])
                ];
            }
        }

        return array_merge($errors, parent::ValidateProperties($testProperties, $user));
    }

    /**
     * User provided properties
     * 
     * @return array
     */
    protected static function getPropertiesDialogMap()
    {
        return [
            "OriginalDocument" => [
                "Name" => Loc::getMessage("ASPOSE_AD_ORIGINAL_DOCUMENT_FIELD"),
                "FieldName" => "OriginalDocument",
                "Type" => FieldType::FILE,
                "Required" => true
            ],
            "DocumentsToAppend" => [
                "Name" => Loc::getMessage("ASPOSE_AD_DOCUMENTS_TO_APPEND_FIELD"),
                "FieldName" => "DocumentsToAppend",
                "Type" => FieldType::FILE,
                "Required" => true,
                "Multiple" => true
            ],
        ];
    }
}
