/**
 * @NApiVersion 2.0
 * @NScriptType MapReduceScript
 */
define(['N/search', 'N/file', 'N/record', 'N/log', 'N/sftp', 'N/task'],
    function (search, file, record, log, sftp, task) {

        function getInputData() {
            try {
                log.debug('Start Script -- getInput')
                var projectSearchObj = search.create({
                    type: "customrecord_cseg_bc_project",
                    filters:
                        [
                            ["isinactive","is","F"],
                            "AND",
                            ["custrecord_bc_sent_to_df","is","F"],
                            "AND",
                            ["custrecord_bc_proj_customer","noneof","601","600","598","603","609","608","610","596","604","602","597","595","599","606","605","607","2658"],
                            "AND", 
                            ["custrecord_bc_proj_subsidiary","anyof","15","14","16","13"]
                        ],
                    columns:
                        [
                            search.createColumn({name: "name", label: "ProjectName"}),
                            search.createColumn({name: "externalid", label: "ProjectXRefCode"}),
                            search.createColumn({name: "custrecord_cponum", label: "ProjectDescription"}),
                            search.createColumn({name: "custrecordbc_proj_contract_date", label: "StartDate"}),
                            search.createColumn({
                                name: "namenohierarchy",
                                join: "CUSTRECORD_BC_PROJ_SUBSIDIARY",
                                label: "CreationOrgXRefCode"
                            }),
                            search.createColumn({
                                name: "externalid",
                                join: "CUSTRECORD_BC_PROJ_CUSTOMER",
                                label: "ParentXRefCode"
                            }),
                            search.createColumn({name: "custrecord_bc_reporting_project", label: "LedgerCode"})
                        ]
                });
                return projectSearchObj;
            }catch (e) {
                log.debug('Get Input Error', e);
            }
        }

        function map(context) {
            try {
                log.debug('Start Script -- map')
                var result = JSON.parse(context.value);
                //log.debug('result', result)
                var values = result.values || {};

                var id = result.id;
                var name = values.name || '';
                var projectField = values.externalid;
                var projectExId = projectField ? String(projectField.value || projectField) : '';
                log.debug('projectExId', projectExId);

                var poNumber = values.custrecord_cponum || '';
                var rawDate = values.custrecordbc_proj_contract_date;
                var contractDate;
                if (rawDate) {
                    var parts = String(rawDate).split('/');
                    var mm = ('0' + parts[0]).slice(-2);
                    var dd = ('0' + parts[1]).slice(-2);
                    var yyyy = parts[2];
                    contractDate = yyyy + '-' + mm + '-' + dd;
                } else {
                    contractDate = '2000-01-01';
                }
                var subsidiaryName = values['namenohierarchy.CUSTRECORD_BC_PROJ_SUBSIDIARY'] || '';
                //log.debug('subsidiaryName', subsidiaryName)

                var parentField = values['externalid.CUSTRECORD_BC_PROJ_CUSTOMER'];
                var parentExternalId = parentField ? String(parentField.value || parentField) : '';
                log.debug('parentExternalId', parentExternalId);

                var ledgerCode = ''
                var ledgerCodeField = values['custrecord_bc_reporting_project'];
                if (ledgerCodeField && typeof ledgerCodeField === 'object' && ledgerCodeField.value) {
                    ledgerCode = ledgerCodeField.text;
                }
                log.debug('ledgerCode', ledgerCode);

                var row = [
                    'D',                //Default
                    'Project',          //Project
                    'Project',          //Project
                    projectExId,        //XRefCode
                    name,               //Name
                    poNumber,           //Description
                    subsidiaryName,     //CreationOrgXRefCode
                    parentExternalId,   //ParentXRefCode
                    '0',                //BudgetedHours
                    '0',                //BudgetedAmount
                    contractDate,       //StartDate
                    '',                 //DueDate
                    '',                 //CompletedDate
                    '0',                //ProjectPriority
                    '0',                //PercentComplete
                    '',                 //ProductGroupXRefCode
                    '',                 //ProductModuleXRefCode
                    '',                 //ProjectTypeXRefCode
                    '',                 //ProjectPhaseXRefCode
                    '',                 //AccountNum
                    '0',                //IFRSClassification
                    '',                 //ClockTransferCode
                    ledgerCode,         //LedgerCode
                    '',                 //TaxationAddressXRefCode
                ];

                log.debug('row', row)
                context.write({
                    key: 'all',
                    value: JSON.stringify({row: row,id: result.id})
                });
            }catch (e) {
                log.debug('Error in map', e)
            }
        }

        function reduce(context) {
            log.debug('Start Script -- reduce')
            try {
                var values = context.values;
                log.debug('values', values);

                var firstRow = ['ProjectImport','','','','','','','','','','','','','','','','','','','','','','',''];
                var headers = [
                    'H',
                    'Project',
                    'Project',
                    'XRefCode',
                    'Name',
                    'Description',
                    'CreationOrgXRefCode',
                    'ParentXRefCode',
                    'BudgetHours',
                    'BudgetAmount',
                    'StartDate',
                    'DueDate',
                    'CompletedDate',
                    'ProjectPriority',
                    'PercentComplete',
                    'ProductGroupXRefCode',
                    'ProductModuleXRefCode',
                    'ProjectTypeXRefCode',
                    'ProjectPhaseXRefCode',
                    'AccountNum',
                    'IFRSClassification',
                    'ClockTransferCode',
                    'LedgerCode',
                    'TaxationAddressXRefCode'
                ];

                var rows = [];
                var projectIds = [];

                for (var i = 0; i < values.length; i++) {
                    var parsed = JSON.parse(values[i]);
                    var rowData = parsed.row;
                    log.debug('rowData', rowData)
                    projectIds.push(parsed.id);

                    // Clean XRefCode (keeps underscores, removes commas/periods)
                    var xrefCode = (rowData[3] || '')
                        .replace(/&/g, 'AND')         // & → AND
                        .replace(/-/g, ' ')           // dashes → space
                        .replace(/[.,]/g, '')         // remove commas & periods
                        .replace(/\s+/g, '_')         // spaces → underscore
                        .replace(/[^A-Za-z0-9_]/g, '')// strip anything not alphanumeric or underscore
                        .toUpperCase();

                    var projectName = (rowData[4] || '')
                        .replace(/&/g, 'and')
                        .replace(/['.,]/g, '')
                        .replace(/-/g, ' ')
                        .replace(/[^0-9A-Za-z ]/g, '')
                        .replace(/\s+/g, ' ')
                        .trim();

// Clean Description (rowData[5]) — THIS is the one that breaks CSV if unquoted
                    var descriptionClean = (rowData[5] || '')
                        .replace(/&/g, 'and')
                        .replace(/['.,]/g, '')       // <-- removes commas and periods
                        .replace(/-/g, ' ')
                        .replace(/[^0-9A-Za-z ]/g, '')
                        .replace(/\s+/g, ' ')
                        .trim();

                    var parentXrefCode = (rowData[7] || '')
                        .replace(/&/g, 'AND')
                        .replace(/-/g, ' ')
                        .replace(/[.,]/g, '')
                        .replace(/\s+/g, '_')
                        .replace(/[^A-Za-z0-9_]/g, '')
                        .toUpperCase();

                    var subsidiaryName = (rowData[6] || '')
                        .replace(/&/g, 'AND')
                        .replace(/-/g, ' ')
                        .replace(/\s+/g, '_')
                        .replace(/[^\w]/g, '')
                        .toUpperCase();

                    var projectName = (rowData[4] || '')
                        .replace(/&/g, 'and')
                        .replace(/['.,]/g, '')
                        .replace(/-/g, ' ')
                        .replace(/[^0-9A-Za-z ]/g, '')
                        .replace(/\s+/g, ' ')
                        .trim();

                    var fullRow = [
                        rowData[0],
                        rowData[1],
                        rowData[2],
                        xrefCode || '',
                        projectName || '',
                        descriptionClean || '',
                        subsidiaryName || '',
                        parentXrefCode || '',
                        rowData[8] || '',
                        rowData[9] || '',
                        rowData[10] || '',
                        rowData[11] || '',
                        rowData[12] || '',
                        rowData[13] || '',
                        rowData[14] || '',
                        rowData[15] || '',
                        rowData[16] || '',
                        rowData[17] || '',
                        rowData[18] || '',
                        rowData[19] || '',
                        rowData[20] || '',
                        rowData[21] || '',
                        rowData[22] || '',
                        rowData[23] || ''
                    ];
                    rows.push(fullRow);

                }

                // Check if there are no rows to export
                if (rows.length == 0) {
                    log.debug('No Data Found', 'Skipping file creation and SFTP upload because there are no customers to export.');
                    return; // Exit reduce early
                }
                //log.debug('rows', rows)
                //log.debug('project ids', projectIds)

                //Get today's date and timestamp for filename
                var now = new Date();
                var timeStamp = ('0' + now.getHours()).slice(-2) +
                    ('0' + now.getMinutes()).slice(-2) +
                    ('0' + now.getSeconds()).slice(-2);
                var fileName = 'BlueCollar_ProjectList_' + getTodayString() + '_' + timeStamp + '.csv.ready';
                //var fileName = 'BlueCollar_ProjectList_' + getTodayString() + '_' + timeStamp + '.csv';

                //Create Export Log custom record
                var exportLog = record.create({
                    type: 'customrecord_bc_df_export_log',
                    isDynamic: true
                });
                exportLog.setValue({ fieldId: 'name', value: fileName });
                exportLog.setValue({ fieldId: 'custrecord_bc_df_export_type', value: 2 }); //Project
                exportLog.setValue({ fieldId: 'custrecord_bc_df_export_date', value: new Date() });
                exportLog.setValue({ fieldId: 'custrecord_bc_df_export_status', value: 1 }); //Pending
                exportLog.setValue({ fieldId: 'custrecord_bc_df_export_record_ids', value: projectIds.join(', ') });
                var exportLogId = exportLog.save();
                log.debug('exportLogId', exportLogId)

                var fileId = null;

                //Start CSV creation
                try {
                    //var csvContent = firstRow.join(',') + '\n' + headers.join(',') + '\n' + rows.join('\n');

                    var esc = function (s) {
                        var t = (s == null ? '' : String(s));
                        // Standard CSV escaping only: quote if it contains quotes, commas, or newlines
                        if (/[",\n]/.test(t)) {
                            t = '"' + t.replace(/"/g, '""') + '"';
                        }
                        return t;
                    };

                    var csvFirst = firstRow.map(esc).join(',');
                    var csvHead  = headers.map(esc).join(',');
                    var csvRows  = rows.map(function(r){
                        return r.map(esc).join(',');
                    });

                    var csvContent = csvFirst + '\n' + csvHead + '\n' + csvRows.join('\n');

                    var fileObj = file.create({
                        name: fileName,
                        fileType: file.Type.CSV,
                        contents: csvContent,
                        folder: 1670
                    });
                    var fileId = fileObj.save();
                    log.debug('CSV file saved', fileName + ' (ID: ' + fileId + ')');

                    //PRODUCTION ENVIRONMENT
                    var sftpConnection = sftp.createConnection({
                        username: 'c2ogroup',
                        keyId: 'custkey1',
                        url: 'fts01.dayforcehcm.com',
                        port: 22,
                        directory: '/',
                        hostKey: 'AAAAB3NzaC1yc2EAAAADAQABAAABgQClPF7ps1px0k7dAf5eaaYRymvKFcn3/JNRA6dvC+pC1K2SQ0YkP22nsY/BEGrQf7Q+wmGw2gVXhKmLmum63qOn/7b2xdtS0oKEOdhbL3pl18O/GtHPbPG21SGMOBbr+4MzRmFfypnMPUNPRaDTANhPhJXO0CMbJb+ho3ME2kFben4DOe+gg8ZXhlY+kUagc/hEOfEFX9vY+jZgyXs2lhBIpnrZroa7fuJf43dCNkr1k2lHurvOv7ND/EhUeCnOrpTJiVHzs8jwS0GdYzNRDd3wKU6T/chVhNm5svdPuVgeVRXrvrF2oQpecBVCdW15iBDV99mBy8mEoHpG5mvLH3bhdDDpH6v30Wj80iHbBKmQ1GdZFwbHMtsyhdyJw0a7fuDCPW2Z8inRDHAyb26OrvUMMYm71plw/s7gqtbc00XtM0zLuCRkYhhEcxY7VkGUTI9mDqBgb8K+4mjLtGbH3KPwN9lR7vmPgQSmrZWAaz16/zwXf4kBbpi8cF3pCg2Q3eE=',
                    });
                    log.debug('SFTP Connection', 'Successfully connected.');

                    //STAGE ENVIRONMENT
                    /*var sftpConnection = sftp.createConnection({
                        username: 'c2ogroupstage',
                        keyId: 'custkey6',
                        url: 'ftstest01.dayforcehcm.com',
                        port: 22,
                        directory: '/',
                        hostKey: 'AAAAB3NzaC1yc2EAAAADAQABAAABgQDx5b1w5xcUeEtwwyURMn9zkeoDQZP8JPcD5fGGEOE9jG0w6wVzjyl9eeT+7s+b6Im4ipteeO2A4ErqTmC4IYiDlbehvAzBdIg3VV/II8KneLP25J9+/wvEHCmik55aXY4wphS9nOfoVzwV04vBTQy01kBknTCii+CLXHcbKTs7vFvPbzDh4Sn0eOO3PK9rbtAsglzd7rJdHm7BEny8TEceaja/3j4KKgKcfbfqLd6+EMTmRv361/lOmFqBmfbGbCKzg6TJZxD3mnGeAA8PhSRl6smzi/Mb7UmNVfYpSvjfy0IWC/Te1K26Aj/Mo9Q+9K173ytoBpTUsl0US2HtDk5ZsG24syPEczkHW1uMIgxscBVRBdqhaqeyVH95nWc6g7XICrAv7W7eZ80Zndke4IOfxe1uTVGoJ8CCIS7S1CsWMcpKXsAtA9JWrzmfgDzgd0ghcCPkADFXSGisUN6ji64ZWlb8yy3OOOGsqCIT0dM2yRojCLD2cVcSz3Zxj/uhzWk=',

                        // Temporarily omit hostKey to test first
                    });
                    log.debug('SFTP Connection', 'Successfully connected.');*/
                    log.debug('sftpConnection', sftpConnection)

                    sftpConnection.upload({
                        directory: '/Import/ProjectImport',
                        filename: fileObj.name,
                        file: fileObj,
                        //replaceExisting: true
                    });

                    /*sftpConnection.upload({
                        directory: '/Import/ProjectImport',
                        filename: fileObj.name,
                        file: fileObj,
                        //replaceExisting: true
                    });*/

                    var files = sftpConnection.list({path: '/Import/ProjectImport'});
                    log.debug('Directory Listing', JSON.stringify(files));
                    log.debug('sftpConnection', sftpConnection)

                    for (var j = 0; j < projectIds.length; j++) {
                        record.submitFields({
                            type: 'customrecord_cseg_bc_project',
                            id: projectIds[j],
                            values: {
                                'custrecord_bc_sent_to_df': true,
                            }
                        });
                    }

                    record.submitFields({
                        type: 'customrecord_bc_df_export_log',
                        id: exportLogId,
                        values: {
                            custrecord_bc_df_export_file: fileId,
                            custrecord_bc_df_export_status: 2,
                            custrecord_bc_df_export_logs: 'File exported and uploaded successfully.'
                        }
                    });

                }catch (e) {
                    var logValues = {
                        custrecord_bc_df_export_status: 3,
                        custrecord_bc_df_export_logs: 'Export failed: ' + e.message
                    };
                    if (fileId) {
                        logValues.custrecord_bc_df_export_file = fileId;
                    }
                    record.submitFields({
                        type: 'customrecord_bc_df_export_log',
                        id: exportLogId,
                        values: logValues
                    });
                    log.debug('Export Failed', e);
                }

                /*var mrTask = task.create({
                    taskType: task.TaskType.MAP_REDUCE,
                    scriptId: 'customscript_bc_df_mr_export_costcode',
                    deploymentId: 'customdeploy_bc_df_mr_export_costcode',
                    params: {
                        custscript_bc_df_project_ids: JSON.stringify(projectIds)
                    }
                });
                mrTask.submit();
                log.debug('Triggered Cost Code MR', 'Project IDs passed: ' + JSON.stringify(projectIds));*/

            } catch (e) {
                log.debug('Error in reduce', e)
            }
        }

        function summarize(summary) {

        }

        function getTodayString() {
            var today = new Date();
            var yyyy = today.getFullYear();
            var mm = ('0' + (today.getMonth() + 1)).slice(-2);
            var dd = ('0' + today.getDate()).slice(-2);
            return yyyy + mm + dd;
        }

        return {
            getInputData: getInputData,
            map: map,
            reduce: reduce,
            //summarize: summarize
        };
    });