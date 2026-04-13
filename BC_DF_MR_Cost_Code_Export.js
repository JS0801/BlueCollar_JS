/**
 * @NApiVersion 2.0
 * @NScriptType MapReduceScript
 */
define(['N/search', 'N/file', 'N/record', 'N/log', 'N/sftp', 'N/runtime', 'N/task'],
    function (search, file, record, log, sftp, runtime, task) {

        function getInputData() {
            try {
                log.debug('Start Script -- getInput')
                /*var currentScript = runtime.getCurrentScript();
                var projectIdParam = currentScript.getParameter({name: 'custscript_bc_df_project_ids'});
                var projectIds = projectIdParam ? JSON.parse(projectIdParam) : [];
                log.debug('Project IDs', projectIds);
                var filters = [
                    ["isinactive", "is", "F"],
                    "AND",
                    ["custrecord_bc_costcode_sent_to_df", "is", "F"],
                ];
                if (projectIds && projectIds.length > 0) {
                    filters.push("AND");
                    filters.push(["custrecord_bc_budget_project", "anyof"].concat(projectIds));
                }*/

                var filters = [
                    ["isinactive", "is", "F"],
                    "AND",
                    ["custrecord_bc_costcode_sent_to_df", "is", "F"],
                    "AND",
                    ["custrecord_bc_budget_project.custrecord_bc_sent_to_df", "is", "T"],
                    //"AND",
                    //["custrecord_bc_budget_project","anyof","1611"]
                    "AND",
                    ["custrecord_bc_budget_project.custrecord_bc_proj_subsidiary","anyof","15","14","16"]

                ];

                var costCodeSearchObj = search.create({
                    type: "customrecord_bc_budget_item",
                    filters: filters,
                    columns: [
                        search.createColumn({
                            name: "externalid",
                            summary: "GROUP",
                            label: "External ID"
                        }),
                        search.createColumn({
                            name: "custrecord_bc_budget_code",
                            summary: "GROUP",
                            label: "Cost Code",
                            sort: search.Sort.ASC
                        }),
                        search.createColumn({
                            name: "internalid",
                            join: "CUSTRECORD_BC_BUDGET_CODE",
                            summary: "GROUP",
                            label: "Cost Code ID"
                        }),
                        search.createColumn({
                            name: "custrecord_bc_budget_project",
                            summary: "GROUP",
                            label: "Project Name"
                        }),
                        search.createColumn({
                            name: "custrecord_bc_proj_subsidiary",
                            join: "CUSTRECORD_BC_BUDGET_PROJECT",
                            summary: "GROUP",
                            label: "CreationOrgXRefCode"
                        }),
                        search.createColumn({
                            name: "custrecord_bc_proj_customer",
                            join: "CUSTRECORD_BC_BUDGET_PROJECT",
                            summary: "GROUP",
                            label: "Customer"
                        }),
                        search.createColumn({
                            name: "externalid",
                            join: "CUSTRECORD_BC_BUDGET_PROJECT",
                            summary: "GROUP",
                            label: "ParentXRefCode"
                        }),
                        search.createColumn({
                            name: "custrecord_bc_proj_number",
                            join: "CUSTRECORD_BC_BUDGET_PROJECT",
                            summary: "GROUP",
                            label: "Project Number"
                        }),
                        search.createColumn({
                            name: "formulatext",
                            summary: "GROUP",
                            formula: "{custrecord_bc_budget_project.custrecord_bc_taxation_xref_xcode}",
                            label: "TaxationXRefCode"
                        }),
                        search.createColumn({
                            name: "custrecordbc_proj_contract_date",
                            join: "CUSTRECORD_BC_BUDGET_PROJECT",
                            summary: "GROUP",
                            label: "Contract Date"
                        })
                    ]
                });
                return costCodeSearchObj;
            } catch (e) {
                log.debug('Get Input Error', e)
            }
        }

        function map(context) {
            function cleanString(str) {
                return (str || '')
                    .replace(/-/g, ' ')
                    .replace(/[^\x00-\x7F]/g, '')
                    .replace(/&/g, 'and')
                    .replace(/[^a-zA-Z0-9 ]/g, '')
                    .replace(/\s+/g, ' ')
                    .trim();
            }

            try {
                var result = JSON.parse(context.value);
                var values = result.values;
                log.debug('result', result)

                var externalId = values["GROUP(externalid)"] || '';

                // Project name
                var projectField = values["GROUP(custrecord_bc_budget_project)"];
                var projectName = typeof projectField === 'object' ? projectField.text : projectField || '';

                // Cost code name (text comes through because it's direct)
                var costCodeField = values["GROUP(custrecord_bc_budget_code)"];
                var costCodeName = typeof costCodeField === 'object' ? costCodeField.text : costCodeField || '';

                // Cost code ID (used for later lookups if needed)
                var costCodeId = values["GROUP(internalid.CUSTRECORD_BC_BUDGET_CODE)"] || '';

                // Customer
                var customerField = values["GROUP(custrecord_bc_proj_customer.CUSTRECORD_BC_BUDGET_PROJECT)"];
                var customerName = (customerField && typeof customerField === 'object') ? customerField.text : '';
                //log.debug('customerName', customerName)

                // Subsidiary (joined)
                var subsidiaryField = values["GROUP(custrecord_bc_proj_subsidiary.CUSTRECORD_BC_BUDGET_PROJECT)"];
                var subsidiary = (subsidiaryField && typeof subsidiaryField === 'object') ? subsidiaryField.text : '';
                var subsidiaryName = subsidiary.split(' : ').pop()
                log.debug('subsidiaryName', subsidiaryName)

                // Other fields
                var parentExternalId = values["GROUP(CUSTRECORD_BC_BUDGET_PROJECT.externalid)"] || '';
                var projectNumber = values["GROUP(CUSTRECORD_BC_BUDGET_PROJECT.custrecord_bc_proj_number)"] || '';
                var taxationXRef = values["GROUP(formulatext)"];
                if (taxationXRef === '- None -') {
                    taxationXRef = '';
                }

                if (!parentExternalId) parentExternalId = projectName;

                var match = (costCodeName || '').toString().match(/^(\d+)/);
                var costCodeNumber = match ? match[1] : '';

                var projectPlusCostCode = cleanString(projectName) + ' - ' + cleanString(costCodeName);
                var xRefCode = cleanString(customerName) + '_' + cleanString(projectName) + '_' + cleanString(costCodeNumber);

                var clockTransferCode = '';
                if (projectNumber && costCodeNumber) {
                    clockTransferCode = projectNumber + '_' + costCodeNumber;
                } else if (projectNumber) {
                    clockTransferCode = projectNumber;
                } else if (costCodeNumber) {
                    clockTransferCode = costCodeNumber;
                }

                var rawDate = values["GROUP(custrecordbc_proj_contract_date.CUSTRECORD_BC_BUDGET_PROJECT)"] || '';
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

                var row = [
                    'D',
                    'Project',
                    'Project',
                    xRefCode,
                    projectPlusCostCode,
                    '',
                    subsidiaryName,
                    parentExternalId,
                    '0',
                    '0',
                    contractDate,
                    '',
                    '',
                    '0',
                    '0',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    taxationXRef
                ];

                context.write({
                    key: 'all',
                    value: JSON.stringify({row: row, id: result.id})
                });

            } catch (e) {
                log.debug('Error in Map', e)
            }

        }

        function reduce(context) {
            log.debug('Start Script -- reduce')

            try {
                var values = context.values;

                var firstRow = ['ProjectImport', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''];
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
                var costCodeIds = [];

                for (var i = 0; i < values.length; i++) {
                    var parsed = JSON.parse(values[i]);
                    var rowData = parsed.row;
                    costCodeIds.push(parsed.id);

                    // Clean XRefCode
                    /*var xrefCode = (rowData[3] || '')
                        .replace(/&/g, 'AND')
                        .replace(/\s+/g, '_')
                        .replace(/[^\w]/g, '')
                        .toUpperCase();*/

                    var parentXrefCode = (rowData[7] || '')
                        .replace(/&/g, 'AND')
                        .replace(/\s+/g, '_')
                        .replace(/[^\w]/g, '')
                        .toUpperCase();

                    var subsidiaryName = (rowData[6] || '')
                        .replace(/&/g, 'AND')
                        .replace(/\s+/g, '_')
                        .replace(/[^\w]/g, '')
                        .toUpperCase();

                    var fullRow = [
                        rowData[0],
                        rowData[1],
                        rowData[2],
                        rowData[3] || '',
                        rowData[4] || '',
                        rowData[5] || '',
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

                log.debug('cost code ids', costCodeIds)

                //Get today's date and timestamp for filename
                var now = new Date();
                var timeStamp = ('0' + now.getHours()).slice(-2) +
                    ('0' + now.getMinutes()).slice(-2) +
                    ('0' + now.getSeconds()).slice(-2);
                var fileName = 'BlueCollar_CostCodeList_' + getTodayString() + '_' + timeStamp + '.csv.ready';
                //var fileName = 'BlueCollar_CostCodeList_' + getTodayString() + '_' + timeStamp + '.csv';

                //Create Export Log custom record
                var exportLog = record.create({
                    type: 'customrecord_bc_df_export_log',
                    isDynamic: true
                });
                exportLog.setValue({fieldId: 'name', value: fileName});
                exportLog.setValue({fieldId: 'custrecord_bc_df_export_type', value: 3}); //Project
                exportLog.setValue({fieldId: 'custrecord_bc_df_export_date', value: new Date()});
                exportLog.setValue({fieldId: 'custrecord_bc_df_export_status', value: 1}); //Pending
                exportLog.setValue({fieldId: 'custrecord_bc_df_export_record_ids', value: costCodeIds.join(', ')});
                var exportLogId = exportLog.save();
                log.debug('exportLogId', exportLogId)

                var fileId = null;

                //Start CSV creation
                try {
                    var csvContent = firstRow.join(',') + '\n' + headers.join(',') + '\n' + rows.join('\n');
                    var fileObj = file.create({
                        name: fileName,
                        fileType: file.Type.CSV,
                        contents: csvContent,
                        folder: 1669
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

                    sftpConnection.upload({
                        directory: '/Import/ProjectImport',
                        filename: fileObj.name,
                        file: fileObj,
                        //replaceExisting: true
                    });

                    var files = sftpConnection.list({path: '/Import/ProjectImport'});
                    log.debug('Directory Listing', JSON.stringify(files));
                    log.debug('sftpConnection', sftpConnection)

                    record.submitFields({
                        type: 'customrecord_bc_df_export_log',
                        id: exportLogId,
                        values: {
                            custrecord_bc_df_export_file: fileId,
                            custrecord_bc_df_export_status: 2,
                            custrecord_bc_df_export_logs: 'File exported and uploaded successfully.'
                        }
                    });
                } catch (e) {
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

            } catch (e) {
                log.debug('Error', e)
            }
        }

        function summarize(summary) {

            var mrTask = task.create({
                    taskType: task.TaskType.MAP_REDUCE,
                    scriptId: 'customscript_bc_df_mr_update_costcode',
                    deploymentId: 'customdeploy_bc_df_mr_update_costcode',
                });
            mrTask.submit();
            log.debug('Triggered Cost Code Update MR');

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
            summarize: summarize
        };
    });