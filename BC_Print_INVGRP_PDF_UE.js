/**
* @NApiVersion 2.x
* @NScriptType UserEventScript
* @NModuleScope SameAccount
*/
define(['N/ui/serverWidget', 'N/url', 'N/search', 'N/record'], function(ui, url, search, record) {
  
  function beforeLoad(context) {
    if (context.type === context.UserEventType.VIEW) {
      var form = context.form;
      var current_rec = context.newRecord;
      var recordId = context.newRecord.id;
      var subid = current_rec.getValue('subsidiary');
      var trandate = current_rec.getValue('trandate');
      var internalIdString = null;


      const salesorderSearchObj = search.create({
    type: "salesorder",
    settings: [{ "name": "consolidationtype", "value": "ACCTTYPE" }],
    filters: [
        ["type", "anyof", "SalesOrd"],
        "AND",
        ["internalid", "anyof", recordId],
        "AND",
        ["mainline", "is", "T"]
    ],
    columns: [
        search.createColumn({
            name: "internalid",
            join: "CUSTBODY_ASSOCIATED_SALES_ORDERS",
            label: "Internal ID"
        }),
        search.createColumn({
            name: "trandate",
            join: "CUSTBODY_ASSOCIATED_SALES_ORDERS",
            label: "Date",
            sort: search.Sort.ASC
        })
    ]
});

const results = [];
salesorderSearchObj.run().each(function (result) {
    const id = result.getValue({
        name: "internalid",
        join: "CUSTBODY_ASSOCIATED_SALES_ORDERS"
    });
    const date = result.getValue({
        name: "trandate",
        join: "CUSTBODY_ASSOCIATED_SALES_ORDERS"
    });
    results.push({ id: id, date: new Date(date) });
    return true;
});

      log.debug('results', results)

      if (results.length == 0) {
        internalIdString = recordId;
      }else{
        results.push({ id: recordId, date: new Date(trandate) });
        results.sort(function (a, b) {
            return a.date - b.date;
        });

        internalIdString = results.map(function (entry) {
            return entry.id;
        }).join(',');
      }


log.debug('Sorted Internal IDs', internalIdString);
      // var relatedSO = current_rec.getValue('custbody_associated_sales_orders');
      // log.debug('relatedSO', relatedSO)
      // if (relatedSO && relatedSO.length > 0) relatedSO = relatedSO.join(',') + ',' + recordId;
      // else relatedSO = recordId;
      // log.debug('relatedSO', relatedSO)
      
      // Add a custom button to the form
      if (current_rec.type == 'salesorder') {
        
        var scriptUrl = url.resolveScript({
          scriptId: 'customscript_bc_sl_so_pdf_helper',
          deploymentId: 'customdeploy1'
        });

        if (recordId == 7282){
        scriptUrl = url.resolveScript({
          scriptId: 'customscript_bc_export_timesheet_v2',
          deploymentId: 'customdeploy_bc_export_timesheet_v2'
        });
        }
        
        var buttonScript =
        "(function(){\
          function rm(){\
            var ov=document.getElementById('csvOverlay'); if(ov) ov.remove();\
            var st=document.getElementById('csvOverlayStyle'); if(st) st.remove();\
          }\
          var o=document.createElement('div');\
          o.id='csvOverlay';\
          o.style.cssText='position:fixed;inset:0;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.45);z-index:9999;';\
          o.innerHTML='<div style=\"display:flex;flex-direction:column;align-items:center;gap:12px;background:#fff;padding:24px 28px;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.2);\">\
          <div style=\\\"width:40px;height:40px;border:4px solid #e5e7eb;border-top-color:#111;border-radius:50%;animation:csvspin 1s linear infinite\\\"></div>\
          <div style=\\\"font:600 14px/1.2 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;\\\">Exporting Excel file…</div>\
          </div>';\
          document.body.appendChild(o);\
          var st=document.createElement('style'); st.id='csvOverlayStyle'; st.textContent='@keyframes csvspin{to{transform:rotate(360deg)}}'; document.head.appendChild(st);\
          var url='" + scriptUrl + "&export=excel&tranid=" + internalIdString + "';\
          var d=new Date(), yyyy=d.getFullYear(), mm=('0'+(d.getMonth()+1)).slice(-2), dd=('0'+d.getDate()).slice(-2);\
          var filename='Weekly_Timesheet_'+yyyy+'-'+mm+'-'+dd+'_" + recordId + ".xls';\
          var failTimer=setTimeout(rm,4000); /* failsafe cleanup */ \
          fetch(url,{credentials:'include'}).then(function(r){\
            if(!r.ok) throw new Error(r.statusText);\
            return r.blob();\
          }).then(function(blob){\
            var a=document.createElement('a');\
            a.href=URL.createObjectURL(blob);\
            a.download=filename;\
            document.body.appendChild(a); a.click();\
            URL.revokeObjectURL(a.href); a.remove();\
          }).catch(function(){\
            /* Fallback: navigate to Suitelet (browser will download attachment) */\
            window.location.href=url;\
            window.addEventListener('focus', rm, { once:true });\
          }).finally(function(){\
            clearTimeout(failTimer);\
            rm();\
          });\
        })();";
        
        form.addButton({
          id: 'custpage_export',
          label: 'Export Timesheet',
          functionName: buttonScript
        });
        
      } else {
        var custid = current_rec.getValue('customer');
        var scriptUrl = url.resolveScript({
          scriptId: 'customscript_bc_sl_inv_grp_pdf_summary',
          deploymentId: 'customdeploy1'
        });
        
        var buttonScript = "window.open('" + scriptUrl + "&recid=" + recordId + "&subid=" + subid + "&custid=" + custid +"', '_blank');";
        
        form.addButton({
          id: 'custpage_my_button',
          label: 'Summary PDF',
          functionName: buttonScript
        });
        
        var scriptUrl1 = url.resolveScript({
          scriptId: 'customscript_bc_inv_detailed_pdf',
          deploymentId: 'customdeploy1'
        });
        
        var buttonScript1 = "window.open('" + scriptUrl1 + "&recid=" + recordId + "&subid=" + subid + "&custid=" + custid +"', '_blank');";
        
        form.addButton({
          id: 'custpage_my_button1',
          label: 'Detail PDF',
          functionName: buttonScript1
        });
        
        form.addButton({
          id: 'custpage_my_buttoncsv',
          label: 'Export Excel',
          functionName:
          "var o=document.createElement('div');" +
          "o.id='csvOverlay';" +
          "o.style.cssText='position:fixed;inset:0;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.45);z-index:9999;';" +
          "o.innerHTML='<div style=\"display:flex;flex-direction:column;align-items:center;gap:12px;background:#fff;padding:24px 28px;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.2);\">" +
          "<div style=\\\"width:40px;height:40px;border:4px solid #e5e7eb;border-top-color:#111;border-radius:50%;animation:csvspin 1s linear infinite\\\"></div>" +
          "<div style=\\\"font:600 14px/1.2 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;\\\">Exporting Excel file…</div>" +
          "</div>';document.body.appendChild(o);" +
          "(function(){" +
          "var st=document.createElement('style');st.id='csvOverlayStyle';st.textContent='@keyframes csvspin{to{transform:rotate(360deg)}}';document.head.appendChild(st);" +
          "var url='" + scriptUrl1 + "&type=CSV&recid=" + recordId + "&subid=" + subid + "&custid=" + custid + "';" +
          "fetch(url,{credentials:\"include\"}).then(function(r){if(!r.ok)throw new Error(r.statusText);return r.blob();}).then(function(blob){" +
          "var a=document.createElement('a');" +
          "a.href=URL.createObjectURL(blob);" +
          "a.download='Invoice_Detailed_" + recordId + ".xls';" +
          "document.body.appendChild(a);a.click();URL.revokeObjectURL(a.href);a.remove();" +
          "var ov=document.getElementById(\"csvOverlay\");if(ov)ov.remove();var s=document.getElementById(\"csvOverlayStyle\");if(s)s.remove();" +
          "}).catch(function(e){" +
          "window.location.href=url;" +
          "function rm(){var ov=document.getElementById(\"csvOverlay\");if(ov)ov.remove();var s=document.getElementById(\"csvOverlayStyle\");if(s)s.remove();window.removeEventListener(\"focus\",rm);}window.addEventListener(\"focus\",rm);" +
          "});" +
          "})();"
        });
        
        var scriptUrlUS = url.resolveScript({
          scriptId: 'customscript_bc_invoice_detaield_pdf_us',
          deploymentId: 'customdeploy_bc_invoice_detaield_pdf_us'
        });
        
        var buttonScript1 = "window.open('" + scriptUrlUS + "&recid=" + recordId + "&subid=" + subid + "&custid=" + custid +"', '_blank');";
        
        form.addButton({
          id: 'custpage_my_button1',
          label: 'Detail PDF (Subtotal)',
          functionName: buttonScript1
        });
        
        form.addButton({
          id: 'custpage_my_buttoncsv',
          label: 'Export Excel (Subtotal)',
          functionName:
          "var o=document.createElement('div');" +
          "o.id='csvOverlay';" +
          "o.style.cssText='position:fixed;inset:0;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.45);z-index:9999;';" +
          "o.innerHTML='<div style=\"display:flex;flex-direction:column;align-items:center;gap:12px;background:#fff;padding:24px 28px;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.2);\">" +
          "<div style=\\\"width:40px;height:40px;border:4px solid #e5e7eb;border-top-color:#111;border-radius:50%;animation:csvspin 1s linear infinite\\\"></div>" +
          "<div style=\\\"font:600 14px/1.2 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;\\\">Exporting Excel file…</div>" +
          "</div>';document.body.appendChild(o);" +
          "(function(){" +
          "var st=document.createElement('style');st.id='csvOverlayStyle';st.textContent='@keyframes csvspin{to{transform:rotate(360deg)}}';document.head.appendChild(st);" +
          "var url='" + scriptUrlUS + "&type=CSV&recid=" + recordId + "&subid=" + subid + "&custid=" + custid + "';" +
          "fetch(url,{credentials:\"include\"}).then(function(r){if(!r.ok)throw new Error(r.statusText);return r.blob();}).then(function(blob){" +
          "var a=document.createElement('a');" +
          "a.href=URL.createObjectURL(blob);" +
          "a.download='Invoice_Detailed_" + recordId + ".xls';" +
          "document.body.appendChild(a);a.click();URL.revokeObjectURL(a.href);a.remove();" +
          "var ov=document.getElementById(\"csvOverlay\");if(ov)ov.remove();var s=document.getElementById(\"csvOverlayStyle\");if(s)s.remove();" +
          "}).catch(function(e){" +
          "window.location.href=url;" +
          "function rm(){var ov=document.getElementById(\"csvOverlay\");if(ov)ov.remove();var s=document.getElementById(\"csvOverlayStyle\");if(s)s.remove();window.removeEventListener(\"focus\",rm);}window.addEventListener(\"focus\",rm);" +
          "});" +
          "})();"
        });
      }
    }
  }
  
  function afterSubmit(context) {
    var current_rec = context.newRecord;
    var recordId = current_rec.id;
    
    //Only run on CREATE, and skip sales orders
    if (context.type !== context.UserEventType.CREATE || current_rec.type === 'salesorder') return;
    
    var invoiceSearchObj = search.create({
      type: "invoice",
      settings:[{"name":"consolidationtype","value":"NONE"}],
      filters:
      [
        ["type","anyof","CustInvc"], 
        "AND", 
        ["groupedto","anyof",recordId], 
        "AND", 
        ["custcol_invoicing_category","noneof","@NONE@"]
      ],
      columns:
      [
        search.createColumn({
          name: "custrecord_cponum",
          join: "cseg_bc_project",
          summary: "MAX"
        }),
        search.createColumn({
          name: "otherrefnum",
          summary: "MAX"
        })
      ]
    });
    
    invoiceSearchObj.run().each(function(result){
      record.submitFields({
        type: 'invoicegroup',
        id: recordId,
        values:{
          ponumber: result.getValue({ name: "custrecord_cponum", join: "cseg_bc_project", summary: "MAX"}),
          custrecord_cust_ref: result.getValue({ name: "otherrefnum", summary: "MAX"})
        }
      });
      return true;
    });
  }
  
  return {
    beforeLoad: beforeLoad,
    afterSubmit: afterSubmit
  };
});