/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['N/search', 'N/record', 'N/task', 'N/runtime', 'N/url', 'N/log', 'N/file'],
function (search, record, task, runtime, url, log, file) {

  const ZIP_MR_SCRIPT_ID = 'customscript_bc_mr_search_file_downloade';
  const ZIP_MR_DEPLOY_ID = 'customdeploy_bc_mr_search_file_downloade';

  function esc(s) {
    if (s === null || s === undefined) return '';
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function toDownloadUrl(u) {
  if (!u) return '';
  u = String(u);

  // If already has download flag, keep it
  if (u.indexOf('download=') !== -1) return u;

  // Add download=T to force attachment download
  return u + (u.indexOf('?') !== -1 ? '&' : '?') + 'download=T';
}

  function formatBytes(bytes) {
    if (!bytes) return '0 B';
    var k = 1024;
    var sizes = ['B', 'KB', 'MB', 'GB'];
    var i = Math.floor(Math.log(bytes) / Math.log(k));
    return (bytes / Math.pow(k, i)).toFixed(2) + ' ' + sizes[i];
  }

  function parseParams(req) {
    return {
      internalid: req.parameters.internalid || '',
      soid: req.parameters.soid || '',
      downloadall: req.parameters.downloadall || ''
    };
  }

  /**
   * YOUR LOGIC (kept) - builds arrayids using:
   * - Sales Order grouped + Vendor Bill grouped
   * - Vendor Bill -> Applied To PO
   */
  function buildArrayIds(p) {
    var arrayids = [];

    // If user clicked a specific record link, include those too (optional but useful)
    if (p.internalid) arrayids.push(p.internalid);
    if (p.soid) arrayids.push(p.soid);

    var salesorderSearchObj = search.create({
      type: "salesorder",
      settings: [{ "name": "consolidationtype", "value": "NONE" }],
      filters: [
        ["type", "anyof", "SalesOrd"],
        "AND",
        ["mainline", "is", "F"],
        "AND",
        ["formulatext: case when {custcol_bc_tm_line_id} = {custcol_bc_tm_source_transaction.line} then 1 else 0 end", "is", "1"],
        "AND",
        ["custcol_bc_tm_source_transaction.mainline", "is", "F"],
        "AND",
        ["custcol_invoicing_category", "anyof", "4"]
      ],
      columns: [
        search.createColumn({ name: "internalid", summary: "GROUP", label: "internalid" }),
        search.createColumn({ name: "entity", summary: "GROUP", label: "Name" }),
        search.createColumn({ name: "internalid", join: "CUSTCOL_BC_TM_SOURCE_TRANSACTION", summary: "GROUP", label: "Vendor Bill #" })
      ]
    });

    var soCount = salesorderSearchObj.runPaged().count;
    log.debug("salesorderSearchObj result count", soCount);

    salesorderSearchObj.run().each(function (result) {
      var vb = result.getValue({ name: 'internalid', join: "CUSTCOL_BC_TM_SOURCE_TRANSACTION", summary: "GROUP" });
      var so = result.getValue({ name: 'internalid', summary: "GROUP" });

      if (vb) arrayids.push(vb);
      if (so) arrayids.push(so);
      return true;
    });

    var vendorbillSearchObj = search.create({
      type: "vendorbill",
      settings: [{ "name": "consolidationtype", "value": "ACCTTYPE" }],
      filters: [
        ["type", "anyof", "VendBill"],
        "AND",
        ["internalid", "anyof", arrayids],
        "AND",
        ["appliedtotransaction.type", "anyof", "PurchOrd"]
      ],
      columns: [
        search.createColumn({
          name: "appliedtotransaction",
          summary: "GROUP",
          label: "Applied To Transaction",
          sort: search.Sort.ASC
        })
      ]
    });

    var vbCount = vendorbillSearchObj.runPaged().count;
    log.debug("vendorbillSearchObj result count", vbCount);

    vendorbillSearchObj.run().each(function (result) {
      var poId = result.getValue({ name: "appliedtotransaction", summary: "GROUP" });
      if (poId) arrayids.push(poId);
      return true;
    });

    // De-dupe
    var dedupe = {};
    var finalIds = [];
    for (var i = 0; i < arrayids.length; i++) {
      var id = String(arrayids[i] || '');
      if (!id || dedupe[id]) continue;
      dedupe[id] = true;
      finalIds.push(id);
    }

    return finalIds;
  }

  function getFilesForTransactions(arrayids) {
    var out = [];
    var seen = {};

    if (!arrayids || !arrayids.length) return out;

    var transactionSearchObj = search.create({
      type: "transaction",
      settings: [{ name: "consolidationtype", value: "ACCTTYPE" }],
      filters: [
        ["internalid", "anyof", arrayids],
        "AND",
        ["file.internalidnumber", "isnotempty", ""]
      ],
      columns: [
        search.createColumn({ name: "internalid", join: "file", summary: "GROUP", sort: search.Sort.ASC }),
        search.createColumn({ name: "name", join: "file", summary: "GROUP" }),
        search.createColumn({ name: "documentsize", join: "file", summary: "GROUP" }),
        search.createColumn({ name: "url", join: "file", summary: "GROUP" })
      ]
    });

    transactionSearchObj.run().each(function (r) {
      var fileId = r.getValue({ name: "internalid", join: "file", summary: "GROUP" });
      if (!fileId || seen[fileId]) return true;
      seen[fileId] = true;

      out.push({
        fileId: fileId,
        name: r.getValue({ name: "name", join: "file", summary: "GROUP" }) || ('File ' + fileId),
        size: parseInt(r.getValue({ name: "documentsize", join: "file", summary: "GROUP" }), 10) || 0,
        url: r.getValue({ name: "url", join: "file", summary: "GROUP" }) || ''
      });
      return true;
    });

    return out;
  }

  function renderModal(res, files, params, msgHtml) {
    var rows = '';
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      rows += ''
        + '<tr>'
        + '  <td style="padding:8px;border-bottom:1px solid #eee;">' + esc(f.name) + '</td>'
        + '  <td style="padding:8px;border-bottom:1px solid #eee;white-space:nowrap;">' + esc(formatBytes(f.size)) + '</td>'
        + '  <td style="padding:8px;border-bottom:1px solid #eee;white-space:nowrap;">'
        + (f.url ? '<a href="' + esc(f.url) + "&_xd=T" + '" rel="noopener">Download</a>' : '-')
        + '  </td>'
        + '</tr>';
    }

    var html =
      '<!doctype html><html><head><meta charset="utf-8"/><title>Attachments</title>'
      + '<style>'
      + 'body{font-family:Arial, sans-serif;margin:0;}'
      + '.overlay{position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;padding:24px;}'
      + '.modal{background:#fff;width:min(1100px,96vw);max-height:92vh;border-radius:10px;box-shadow:0 10px 40px rgba(0,0,0,.35);overflow:hidden;}'
      + '.hdr{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid #e6e6e6;}'
      + '.ttl{font-size:16px;font-weight:700;}'
      + '.btn{border:1px solid #ccc;background:#fafafa;padding:7px 10px;border-radius:8px;cursor:pointer;}'
      + '.btnPrimary{border:1px solid #1b73e8;background:#1b73e8;color:#fff;}'
      + '.body{padding:14px 16px;overflow:auto;max-height:calc(92vh - 120px);}'
      + 'table{width:100%;border-collapse:collapse;}'
      + 'th{font-size:12px;text-align:left;color:#555;padding:8px;border-bottom:1px solid #ddd;position:sticky;top:0;background:#fff;}'
      + '.note{background:#f7f7f7;border:1px solid #e5e5e5;padding:10px;border-radius:8px;margin-bottom:10px;}'
      + '.msg{margin-bottom:10px;}'
      + '</style></head><body>'
      + '<div class="overlay"><div class="modal">'
      + '  <div class="hdr">'
      + '    <div class="ttl">Attachments (' + files.length + ')</div>'
      + '    <div>'
      + '      <button class="btn" onclick="window.close()">Close</button> '
      + '      <button class="btn btnPrimary" onclick="downloadAllEmailZip()">Download All (Email Zip)</button>'
      + '    </div>'
      + '  </div>'
      + '  <div class="body">'
      + '    <div class="note"><b>Tip:</b> Use individual download links, or click <b>Download All (Email Zip)</b> to receive zip link(s) by email.</div>'
      + (msgHtml ? '<div class="msg">' + msgHtml + '</div>' : '')
      + '    <table>'
      + '      <thead><tr><th>File Name</th><th>Size</th><th>Action</th></tr></thead>'
      + '      <tbody>' + (rows || '<tr><td colspan="3" style="padding:10px;">No files found.</td></tr>') + '</tbody>'
      + '    </table>'
      + '  </div>'
      + '</div></div>'

      + '<form id="zipForm" method="POST">'
      + '  <input type="hidden" name="internalid" value="' + esc(params.internalid) + '"/>'
      + '  <input type="hidden" name="soid" value="' + esc(params.soid) + '"/>'
      + '  <input type="hidden" name="downloadall" value="' + esc(params.downloadall) + '"/>'
      + '  <input type="hidden" name="email" id="emailFld" value=""/>'
      + '</form>'

      + '<script>'
      + 'function downloadAllEmailZip(){'
      + '  var email = prompt("Enter email to receive zip download link(s):");'
      + '  if(!email) return;'
      + '  email = String(email).trim();'
      + '  if(email.indexOf("@") === -1){ alert("Please enter a valid email."); return; }'
      + '  document.getElementById("emailFld").value = email;'
      + '  document.getElementById("zipForm").submit();'
      + '}'
      + '</script>'

      + '</body></html>';

    res.write(html);
  }

  function onRequest(context) {
    var req = context.request;
    var res = context.response;

    if (req.method === 'GET') {
      var p = parseParams(req);
      var arrayids = buildArrayIds(p);
      var files = getFilesForTransactions(arrayids);
      renderModal(res, files, p, '');
      return;
    }

    // POST: create job + trigger MR + show message
    try {
      var p2 = parseParams(req);
      var emailAddr = (req.parameters.email || '').trim();

      var arrayids2 = buildArrayIds(p2);
      var files2 = getFilesForTransactions(arrayids2);

      if (!emailAddr) {
        renderModal(res, files2, p2, '<div style="color:#b00020;font-weight:700;">Missing email.</div>');
        return;
      }

      // Create Zip Job record
      var jobRec = record.create({ type: 'customrecord_tc_zip_job', isDynamic: true });
      jobRec.setValue({ fieldId: 'custrecord_tc_zip_status', value: 'Pending' });
      jobRec.setValue({ fieldId: 'custrecord_tc_zip_email', value: emailAddr });
      jobRec.setValue({
        fieldId: 'custrecord_tc_zip_params',
        value: JSON.stringify({
          internalid: p2.internalid,
          soid: p2.soid,
          downloadall: p2.downloadall,
          requestedBy: runtime.getCurrentUser().id,
          requestedOn: new Date().toISOString()
        })
      });
      var jobId = jobRec.save();

      // Trigger Map/Reduce
      var mrTask = task.create({
        taskType: task.TaskType.MAP_REDUCE,
        scriptId: ZIP_MR_SCRIPT_ID,
        deploymentId: ZIP_MR_DEPLOY_ID,
        params: { custscript_tc_zip_job_id: String(jobId) }
      });
      mrTask.submit();

      renderModal(
        res,
        files2,
        p2,
        '<div style="color:#137333;font-weight:700;">Zip request submitted.</div>'
        + '<div>Email will be sent to <b>' + esc(emailAddr) + '</b> when ready. Job ID: <b>' + esc(jobId) + '</b></div>'
      );

    } catch (e) {
      log.error('Suitelet POST Error', e);
      var pE = parseParams(req);
      var idsE = buildArrayIds(pE);
      var filesE = getFilesForTransactions(idsE);
      renderModal(res, filesE, pE, '<div style="color:#b00020;">Error: ' + esc(e.message || e) + '</div>');
    }
  }

  return { onRequest: onRequest };
});