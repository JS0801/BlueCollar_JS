/**
 * @NApiVersion 2.1
 * @NScriptType MapReduceScript
 */
define(['N/search', 'N/record', 'N/log'], function (search, record, log) {

    function getInputData() {
        // Return ALL budget items that are unsent and belong to a sent project
        return search.create({
            type: 'customrecord_bc_budget_item',
            filters: [
                ['isinactive', 'is', 'F'], 'AND',
                ['custrecord_bc_costcode_sent_to_df', 'is', 'F'], 'AND',
                ['custrecord_bc_budget_project.custrecord_bc_sent_to_df', 'is', 'T']
            ],
            columns: ['internalid']
        });
    }

    function map(context) {
        var row = JSON.parse(context.value);
        var id = row.id;
        try {
            record.submitFields({
                type: 'customrecord_bc_budget_item',
                id: id,
                values: {custrecord_bc_costcode_sent_to_df: true},
                options: {enablesourcing: false, ignoreMandatoryFields: true}
            });
            log.debug('Updated budget item', id);
        } catch (e) {
            log.error('Update failed ' + id, e && e.message ? e.message : e);
        }
    }

    function reduce() {
    } // not used

    function summarize(summary) {
        var errors = 0;
        summary.mapSummary.errors.iterator().each(function (key, e) {
            errors++;
            log.debug('Map error for ID ' + key, e);
            return true;
        });
        log.debug('Mark-as-sent complete', {
            totalKeys: summary.inputSummary ? summary.inputSummary.keys : null,
            errors: errors
        });
    }

    return {getInputData, map, reduce, summarize};
});
