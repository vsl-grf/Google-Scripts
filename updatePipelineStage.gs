function updatePipelineStage() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DCP_payments_checker');
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WorkdeskQ125');

  if (!sourceSheet || !targetSheet) {
    Logger.log('Error: One or both sheets not found.');
    return;
  }

  var sourceHeaders = sourceSheet.getDataRange().getValues()[0];
  var targetHeaders = targetSheet.getDataRange().getValues()[0];

  var dcpInvoiceIndex = sourceHeaders.indexOf('DCP_invoice');
  var dcpStatusIndex = sourceHeaders.indexOf('DCP_status');
  var refresherStatusIndex = sourceHeaders.indexOf('refresher_status');
  var cbInvoiceIndex = targetHeaders.indexOf('CB INV Number');
  var pipelineStageIndex = targetHeaders.indexOf('Pipeline Stage');

  if ([dcpInvoiceIndex, dcpStatusIndex, refresherStatusIndex, cbInvoiceIndex, pipelineStageIndex].includes(-1)) {
    Logger.log('Error: One or more specified columns not found.');
    return;
  }

  var sourceData = sourceSheet.getDataRange().getValues();
  var targetData = targetSheet.getDataRange().getValues();

  // Counter for successful updates
  var updatesMade = 0;

  for (var i = 1; i < sourceData.length; i++) {
    var sourceInvoice = sourceData[i][dcpInvoiceIndex]?.toString().trim();
    var dcpStatus = sourceData[i][dcpStatusIndex]?.toString().trim().toLowerCase();
    var refresherStatus = sourceData[i][refresherStatusIndex]?.toString().trim().toLowerCase();

    // Skip empty rows or if any key value is missing
    if (!sourceInvoice || !dcpStatus || !refresherStatus) continue;

    if (dcpStatus === 'collected' && refresherStatus === 'paid') {
      for (var j = 1; j < targetData.length; j++) {
        var targetInvoice = targetData[j][cbInvoiceIndex]?.toString().trim();

        if (targetInvoice === sourceInvoice) {
          var currentStage = targetSheet.getRange(j + 1, pipelineStageIndex + 1).getValue().toString().trim().toLowerCase();

          // Prevent redundant updates
          if (currentStage !== 'collected') {
            targetSheet.getRange(j + 1, pipelineStageIndex + 1).setValue('collected');
            updatesMade++;
            Logger.log('Updated row ' + (j + 1));
          } else {
            Logger.log('No change required for row ' + (j + 1));
          }
        }
      }
    }
  }

  Logger.log('Total updates made: ' + updatesMade);
}
