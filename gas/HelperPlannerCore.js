/**
 * HELPER PLANNER CORE CORE ENGINE
 * Implements Google Sheets worksheets database operations, active override scaling, and backend API routes.
 */

const SHEET_HELPER_PLAN = 'HELPER_PLAN';
const SHEET_HELPER_BOM_MASTER = 'HELPER_BOM_MASTER';
const SHEET_HELPER_BOM_ADJUSTMENTS = 'HELPER_BOM_ADJUSTMENTS';
const SHEET_HELPER_BOM_ACCESS = 'HELPER_BOM_ACCESS';
const SHEET_HELPER_AUDIT_LOG = 'HELPER_AUDIT_LOG';

let _sheetsInitialized = false;

/**
 * Setup and initialize all database sheets if they do not exist
 */
function setupHelperPlannerSheets() {
  if (_sheetsInitialized) return;
  
  const requiredSheets = [
    {
      name: SHEET_HELPER_PLAN,
      headers: ['Plan ID', 'Production Date', 'Physical Area', 'Material / SKU Code', 'Material Description', 'Machine', 'Shift', 'Original Output (CS)', 'Current Output (CS)', 'Source', 'Status']
    },
    {
      name: SHEET_HELPER_BOM_MASTER,
      headers: ['SKU Code', 'SKU Description', 'Component Code', 'Component Description', 'Base Quantity per 1000 CS', 'UOM']
    },
    {
      name: SHEET_HELPER_BOM_ADJUSTMENTS,
      headers: ['BOM Version', 'SKU Code', 'Original Component Code', 'Adjusted Component Code', 'Original Quantity', 'Adjusted Quantity', 'UOM', 'Adjusted By', 'Adjusted At', 'Comment', 'Status', 'Effective At']
    },
    {
      name: SHEET_HELPER_BOM_ACCESS,
      headers: ['Email', 'Authorized By', 'Authorized At', 'Status']
    },
    {
      name: SHEET_HELPER_AUDIT_LOG,
      headers: ['Timestamp', 'User', 'Category', 'Target ID', 'Action', 'Old Value', 'New Value', 'Details']
    }
  ];

  requiredSheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.getRange(1, 1, 1, s.headers.length).setValues([s.headers])
           .setBackground('#1e293b')
           .setFontColor('white')
           .setFontWeight('bold');
      sheet.setFrozenRows(1);
      
      // Auto-authorize active user on creation to avoid lockout
      if (s.name === SHEET_HELPER_BOM_ACCESS) {
        const activeEmail = Session.getActiveUser().getEmail() || 'planner@example.com';
        sheet.appendRow([activeEmail, 'SYSTEM', new Date(), 'ACTIVE']);
      }
    }
  });

  // Schema check & auto-migration for adjusted sheet
  const adjSheet = ss.getSheetByName(SHEET_HELPER_BOM_ADJUSTMENTS);
  if (adjSheet) {
    const lastCol = adjSheet.getLastColumn();
    if (lastCol > 0) {
      const headers = adjSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.indexOf('Original Component Code') < 0 || headers.indexOf('Effective At') < 0) {
        if (headers.indexOf('Original Component Code') >= 0 && headers.indexOf('Effective At') < 0) {
          // Gracefully append the column to avoid destroying active overrides
          adjSheet.getRange(1, lastCol + 1).setValue('Effective At')
                  .setBackground('#1e293b')
                  .setFontColor('white')
                  .setFontWeight('bold');
        } else {
          // Re-create headers for clean migration
          adjSheet.clear();
          const newHeaders = ['BOM Version', 'SKU Code', 'Original Component Code', 'Adjusted Component Code', 'Original Quantity', 'Adjusted Quantity', 'UOM', 'Adjusted By', 'Adjusted At', 'Comment', 'Status', 'Effective At'];
          adjSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders])
               .setBackground('#1e293b')
               .setFontColor('white')
               .setFontWeight('bold');
          adjSheet.setFrozenRows(1);
        }
      }
    }
  }

  _sheetsInitialized = true;
}

/**
 * Sync plan input and SAP BOM from raw sheets into Helper Planner database worksheets
 * Keeps manual overrides and ad-hoc plans intact.
 */
function syncHelperPlannerMasterData() {
  setupHelperPlannerSheets();

  // Verify Admin Privilege
  const activeUserEmail = Session.getActiveUser().getEmail() || '';
  let hasBomAdjustPrivilege = false;
  
  const accessSheet = ss.getSheetByName(SHEET_HELPER_BOM_ACCESS);
  if (accessSheet) {
    const accessValues = accessSheet.getDataRange().getValues();
    if (accessValues.length > 1) {
      const emailIdx = accessValues[0].indexOf('Email');
      const statusIdx = accessValues[0].indexOf('Status');
      for (let i = 1; i < accessValues.length; i++) {
        const email = accessValues[i][emailIdx].toString().trim().toLowerCase();
        const status = accessValues[i][statusIdx].toString().trim().toUpperCase();
        if (email === activeUserEmail.toLowerCase() && status === 'ACTIVE') {
          hasBomAdjustPrivilege = true;
          break;
        }
      }
    }
  }
  
  if (!hasBomAdjustPrivilege) {
    throw new Error("Bạn không có quyền đồng bộ cơ sở dữ liệu! Vui lòng liên hệ Admin.");
  }

  // 1. Sync Planning input to HELPER_PLAN
  const planInputSheet = ss.getSheetByName('Planning input');
  if (!planInputSheet) {
    throw new Error("Không tìm thấy sheet 'Planning input'!");
  }

  const rawPlanValues = planInputSheet.getDataRange().getValues();
  if (rawPlanValues.length <= 1) return;

  const planHeaders = rawPlanValues[0];
  const dateIdx = planHeaders.indexOf('Production Date');
  const areaIdx = planHeaders.indexOf('Physical Area');
  const skuIdx = planHeaders.indexOf('Material / SKU Code');
  const descIdx = planHeaders.indexOf('Material Description');
  const machineIdx = planHeaders.indexOf('machine');
  const shiftIdx = planHeaders.indexOf('Shift');
  const outputIdx = planHeaders.indexOf('Output');

  if (dateIdx < 0 || areaIdx < 0 || skuIdx < 0 || outputIdx < 0) {
    throw new Error("Cấu trúc sheet 'Planning input' không đúng!");
  }

  // Fetch only active preserved ad-hoc plans and modified current outputs
  const helperPlanSheet = ss.getSheetByName(SHEET_HELPER_PLAN);
  let adhocRows = [];
  let existingOutputs = {};
  
  const existingHelperData = helperPlanSheet.getDataRange().getValues();
  if (existingHelperData.length > 1) {
    const hpHeaders = existingHelperData[0];
    const hpIdIdx = hpHeaders.indexOf('Plan ID');
    const hpOutputIdx = hpHeaders.indexOf('Current Output (CS)');
    const hpSourceIdx = hpHeaders.indexOf('Source');
    const hpStatusIdx = hpHeaders.indexOf('Status');
    const hpOrigIdx = hpHeaders.indexOf('Original Output (CS)');

    for (let i = 1; i < existingHelperData.length; i++) {
      const row = existingHelperData[i];
      const source = row[hpSourceIdx];
      const status = row[hpStatusIdx];
      const planId = row[hpIdIdx];
      
      if (source === 'ADHOC' && status === 'ACTIVE') {
        adhocRows.push(row);
      } else if (source === 'SYNC') {
        const origOutput = row[hpOrigIdx];
        const currOutput = row[hpOutputIdx];
        if (origOutput !== currOutput) {
          existingOutputs[planId] = currOutput;
        }
      }
    }
  }

  // Load Size Type lookup for Case Count conversion (PC -> CS)
  const sizeTypeSheet = ss.getSheetByName('Size Type');
  let caseCountLookup = {};
  if (sizeTypeSheet) {
    const sizeTypeValues = sizeTypeSheet.getDataRange().getValues();
    if (sizeTypeValues.length > 1) {
      let headerRowIndex = -1;
      let skuColIdx = -1;
      let factorColIdx = -1;
      
      // Scan the first 3 rows to find where the headers are located
      for (let i = 0; i < Math.min(3, sizeTypeValues.length); i++) {
        const row = sizeTypeValues[i];
        const skuIdx = row.indexOf('Material / SKU Code');
        const factorIdx = row.indexOf('PrimaryConvFactor');
        
        if (skuIdx >= 0 && factorIdx >= 0) {
          headerRowIndex = i;
          skuColIdx = skuIdx;
          factorColIdx = factorIdx;
          break; // Stop scanning once we find the perfect header row!
        }
      }
      
      // Fallbacks if not found
      if (skuColIdx < 0) skuColIdx = 26; // Default to Column AA
      if (factorColIdx < 0) factorColIdx = 28; // Default to Column AC
      if (headerRowIndex < 0) headerRowIndex = 0;
      
      // Start looping from the row immediately after the header
      for (let j = headerRowIndex + 1; j < sizeTypeValues.length; j++) {
        const r = sizeTypeValues[j];
        if (r.length > Math.max(skuColIdx, factorColIdx)) {
          const skuCodeVal = r[skuColIdx] ? r[skuColIdx].toString().trim().toUpperCase().replace(/^0+/, '') : '';
          const caseCountVal = Number(r[factorColIdx]) || 0;
          if (skuCodeVal && caseCountVal > 0) {
            caseCountLookup[skuCodeVal] = caseCountVal;
          }
        }
      }
    }
  }

  const newPlanRows = [];
  let syncCount = 0;

  for (let i = 1; i < rawPlanValues.length; i++) {
    const row = rawPlanValues[i];
    const area = row[areaIdx].toString().trim().toUpperCase();
    if (area === 'PKG' || area === 'SE') {
      syncCount++;
      const planId = 'PLN_SYNC_' + syncCount;
      const prodDate = row[dateIdx];
      const sku = row[skuIdx].toString().trim();
      const desc = row[descIdx] ? row[descIdx].toString().trim() : '';
      const machine = row[machineIdx] ? row[machineIdx].toString().trim() : '';
      const shift = row[shiftIdx] || 1;
      
      const rawOutputVal = Number(row[outputIdx]) || 0;
      const cleanSkuKey = sku.toString().trim().toUpperCase().replace(/^0+/, '');
      let caseCount = caseCountLookup[cleanSkuKey] || 1;
      if (caseCount <= 0) caseCount = 1;
      const origOutput = Math.round((rawOutputVal / caseCount) * 100) / 100;
      
      const currOutput = existingOutputs.hasOwnProperty(planId) ? existingOutputs[planId] : origOutput;

      newPlanRows.push([
        planId,
        prodDate,
        area,
        sku,
        desc,
        machine,
        shift,
        origOutput,
        currOutput,
        'SYNC',
        'ACTIVE'
      ]);
    }
  }

  const allPlanHeaders = ['Plan ID', 'Production Date', 'Physical Area', 'Material / SKU Code', 'Material Description', 'Machine', 'Shift', 'Original Output (CS)', 'Current Output (CS)', 'Source', 'Status'];
  const allRowsToWrite = [allPlanHeaders].concat(newPlanRows).concat(adhocRows);
  helperPlanSheet.clearContents();
  helperPlanSheet.getRange(1, 1, allRowsToWrite.length, allRowsToWrite[0].length).setValues(allRowsToWrite);

  // 2. Sync ZMNU to HELPER_BOM_MASTER
  const zmnuSheet = ss.getSheetByName('ZMNU');
  if (!zmnuSheet) {
    throw new Error("Không tìm thấy sheet 'ZMNU'!");
  }

  const rawZmnuValues = zmnuSheet.getDataRange().getValues();
  if (rawZmnuValues.length <= 1) return;

  const zmnuHeaders = rawZmnuValues[0];
  const matIdx = zmnuHeaders.indexOf('Material');
  const matDescIdx = zmnuHeaders.indexOf('Material Description');
  const compIdx = zmnuHeaders.indexOf('Component');
  const compDescIdx = zmnuHeaders.indexOf('Component Description');
  const qtyIdx = zmnuHeaders.indexOf('Quantity');
  const uomIdx = zmnuHeaders.indexOf('UOM');

  if (matIdx < 0 || compIdx < 0 || qtyIdx < 0) {
    throw new Error("Cấu trúc sheet 'ZMNU' không đúng!");
  }

  const newBomRows = [];
  for (let i = 1; i < rawZmnuValues.length; i++) {
    const row = rawZmnuValues[i];
    const compCode = row[compIdx].toString().trim();
    
    if (compCode.toUpperCase().indexOf('P') === 0) {
      const skuCode = row[matIdx].toString().trim();
      const skuDesc = row[matDescIdx] ? row[matDescIdx].toString().trim() : '';
      const compDesc = row[compDescIdx] ? row[compDescIdx].toString().trim() : '';
      const qty = Number(row[qtyIdx]) || 0;
      const uom = row[uomIdx] ? row[uomIdx].toString().trim() : '';

      newBomRows.push([
        skuCode,
        skuDesc,
        compCode,
        compDesc,
        qty,
        uom
      ]);
    }
  }

  const helperBomSheet = ss.getSheetByName(SHEET_HELPER_BOM_MASTER);
  const bomHeaders = ['SKU Code', 'SKU Description', 'Component Code', 'Component Description', 'Base Quantity per 1000 CS', 'UOM'];
  
  helperBomSheet.clearContents();
  const bomDataToWrite = [bomHeaders].concat(newBomRows);
  helperBomSheet.getRange(1, 1, bomDataToWrite.length, bomDataToWrite[0].length).setValues(bomDataToWrite);

  const timestamp = new Date();
  const details = `Triggered full master database sync: Synced ${syncCount} plans, ${newBomRows.length} BOM master rows. By: ${activeUserEmail}.`;
  _writeToAuditLog(timestamp, activeUserEmail, 'DATABASE_SYNC', 'SYSTEM', 'SYNC', '', '', details);

  return {
    success: true,
    syncedPlans: syncCount,
    syncedBoms: newBomRows.length
  };
}

/**
 * Retrieve plans with live-scaled materials, adjustments, and timelines
 */
function getHelperPlanData() {
  setupHelperPlannerSheets();
  
  // Auto-sync database from raw source worksheets on load to ensure fresh data
  try {
    syncHelperPlannerMasterData();
  } catch (syncErr) {
    console.error("Auto-sync on load skipped or failed: " + syncErr.message);
  }
  
  const helperPlanSheet = ss.getSheetByName(SHEET_HELPER_PLAN);
  const planValues = helperPlanSheet.getDataRange().getValues();
  
  const helperBomSheet = ss.getSheetByName(SHEET_HELPER_BOM_MASTER);
  const bomValues = helperBomSheet.getDataRange().getValues();
  
  const helperAdjSheet = ss.getSheetByName(SHEET_HELPER_BOM_ADJUSTMENTS);
  const adjValues = helperAdjSheet.getDataRange().getValues();

  // 1. Verify User Privilege
  const activeUserEmail = Session.getActiveUser().getEmail() || '';
  let hasBomAdjustPrivilege = false;
  
  const accessSheet = ss.getSheetByName(SHEET_HELPER_BOM_ACCESS);
  if (accessSheet) {
    const accessValues = accessSheet.getDataRange().getValues();
    if (accessValues.length > 1) {
      const emailIdx = accessValues[0].indexOf('Email');
      const statusIdx = accessValues[0].indexOf('Status');
      for (let i = 1; i < accessValues.length; i++) {
        const email = accessValues[i][emailIdx].toString().trim().toLowerCase();
        const status = accessValues[i][statusIdx].toString().trim().toUpperCase();
        if (email === activeUserEmail.toLowerCase() && status === 'ACTIVE') {
          hasBomAdjustPrivilege = true;
          break;
        }
      }
    }
  }

  // 2. Parse active adjustments using the upgraded original/adjusted schema
  const adjMap = {};
  const futureOverrides = [];
  if (adjValues.length > 1) {
    const adjHeaders = adjValues[0];
    const versionIdx = adjHeaders.indexOf('BOM Version');
    const skuIdx = adjHeaders.indexOf('SKU Code');
    const origCompIdx = adjHeaders.indexOf('Original Component Code');
    const adjCompIdx = adjHeaders.indexOf('Adjusted Component Code');
    const adjQtyIdx = adjHeaders.indexOf('Adjusted Quantity');
    const statusIdx = adjHeaders.indexOf('Status');
    const uomIdx = adjHeaders.indexOf('UOM');
    const effectiveAtIdx = adjHeaders.indexOf('Effective At');

    for (let i = 1; i < adjValues.length; i++) {
      const row = adjValues[i];
      if (row[statusIdx] === 'ACTIVE') {
        // Scheduled activation check
        if (effectiveAtIdx >= 0 && row[effectiveAtIdx]) {
          const effDate = new Date(row[effectiveAtIdx]);
          const nowTime = new Date().getTime();
          if (!isNaN(effDate.getTime()) && nowTime < effDate.getTime()) {
            futureOverrides.push({
              sku: row[skuIdx].toString().trim(),
              origComp: row[origCompIdx].toString().trim(),
              effectiveAt: row[effectiveAtIdx].toString(),
              epoch: effDate.getTime()
            });
            continue; // Future scheduled override, ignore for now!
          }
        }
        
        const sku = row[skuIdx].toString().trim();
        const origComp = row[origCompIdx].toString().trim();
        const key = sku + '_' + origComp;
        adjMap[key] = {
          adjustedCompCode: row[adjCompIdx] ? row[adjCompIdx].toString().trim() : origComp,
          adjustedQty: Number(row[adjQtyIdx]) || 0,
          version: row[versionIdx],
          uom: row[uomIdx],
          effectiveAt: effectiveAtIdx >= 0 && row[effectiveAtIdx] ? row[effectiveAtIdx].toString() : ''
        };
      }
    }
  }

  // 3. Parse Master BOM
  const bomMap = {};
  if (bomValues.length > 1) {
    const bomHeaders = bomValues[0];
    const skuIdx = bomHeaders.indexOf('SKU Code');
    const skuDescIdx = bomHeaders.indexOf('SKU Description');
    const compIdx = bomHeaders.indexOf('Component Code');
    const compDescIdx = bomHeaders.indexOf('Component Description');
    const qtyIdx = bomHeaders.indexOf('Base Quantity per 1000 CS');
    const uomIdx = bomHeaders.indexOf('UOM');

    for (let i = 1; i < bomValues.length; i++) {
      const row = bomValues[i];
      const sku = row[skuIdx].toString().trim();
      if (!bomMap[sku]) bomMap[sku] = [];
      bomMap[sku].push({
        compCode: row[compIdx].toString().trim(),
        compDesc: row[compDescIdx] ? row[compDescIdx].toString().trim() : '',
        baseQty: Number(row[qtyIdx]) || 0,
        uom: row[uomIdx] ? row[uomIdx].toString().trim() : ''
      });
    }
  }

  // 4. Parse Plan Rows
  const plans = [];
  if (planValues.length > 1) {
    const planHeaders = planValues[0];
    const idIdx = planHeaders.indexOf('Plan ID');
    const dateIdx = planHeaders.indexOf('Production Date');
    const areaIdx = planHeaders.indexOf('Physical Area');
    const skuIdx = planHeaders.indexOf('Material / SKU Code');
    const descIdx = planHeaders.indexOf('Material Description');
    const machineIdx = planHeaders.indexOf('Machine');
    const shiftIdx = planHeaders.indexOf('Shift');
    const origQtyIdx = planHeaders.indexOf('Original Output (CS)');
    const currQtyIdx = planHeaders.indexOf('Current Output (CS)');
    const sourceIdx = planHeaders.indexOf('Source');
    const statusIdx = planHeaders.indexOf('Status');

    for (let i = 1; i < planValues.length; i++) {
      const row = planValues[i];
      if (row[statusIdx] === 'ACTIVE') {
        const planId = row[idIdx];
        const dateVal = row[dateIdx];
        const dateStr = dateVal instanceof Date ? Utilities.formatDate(dateVal, LOCAL_TIME_ZONE, 'yyyy-MM-dd') : dateVal.toString();
        const sku = row[skuIdx].toString().trim();
        const area = row[areaIdx];
        const desc = row[descIdx];
        const machine = row[machineIdx];
        const shift = row[shiftIdx];
        const origQty = Number(row[origQtyIdx]) || 0;
        const currQty = Number(row[currQtyIdx]) || 0;
        const source = row[sourceIdx];

        const rawBoms = bomMap[sku] || [];
        const scaledBoms = rawBoms.map(bom => {
          const adjKey = sku + '_' + bom.compCode;
          const hasOverride = adjMap.hasOwnProperty(adjKey);
          
          const activeCompCode = hasOverride ? adjMap[adjKey].adjustedCompCode : bom.compCode;
          const activeQty = hasOverride ? adjMap[adjKey].adjustedQty : bom.baseQty;
          const requiredQty = (currQty * activeQty) / 1000;
          
          return {
            compCode: activeCompCode,
            origCompCode: bom.compCode,
            compDesc: bom.compDesc,
            baseQty: bom.baseQty,
            activeQty: activeQty,
            uom: bom.uom,
            requiredQty: requiredQty,
            hasOverride: hasOverride,
            overrideQty: hasOverride ? adjMap[adjKey].adjustedQty : null,
            overrideCompCode: hasOverride ? adjMap[adjKey].adjustedCompCode : null,
            version: hasOverride ? adjMap[adjKey].version : null
          };
        });

        plans.push({
          planId: planId,
          date: dateStr,
          area: area,
          sku: sku,
          desc: desc,
          machine: machine,
          shift: shift,
          origQty: origQty,
          currQty: currQty,
          source: source,
          materials: scaledBoms
        });
      }
    }
  }

  // Sort plans by Machine, and secondary by Source (SYNC first, ADHOC second)
  plans.sort((a, b) => {
    const macA = (a.machine || '').toUpperCase();
    const macB = (b.machine || '').toUpperCase();
    if (macA !== macB) return macA.localeCompare(macB);
    
    const srcA = a.source || '';
    const srcB = b.source || '';
    if (srcA === 'SYNC' && srcB !== 'SYNC') return -1;
    if (srcA !== 'SYNC' && srcB === 'SYNC') return 1;
    return 0;
  });

  // 4. Parse Audit Logs (recent 100 entries)
  const auditLogs = [];
  const helperAuditSheet = ss.getSheetByName(SHEET_HELPER_AUDIT_LOG);
  const auditValues = helperAuditSheet.getDataRange().getValues();
  if (auditValues.length > 1) {
    const headers = auditValues[0];
    const tsIdx = headers.indexOf('Timestamp');
    const userIdx = headers.indexOf('User');
    const catIdx = headers.indexOf('Category');
    const targetIdx = headers.indexOf('Target ID');
    const actionIdx = headers.indexOf('Action');
    const oldIdx = headers.indexOf('Old Value');
    const newIdx = headers.indexOf('New Value');
    const detailsIdx = headers.indexOf('Details');

    const limit = Math.max(1, auditValues.length - 100);
    for (let i = auditValues.length - 1; i >= limit; i--) {
      const row = auditValues[i];
      const tsVal = row[tsIdx];
      const tsStr = tsVal instanceof Date ? Utilities.formatDate(tsVal, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss') : tsVal.toString();
      auditLogs.push({
        timestamp: tsStr,
        user: row[userIdx],
        category: row[catIdx],
        targetId: row[targetIdx],
        action: row[actionIdx],
        oldValue: row[oldIdx],
        newValue: row[newIdx],
        details: row[detailsIdx]
      });
    }
  }

  // Determine current active version
  let globalVersion = 'V_1.0.0';
  if (adjValues.length > 1) {
    const versionIdx = adjValues[0].indexOf('BOM Version');
    for (let i = adjValues.length - 1; i >= 1; i--) {
      if (adjValues[i][versionIdx]) {
        globalVersion = adjValues[i][versionIdx].toString();
        break;
      }
    }
  }

  return {
    plans: plans,
    auditLogs: auditLogs,
    latestAuditLog: auditLogs.length > 0 ? auditLogs[0] : null,
    globalVersion: globalVersion,
    hasBomAdjustPrivilege: hasBomAdjustPrivilege,
    userEmail: activeUserEmail,
    futureOverrides: futureOverrides
  };
}

/**
 * Edit Output (CS) in real-time from inline grid
 */
function updatePlanOutput(planId, newOutput) {
  setupHelperPlannerSheets();
  const sheet = ss.getSheetByName(SHEET_HELPER_PLAN);
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  
  const idIdx = headers.indexOf('Plan ID');
  const currQtyIdx = headers.indexOf('Current Output (CS)');
  const origQtyIdx = headers.indexOf('Original Output (CS)');
  const skuIdx = headers.indexOf('Material / SKU Code');
  const areaIdx = headers.indexOf('Physical Area');
  const machineIdx = headers.indexOf('Machine');
  const shiftIdx = headers.indexOf('Shift');

  let foundRowIndex = -1;
  let oldVal = 0;
  let sku = '';
  let area = '';
  let machine = '';
  let shift = '';

  for (let i = 1; i < values.length; i++) {
    if (values[i][idIdx] === planId) {
      foundRowIndex = i;
      oldVal = values[i][currQtyIdx];
      sku = values[i][skuIdx];
      area = values[i][areaIdx];
      machine = values[i][machineIdx];
      shift = values[i][shiftIdx];
      break;
    }
  }

  if (foundRowIndex < 0) {
    throw new Error("Không tìm thấy dòng kế hoạch " + planId);
  }

  const newOutputNum = Number(newOutput);
  if (isNaN(newOutputNum) || newOutputNum < 0) {
    throw new Error("Sản lượng dự kiến phải là số không âm!");
  }

  sheet.getRange(foundRowIndex + 1, currQtyIdx + 1).setValue(newOutputNum);

  const user = Session.getActiveUser().getEmail() || 'Planner';
  const timestamp = new Date();
  const details = `User ${user} changed output of SKU ${sku} (${area}, Machine ${machine}, Shift ${shift}) from ${oldVal} to ${newOutputNum} CS`;
  
  _writeToAuditLog(timestamp, user, 'PLAN_OUTPUT', planId, 'UPDATE', oldVal.toString(), newOutputNum.toString(), details);

  return {
    success: true,
    message: `Đã cập nhật sản lượng SKU ${sku} thành ${newOutputNum} CS`
  };
}

/**
 * Add manual Ad-hoc planning row
 */
function addAdhocPlanRow(rowData) {
  setupHelperPlannerSheets();
  const sheet = ss.getSheetByName(SHEET_HELPER_PLAN);
  
  const user = Session.getActiveUser().getEmail() || 'Planner';
  const timestamp = new Date();
  const adhocId = 'PLN_ADHOC_' + timestamp.getTime();
  
  let prodDate;
  if (rowData.date) {
    prodDate = new Date(rowData.date.replace(/-/g, "/") + " 12:00:00");
  } else {
    prodDate = new Date();
    prodDate.setHours(12, 0, 0, 0);
  }
  const area = rowData.area || 'PKG';
  const sku = rowData.sku ? rowData.sku.toString().trim() : '';
  const desc = rowData.desc ? rowData.desc.toString().trim() : 'Ad-hoc SKU';
  const machine = rowData.machine ? rowData.machine.toString().trim() : '';
  const shift = Number(rowData.shift) || 1;
  const output = Number(rowData.output) || 0;

  if (!sku) {
    throw new Error("SKU Code không được để trống!");
  }

  const row = [
    adhocId,
    prodDate,
    area,
    sku,
    desc,
    machine,
    shift,
    output,
    output,
    'ADHOC',
    'ACTIVE'
  ];

  sheet.appendRow(row);

  const details = `Created ad-hoc plan: SKU ${sku} (${area}, Machine ${machine}, Shift ${shift}) with output ${output} CS`;
  _writeToAuditLog(timestamp, user, 'PLAN_OUTPUT', adhocId, 'CREATE', '', output.toString(), details);

  return {
    success: true,
    message: `Đã thêm ad-hoc plan cho SKU ${sku} thành công!`
  };
}

/**
 * Adjust/Override BOM item base ratio
 */
function adjustBOM(skuCode, origComponentCode, adjustedComponentCode, adjustedQty, comment, effectiveAt) {
  setupHelperPlannerSheets();
  
  // 1. Verify User Privilege
  const activeUserEmail = Session.getActiveUser().getEmail() || '';
  let hasBomAdjustPrivilege = false;
  
  const accessSheet = ss.getSheetByName(SHEET_HELPER_BOM_ACCESS);
  if (accessSheet) {
    const accessValues = accessSheet.getDataRange().getValues();
    if (accessValues.length > 1) {
      const emailIdx = accessValues[0].indexOf('Email');
      const statusIdx = accessValues[0].indexOf('Status');
      for (let i = 1; i < accessValues.length; i++) {
        const email = accessValues[i][emailIdx].toString().trim().toLowerCase();
        const status = accessValues[i][statusIdx].toString().trim().toUpperCase();
        if (email === activeUserEmail.toLowerCase() && status === 'ACTIVE') {
          hasBomAdjustPrivilege = true;
          break;
        }
      }
    }
  }
  
  if (!hasBomAdjustPrivilege) {
    throw new Error("Bạn không có quyền chỉnh sửa định mức BOM! Vui lòng liên hệ Admin.");
  }

  const sheet = ss.getSheetByName(SHEET_HELPER_BOM_ADJUSTMENTS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  const versionIdx = headers.indexOf('BOM Version');
  const skuIdx = headers.indexOf('SKU Code');
  const origCompIdx = headers.indexOf('Original Component Code');
  const adjCompIdx = headers.indexOf('Adjusted Component Code');
  const origQtyIdx = headers.indexOf('Original Quantity');
  const adjQtyIdx = headers.indexOf('Adjusted Quantity');
  const uomIdx = headers.indexOf('UOM');
  const userIdx = headers.indexOf('Adjusted By');
  const atIdx = headers.indexOf('Adjusted At');
  const commentIdx = headers.indexOf('Comment');
  const statusIdx = headers.indexOf('Status');
  const effectiveAtIdx = headers.indexOf('Effective At') >= 0 ? headers.indexOf('Effective At') : headers.length;

  const timestamp = new Date();
  const formattedTimestamp = Utilities.formatDate(timestamp, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');

  const masterSheet = ss.getSheetByName(SHEET_HELPER_BOM_MASTER);
  const masterValues = masterSheet.getDataRange().getValues();
  const mSkuIdx = masterValues[0].indexOf('SKU Code');
  const mCompIdx = masterValues[0].indexOf('Component Code');
  const mQtyIdx = masterValues[0].indexOf('Base Quantity per 1000 CS');
  const mUomIdx = masterValues[0].indexOf('UOM');

  let originalQty = 0;
  let uom = 'PC';
  for (let i = 1; i < masterValues.length; i++) {
    if (masterValues[i][mSkuIdx].toString().trim() === skuCode && masterValues[i][mCompIdx].toString().trim() === origComponentCode) {
      originalQty = Number(masterValues[i][mQtyIdx]) || 0;
      uom = masterValues[i][mUomIdx];
      break;
    }
  }

  let formattedEffectiveAt = '';
  if (effectiveAt) {
    const effDate = new Date(effectiveAt);
    if (!isNaN(effDate.getTime())) {
      formattedEffectiveAt = Utilities.formatDate(effDate, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
    }
  }

  const isFutureScheduled = formattedEffectiveAt && (new Date(formattedEffectiveAt).getTime() > new Date().getTime());

  // Optimize: Batch edit active overrides to SUPERSEDED using bulk setValues
  let currentActiveOverrideVal = originalQty.toString();
  let currentActiveOverrideComp = origComponentCode;
  let statusValues = [];
  let supersededCount = 0;
  
  // Only supersede current active overrides if the new change is immediate (not scheduled in future)
  if (!isFutureScheduled && values.length > 1) {
    statusValues = sheet.getRange(2, statusIdx + 1, values.length - 1, 1).getValues();
    for (let i = 1; i < values.length; i++) {
      const rowSku = values[i][skuIdx].toString().trim();
      const rowOrigComp = values[i][origCompIdx].toString().trim();
      const rowStatus = values[i][statusIdx].toString().trim();
      const rowEffective = headers.indexOf('Effective At') >= 0 ? values[i][headers.indexOf('Effective At')] : '';
      
      if (rowSku === skuCode && rowOrigComp === origComponentCode && rowStatus === 'ACTIVE') {
        const isRowFuture = rowEffective && (new Date(rowEffective).getTime() > new Date().getTime());
        if (!isRowFuture) {
          statusValues[i - 1][0] = 'SUPERSEDED';
          currentActiveOverrideVal = values[i][adjQtyIdx].toString();
          currentActiveOverrideComp = values[i][adjCompIdx] ? values[i][adjCompIdx].toString() : rowOrigComp;
          supersededCount++;
        }
      }
    }
    if (supersededCount > 0) {
      sheet.getRange(2, statusIdx + 1, statusValues.length, 1).setValues(statusValues);
    }
  }

  let major = 1, minor = 0, patch = 0;
  for (let i = 1; i < values.length; i++) {
    const vStr = values[i][versionIdx].toString();
    const vMatch = vStr.match(/V_(\d+)\.(\d+)\.(\d+)/);
    if (vMatch) {
      const vMaj = Number(vMatch[1]), vMin = Number(vMatch[2]), vPat = Number(vMatch[3]);
      if (vMaj > major || (vMaj === major && vMin > minor) || (vMaj === major && vMin === minor && vPat > patch)) {
        major = vMaj;
        minor = vMin;
        patch = vPat;
      }
    }
  }
  
  patch++;
  const nextVersion = `V_${major}.${minor}.${patch}`;

  const row = new Array(Math.max(headers.length, 12)).fill('');
  row[versionIdx] = nextVersion;
  row[skuIdx] = skuCode;
  row[origCompIdx] = origComponentCode;
  row[adjCompIdx] = adjustedComponentCode;
  row[origQtyIdx] = originalQty;
  row[adjQtyIdx] = Number(adjustedQty);
  row[uomIdx] = uom;
  row[userIdx] = activeUserEmail;
  row[atIdx] = formattedTimestamp;
  row[commentIdx] = comment || '';
  row[statusIdx] = 'ACTIVE';
  row[effectiveAtIdx] = formattedEffectiveAt;

  sheet.appendRow(row);

  const schedMsg = formattedEffectiveAt ? ` (Hiệu lực từ: ${formattedEffectiveAt})` : ' (Instant)';
  const details = `Overrode BOM SKU ${skuCode} (${origComponentCode}): Changed P-Code ${currentActiveOverrideComp}->${adjustedComponentCode}, Ratio ${currentActiveOverrideVal}->${adjustedQty} per 1000 CS${schedMsg}. By: ${activeUserEmail}. Comment: ${comment || 'none'}`;
  _writeToAuditLog(timestamp, activeUserEmail, 'BOM_ADJUSTMENT', `${skuCode}_${origComponentCode}`, 'OVERRIDE', `${currentActiveOverrideComp}:${currentActiveOverrideVal}`, `${adjustedComponentCode}:${adjustedQty}`, details);

  return {
    success: true,
    message: `Đã điều chỉnh BOM thành công! Phiên bản mới: ${nextVersion}`
  };
}

/**
 * Reset/Revert custom BOM overrides back to original default Master ZMNU
 */
function resetBOM(skuCode, origComponentCode) {
  setupHelperPlannerSheets();
  
  // 1. Verify User Privilege
  const activeUserEmail = Session.getActiveUser().getEmail() || '';
  let hasBomAdjustPrivilege = false;
  
  const accessSheet = ss.getSheetByName(SHEET_HELPER_BOM_ACCESS);
  if (accessSheet) {
    const accessValues = accessSheet.getDataRange().getValues();
    if (accessValues.length > 1) {
      const emailIdx = accessValues[0].indexOf('Email');
      const statusIdx = accessValues[0].indexOf('Status');
      for (let i = 1; i < accessValues.length; i++) {
        const email = accessValues[i][emailIdx].toString().trim().toLowerCase();
        const status = accessValues[i][statusIdx].toString().trim().toUpperCase();
        if (email === activeUserEmail.toLowerCase() && status === 'ACTIVE') {
          hasBomAdjustPrivilege = true;
          break;
        }
      }
    }
  }
  
  if (!hasBomAdjustPrivilege) {
    throw new Error("Bạn không có quyền khôi phục định mức BOM! Vui lòng liên hệ Admin.");
  }

  const sheet = ss.getSheetByName(SHEET_HELPER_BOM_ADJUSTMENTS);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { success: true, message: "Không tìm thấy điều chỉnh nào cần khôi phục." };
  }
  
  const headers = values[0];
  const skuIdx = headers.indexOf('SKU Code');
  const origCompIdx = headers.indexOf('Original Component Code');
  const statusIdx = headers.indexOf('Status');

  let supersededCount = 0;
  const statusValues = sheet.getRange(2, statusIdx + 1, values.length - 1, 1).getValues();

  for (let i = 1; i < values.length; i++) {
    const rowSku = values[i][skuIdx].toString().trim();
    const rowOrigComp = values[i][origCompIdx].toString().trim();
    const rowStatus = values[i][statusIdx].toString().trim();
    
    if (rowSku === skuCode && rowOrigComp === origComponentCode && rowStatus === 'ACTIVE') {
      statusValues[i - 1][0] = 'SUPERSEDED';
      supersededCount++;
    }
  }

  if (supersededCount > 0) {
    sheet.getRange(2, statusIdx + 1, statusValues.length, 1).setValues(statusValues);
  }

  const timestamp = new Date();
  const details = `Reverted custom override for SKU ${skuCode} (${origComponentCode}) back to master default. By: ${activeUserEmail}.`;
  _writeToAuditLog(timestamp, activeUserEmail, 'BOM_REVERSION', `${skuCode}_${origComponentCode}`, 'REVERT', '', '', details);

  return {
    success: true,
    message: `Đã khôi phục định mức gốc thành công cho SKU ${skuCode}!`
  };
}

/**
 * Common Audit Logging Helper
 */
function _writeToAuditLog(timestamp, user, category, targetId, action, oldValue, newValue, details) {
  setupHelperPlannerSheets();
  const sheet = ss.getSheetByName(SHEET_HELPER_AUDIT_LOG);
  const formattedTimestamp = Utilities.formatDate(timestamp, LOCAL_TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([formattedTimestamp, user, category, targetId, action, oldValue, newValue, details]);
}

/**
 * Diagnostic test function for the Helper Planner
 * Planners can run this function from the GAS Script Editor to verify the backend data engine.
 */
function testHelperPlanner() {
  Logger.log("=== STARTING HELPER PLANNER DIAGNOSTIC TEST ===");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Check Source Sheets
  const planInputSheet = ss.getSheetByName('Planning input');
  const zmnuSheet = ss.getSheetByName('ZMNU');
  
  if (!planInputSheet) {
    Logger.log("❌ ERROR: Sheet 'Planning input' not found!");
  } else {
    const lastRow = planInputSheet.getLastRow();
    Logger.log(`✅ 'Planning input' found with ${lastRow} rows.`);
    
    const values = planInputSheet.getDataRange().getValues();
    const headers = values[0];
    const areaIdx = headers.indexOf('Physical Area');
    const dateIdx = headers.indexOf('Production Date');
    const skuIdx = headers.indexOf('Material / SKU Code');
    
    Logger.log(`Headers found: ${headers.join(", ")}`);
    Logger.log(`Column indexes -> Area: ${areaIdx}, Date: ${dateIdx}, SKU: ${skuIdx}`);
    
    let pkgCount = 0;
    let seCount = 0;
    let otherCount = 0;
    const sampleRows = [];
    
    for (let i = 1; i < values.length; i++) {
      const area = values[i][areaIdx] ? values[i][areaIdx].toString().trim().toUpperCase() : '';
      if (area === 'PKG') {
        pkgCount++;
        if (sampleRows.length < 3) sampleRows.push({ row: i + 1, area, date: values[i][dateIdx], sku: values[i][skuIdx] });
      } else if (area === 'SE') {
        seCount++;
        if (sampleRows.length < 3) sampleRows.push({ row: i + 1, area, date: values[i][dateIdx], sku: values[i][skuIdx] });
      } else {
        otherCount++;
      }
    }
    
    Logger.log(`Found Area counts in raw sheet -> PKG: ${pkgCount}, SE: ${seCount}, Other (e.g. IM): ${otherCount}`);
    if (sampleRows.length > 0) {
      Logger.log("Sample rows found in source:");
      sampleRows.forEach(r => {
        Logger.log(`  - Row ${r.row}: Area=${r.area}, Date=${r.date}, SKU=${r.sku}`);
      });
    } else {
      Logger.log("⚠️ WARNING: No PKG or SE rows found in 'Planning input' sheet!");
    }
  }
  
  if (!zmnuSheet) {
    Logger.log("❌ ERROR: Sheet 'ZMNU' not found!");
  } else {
    Logger.log(`✅ 'ZMNU' sheet found with ${zmnuSheet.getLastRow()} rows.`);
  }
  
  // 2. Perform Trial Sync
  Logger.log("Running syncHelperPlannerMasterData()...");
  try {
    const syncRes = syncHelperPlannerMasterData();
    Logger.log(`✅ Sync successful! Result: ${JSON.stringify(syncRes)}`);
  } catch (err) {
    Logger.log(`❌ SYNC ERROR: ${err.message}`);
  }
  
  // 3. Verify Database Sheets
  const helperPlanSheet = ss.getSheetByName(SHEET_HELPER_PLAN);
  if (helperPlanSheet) {
    const rows = helperPlanSheet.getLastRow();
    Logger.log(`✅ SHEET_HELPER_PLAN (${SHEET_HELPER_PLAN}) contains ${rows} rows.`);
    if (rows > 1) {
      const firstFew = helperPlanSheet.getRange(1, 1, Math.min(6, rows), helperPlanSheet.getLastColumn()).getValues();
      Logger.log("First few rows of HELPER_PLAN:");
      for (let i = 1; i < firstFew.length; i++) {
        Logger.log(`  Row ${i + 1}: ${firstFew[i].join(" | ")}`);
      }
    }
  } else {
    Logger.log(`❌ ERROR: SHEET_HELPER_PLAN not found!`);
  }
  
  const helperBomSheet = ss.getSheetByName(SHEET_HELPER_BOM_MASTER);
  if (helperBomSheet) {
    Logger.log(`✅ SHEET_BOM_MASTER contains ${helperBomSheet.getLastRow()} rows.`);
  }
  
  Logger.log("=== END OF HELPER PLANNER DIAGNOSTIC TEST ===");
}

/**
 * Reset ALL custom BOM overrides in the entire database, restoring everything to master baseline
 */
function resetAllBOMOverrides() {
  setupHelperPlannerSheets();
  
  // 1. Verify User Privilege
  const activeUserEmail = Session.getActiveUser().getEmail() || '';
  let hasBomAdjustPrivilege = false;
  
  const accessSheet = ss.getSheetByName(SHEET_HELPER_BOM_ACCESS);
  if (accessSheet) {
    const accessValues = accessSheet.getDataRange().getValues();
    if (accessValues.length > 1) {
      const emailIdx = accessValues[0].indexOf('Email');
      const statusIdx = accessValues[0].indexOf('Status');
      for (let i = 1; i < accessValues.length; i++) {
        const email = accessValues[i][emailIdx].toString().trim().toLowerCase();
        const status = accessValues[i][statusIdx].toString().trim().toUpperCase();
        if (email === activeUserEmail.toLowerCase() && status === 'ACTIVE') {
          hasBomAdjustPrivilege = true;
          break;
        }
      }
    }
  }
  
  if (!hasBomAdjustPrivilege) {
    throw new Error("Bạn không có quyền khôi phục toàn bộ định mức BOM! Vui lòng liên hệ Admin.");
  }

  const sheet = ss.getSheetByName(SHEET_HELPER_BOM_ADJUSTMENTS);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { success: true, message: "Không tìm thấy điều chỉnh nào cần khôi phục." };
  }
  
  const headers = values[0];
  const statusIdx = headers.indexOf('Status');
  
  let supersededCount = 0;
  const statusValues = sheet.getRange(2, statusIdx + 1, values.length - 1, 1).getValues();

  for (let i = 1; i < values.length; i++) {
    const rowStatus = values[i][statusIdx].toString().trim();
    if (rowStatus === 'ACTIVE') {
      statusValues[i - 1][0] = 'SUPERSEDED';
      supersededCount++;
    }
  }

  if (supersededCount > 0) {
    sheet.getRange(2, statusIdx + 1, statusValues.length, 1).setValues(statusValues);
  }

  const timestamp = new Date();
  const details = `Reset ALL ${supersededCount} active custom BOM overrides back to original default master. By: ${activeUserEmail}.`;
  _writeToAuditLog(timestamp, activeUserEmail, 'BOM_RESET_ALL', 'ALL_BOM', 'RESET', '', '', details);

  return {
    success: true,
    message: `Đã khôi phục toàn bộ ${supersededCount} định mức gốc thành công!`
  };
}

/**
 * Delete manual Ad-hoc planning row (Admin only)
 */
function deleteAdhocPlanRow(planId) {
  setupHelperPlannerSheets();
  
  // 1. Verify User Privilege (Only Admins allowed)
  const activeUserEmail = Session.getActiveUser().getEmail() || '';
  let hasBomAdjustPrivilege = false;
  
  const accessSheet = ss.getSheetByName(SHEET_HELPER_BOM_ACCESS);
  if (accessSheet) {
    const accessValues = accessSheet.getDataRange().getValues();
    if (accessValues.length > 1) {
      const emailIdx = accessValues[0].indexOf('Email');
      const statusIdx = accessValues[0].indexOf('Status');
      for (let i = 1; i < accessValues.length; i++) {
        const email = accessValues[i][emailIdx].toString().trim().toLowerCase();
        const status = accessValues[i][statusIdx].toString().trim().toUpperCase();
        if (email === activeUserEmail.toLowerCase() && status === 'ACTIVE') {
          hasBomAdjustPrivilege = true;
          break;
        }
      }
    }
  }
  
  if (!hasBomAdjustPrivilege) {
    throw new Error("Bạn không có quyền xóa kế hoạch sản xuất! Chỉ Admin mới được thực hiện thao tác này.");
  }

  const sheet = ss.getSheetByName(SHEET_HELPER_PLAN);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idIdx = headers.indexOf('Plan ID');
  const sourceIdx = headers.indexOf('Source');
  const skuIdx = headers.indexOf('Material / SKU Code');
  const areaIdx = headers.indexOf('Physical Area');
  const machineIdx = headers.indexOf('Machine');
  const shiftIdx = headers.indexOf('Shift');

  let foundRowIndex = -1;
  let sku = '';
  let area = '';
  let machine = '';
  let shift = '';
  let source = '';

  for (let i = 1; i < values.length; i++) {
    if (values[i][idIdx] === planId) {
      foundRowIndex = i;
      sku = values[i][skuIdx];
      area = values[i][areaIdx];
      machine = values[i][machineIdx];
      shift = values[i][shiftIdx];
      source = values[i][sourceIdx];
      break;
    }
  }

  if (foundRowIndex < 0) {
    throw new Error("Không tìm thấy dòng kế hoạch " + planId);
  }

  if (source !== 'ADHOC') {
    throw new Error("Chỉ cho phép xóa dòng kế hoạch ADHOC tự tạo!");
  }

  // Delete the row from the sheet
  sheet.deleteRow(foundRowIndex + 1);

  // Write to Audit Log
  const timestamp = new Date();
  const details = `Deleted ad-hoc plan: SKU ${sku} (${area}, Machine ${machine}, Shift ${shift}) (Plan ID: ${planId})`;
  _writeToAuditLog(timestamp, activeUserEmail, 'PLAN_OUTPUT', planId, 'DELETE', sku, '', details);

  return {
    success: true,
    message: `Đã xóa thành công kế hoạch Ad-hoc cho SKU ${sku}`
  };
}
