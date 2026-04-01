import React, { useState, useRef, useMemo } from 'react';
import { Upload, FileSpreadsheet, FileJson, Play, Download, CheckSquare, Square, AlertCircle, FileText, CheckCircle2, Undo2, Redo2, RotateCcw } from 'lucide-react';
import ExcelJS from 'exceljs';
import Papa from 'papaparse';
import { saveAs } from 'file-saver';
import { motion, AnimatePresence } from 'motion/react';

interface NewRow {
  workflowStep: string;
  alternateName: string;
  referenceId: string;
  normalizedWorkflow: string;
  canonicalKey: string;
  keyIgnoreRef: string;
  keyIgnoreAlt: string;
  keyAltAndRef: string;
}

interface PreviewRow {
  id: number;
  rootRowIndex: number;
  oldName: string;
  newName: string;
  oldAlternate: string;
  newAlternate: string;
  oldReferenceId: string;
  newReferenceId: string;
  changeType: string;
  matchedNewRow: NewRow | null;
  parsedStep: any;
}

interface ValidationResult {
  status: 'PASSED' | 'FAILED';
  totalRowsChecked: number;
  uniqueWorkflowSteps: number;
  duplicateWorkflowSteps: number;
  uniqueReferenceIds: number;
  duplicateReferenceIds: number;
  uniqueWorkflowAndRefIds: number;
  duplicateWorkflowAndRefIds: number;
  workflowDuplicates: Record<string, number[]>;
  referenceIdDuplicates: Record<string, number[]>;
  workflowAndRefIdDuplicates: Record<string, number[]>;
  detailedRecords: {
    type: string;
    value: string;
    rowNumber: number;
    name: string;
    alternateName: string;
    referenceId: string;
  }[];
}

function extractWorkflow(name: string): string {
  const parenMatch = name?.match(/\(([^)]+)\)/);
  const workflowName = parenMatch ? parenMatch[1] : '';

  let extractedText = '';
  const actionMatch = name?.match(/Action\s+(.*)/i);
  const conclusionMatch = name?.match(/Conclusion:\s+(.*)/i);

  if (actionMatch) {
    extractedText = actionMatch[1].trim();
  } else if (conclusionMatch) {
    extractedText = conclusionMatch[1].trim();
  }

  let result = workflowName;
  if (extractedText) {
    result += ` - ${extractedText}`;
  }
  return result.trim();
}

function cleanStr(str: string) {
  return (str || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function getCanonicalKey(norm: string, alt: string, ref: string) {
  return `${cleanStr(norm)} | ${cleanStr(alt)} | ${cleanStr(ref)}`;
}
function getKeyIgnoreRef(norm: string, alt: string) {
  return `${cleanStr(norm)} | ${cleanStr(alt)}`;
}
function getKeyIgnoreAlt(norm: string, ref: string) {
  return `${cleanStr(norm)} | ${cleanStr(ref)}`;
}
function getKeyAltAndRef(alt: string, ref: string) {
  return `${cleanStr(alt)} | ${cleanStr(ref)}`;
}

function cleanAlternateName(name: string): string {
  if (!name) return '';
  const trimmed = name.trim();
  if (trimmed.toLowerCase().startsWith('conclusion:')) {
    return trimmed.substring('conclusion:'.length).trim();
  }
  return trimmed;
}

function parsePossibleNextStep(step: string) {
  if (!step) return null;
  const trimmed = step.trim();
  
  // Try to match the exact format: "String","String",Boolean,"String"
  // This handles commas inside the strings perfectly without breaking.
  const exactMatch = trimmed.match(/^"(.*?)","(.*?)",([^,]*),"(.*)"$/);
  
  if (exactMatch) {
    return {
      val1: exactMatch[1],
      alternateName: cleanAlternateName(exactMatch[2]),
      val3: exactMatch[3],
      referenceId: exactMatch[4],
      originalString: step
    };
  }

  // Fallback to Papa Parse if the format is slightly different
  const parsed = Papa.parse(trimmed, { header: false }).data[0] as string[];
  if (!parsed || parsed.length < 4) return null;
  
  const clean = (val: string) => {
    if (!val) return '';
    let c = val.trim();
    if (c.startsWith('"') && c.endsWith('"')) {
      c = c.substring(1, c.length - 1);
      c = c.replace(/""/g, '"');
    }
    return c;
  };

  return {
    val1: clean(parsed[0]),
    alternateName: cleanAlternateName(clean(parsed[1])),
    val3: clean(parsed[2]),
    referenceId: clean(parsed[3]),
    originalString: step
  };
}

function updateReferenceIdInStep(originalStep: string, newRefId: string) {
  if (!originalStep) return newRefId;
  const trimmed = originalStep.trim();
  
  // Try to match the exact format: "String","String",Boolean,"String"
  const exactMatch = trimmed.match(/^"(.*?)","(.*?)",([^,]*),"(.*)"$/);
  
  if (exactMatch) {
    // Reconstruct the string keeping the exact format
    return `"${exactMatch[1]}","${exactMatch[2]}",${exactMatch[3]},"${newRefId}"`;
  }
  
  // Fallback to Papa Parse
  const parsedResult = Papa.parse(trimmed, { header: false });
  const parsed = parsedResult.data[0] as string[];
  
  if (!parsed || parsed.length === 0) {
    return `"${(newRefId || '').replace(/"/g, '""')}"`;
  }
  
  // The reference ID is always the last element in the comma-separated list
  parsed[parsed.length - 1] = newRefId;
  
  // Convert back to CSV string safely
  return Papa.unparse([parsed], { header: false }).trim();
}

function canUpdate(changeType: string) {
  return changeType === 'MATCH' || (changeType.endsWith('Change') && changeType !== 'Invalid Format' && changeType !== 'Relationship Missing');
}

export default function App() {
  const [rootFile, setRootFile] = useState<File | null>(null);
  const [newFile, setNewFile] = useState<File | null>(null);
  const [rootFileRows, setRootFileRows] = useState<number | null>(null);
  const [newFileRows, setNewFileRows] = useState<number | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [previewRows, setPreviewRows] = useState<PreviewRow[]>([]);
  const [selectedRowIds, setSelectedRowIds] = useState<Set<number>>(new Set());
  const [workbook, setWorkbook] = useState<ExcelJS.Workbook | null>(null);
  const [cols, setCols] = useState({ nameCol: -1, possibleNextStepCol: -1 });
  const [verificationSummary, setVerificationSummary] = useState<any>(null);
  const [verificationRows, setVerificationRows] = useState<any[]>([]);
  const [validationResult, setValidationResult] = useState<ValidationResult | null>(null);
  const [updatedWorkbookBuffer, setUpdatedWorkbookBuffer] = useState<ArrayBuffer | null>(null);
  const [updatedWorkbookBufferNoDupes, setUpdatedWorkbookBufferNoDupes] = useState<ArrayBuffer | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [activeFilter, setActiveFilter] = useState<string>('All');
  const [selectionHistory, setSelectionHistory] = useState<Set<number>[]>([new Set()]);
  const [historyIndex, setHistoryIndex] = useState<number>(0);

  const rootFileInputRef = useRef<HTMLInputElement>(null);
  const newFileInputRef = useRef<HTMLInputElement>(null);

  const handleReset = () => {
    setRootFile(null);
    setNewFile(null);
    setRootFileRows(null);
    setNewFileRows(null);
    setPreviewRows([]);
    setSelectedRowIds(new Set());
    setWorkbook(null);
    setVerificationSummary(null);
    setVerificationRows([]);
    setValidationResult(null);
    setUpdatedWorkbookBuffer(null);
    setUpdatedWorkbookBufferNoDupes(null);
    setError(null);
    setActiveFilter('All');
    setSelectionHistory([new Set()]);
    setHistoryIndex(0);
    if (rootFileInputRef.current) rootFileInputRef.current.value = '';
    if (newFileInputRef.current) newFileInputRef.current.value = '';
  };

  const handleRootFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    setRootFile(file);
    if (file) {
      try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await file.arrayBuffer());
        const ws = wb.worksheets[0];
        let count = 0;
        ws.eachRow((row, rowNumber) => {
          if (rowNumber > 9) count++;
        });
        setRootFileRows(count);
      } catch (err) {
        console.error(err);
        setRootFileRows(0);
      }
    } else {
      setRootFileRows(null);
    }
  };

  const handleNewFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    setNewFile(file);
    if (file) {
      try {
        const text = await file.text();
        if (file.name.endsWith('.csv')) {
          const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
          setNewFileRows(parsed.data.length);
        } else if (file.name.endsWith('.json')) {
          const parsed = JSON.parse(text);
          setNewFileRows(parsed.length);
        }
      } catch (err) {
        console.error(err);
        setNewFileRows(0);
      }
    } else {
      setNewFileRows(null);
    }
  };

  const handleProcess = async () => {
    if (!rootFile || !newFile) return;
    setIsProcessing(true);
    setError(null);
    setVerificationSummary(null);
    setVerificationRows([]);

    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(await rootFile.arrayBuffer());
      const ws = wb.worksheets[0];

      const headerRow = ws.getRow(9);
      let nameCol = -1;
      let possibleNextStepCol = -1;

      headerRow.eachCell((cell, colNumber) => {
        if (cell.value === 'NAME') nameCol = colNumber;
        if (cell.value === 'POSSIBLE NEXT STEP') possibleNextStepCol = colNumber;
      });

      if (nameCol === -1 || possibleNextStepCol === -1) {
        throw new Error("Could not find required columns 'NAME' and 'POSSIBLE NEXT STEP' in row 9 of the ROOT XLSX file.");
      }

      setCols({ nameCol, possibleNextStepCol });

      let newRowsData: any[] = [];
      if (newFile.name.endsWith('.csv')) {
        const text = await newFile.text();
        const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
        newRowsData = parsed.data;
      } else if (newFile.name.endsWith('.json')) {
        const text = await newFile.text();
        newRowsData = JSON.parse(text);
      } else {
        throw new Error("Workflow Data File must be a CSV or JSON file.");
      }

      const newRows: NewRow[] = newRowsData.map(row => {
        const norm = extractWorkflow(row.Workflow_Step || '');
        const cleanedAlternateName = cleanAlternateName(row.Workflow_Step_Alternate_Name || '');
        return {
          workflowStep: row.Workflow_Step || '',
          alternateName: cleanedAlternateName,
          referenceId: row.referenceID || '',
          normalizedWorkflow: norm,
          canonicalKey: getCanonicalKey(norm, cleanedAlternateName, row.referenceID),
          keyIgnoreRef: getKeyIgnoreRef(norm, cleanedAlternateName),
          keyIgnoreAlt: getKeyIgnoreAlt(norm, row.referenceID),
          keyAltAndRef: getKeyAltAndRef(cleanedAlternateName, row.referenceID)
        };
      });

      const newByCanonical = new Map<string, NewRow>();
      const newByIgnoreRef = new Map<string, NewRow>();
      const newByIgnoreAlt = new Map<string, NewRow>();
      const newByAltAndRef = new Map<string, NewRow>();

      newRows.forEach(row => {
        newByCanonical.set(row.canonicalKey, row);
        newByIgnoreRef.set(row.keyIgnoreRef, row);
        newByIgnoreAlt.set(row.keyIgnoreAlt, row);
        newByAltAndRef.set(row.keyAltAndRef, row);
      });

      const pRows: PreviewRow[] = [];
      const matchedNewKeys = new Set<string>();

      ws.eachRow((row, rowNumber) => {
        if (rowNumber < 10) return;

        const name = row.getCell(nameCol).text || '';
        const possibleNextStep = row.getCell(possibleNextStepCol).text || '';

        if (!name && !possibleNextStep) return;

        const parsedStep = parsePossibleNextStep(possibleNextStep);
        let changeType = '';
        let matchedNewRow: NewRow | null = null;

        if (!parsedStep) {
          changeType = 'Invalid Format';
        } else {
          const norm = extractWorkflow(name);
          const canonicalKey = getCanonicalKey(norm, parsedStep.alternateName, parsedStep.referenceId);
          const keyIgnoreRef = getKeyIgnoreRef(norm, parsedStep.alternateName);
          const keyIgnoreAlt = getKeyIgnoreAlt(norm, parsedStep.referenceId);
          const keyAltAndRef = getKeyAltAndRef(parsedStep.alternateName, parsedStep.referenceId);

          if (newByCanonical.has(canonicalKey)) {
            matchedNewRow = newByCanonical.get(canonicalKey)!;
          } else if (newByIgnoreRef.has(keyIgnoreRef)) {
            matchedNewRow = newByIgnoreRef.get(keyIgnoreRef)!;
          } else if (newByIgnoreAlt.has(keyIgnoreAlt)) {
            matchedNewRow = newByIgnoreAlt.get(keyIgnoreAlt)!;
          } else if (newByAltAndRef.has(keyAltAndRef)) {
            matchedNewRow = newByAltAndRef.get(keyAltAndRef)!;
          }

          if (matchedNewRow) {
            const nameChanged = name.trim() !== matchedNewRow.workflowStep.trim();
            const altChanged = parsedStep.alternateName.trim() !== matchedNewRow.alternateName.trim();
            const refChanged = parsedStep.referenceId.trim() !== matchedNewRow.referenceId.trim();

            if (!nameChanged && !altChanged && !refChanged) {
              changeType = 'MATCH';
            } else {
              const changes = [];
              if (nameChanged) changes.push('Workflow');
              if (altChanged) changes.push('Label');
              if (refChanged) changes.push('referenceID');
              changeType = changes.join(' + ') + ' Change';
            }
          } else {
            changeType = 'Relationship Missing';
          }
        }

        if (matchedNewRow) {
          matchedNewKeys.add(matchedNewRow.canonicalKey);
        }

        pRows.push({
          id: rowNumber,
          rootRowIndex: rowNumber,
          oldName: name,
          newName: matchedNewRow?.workflowStep || '',
          oldAlternate: parsedStep?.alternateName || '',
          newAlternate: matchedNewRow?.alternateName || '',
          oldReferenceId: parsedStep?.referenceId || '',
          newReferenceId: matchedNewRow?.referenceId || '',
          changeType,
          matchedNewRow,
          parsedStep
        });
      });

      setPreviewRows(pRows);
      setWorkbook(wb);
      setSelectedRowIds(new Set());
      setSelectionHistory([new Set()]);
      setHistoryIndex(0);
    } catch (err: any) {
      setError(err.message || "An error occurred during processing.");
    } finally {
      setIsProcessing(false);
    }
  };

  const counts = useMemo(() => {
    const c: Record<string, number> = {};
    previewRows.forEach(r => {
      c[r.changeType] = (c[r.changeType] || 0) + 1;
    });
    return c;
  }, [previewRows]);

  const filteredRows = useMemo(() => {
    if (activeFilter === 'All') return previewRows;
    return previewRows.filter(r => r.changeType === activeFilter);
  }, [previewRows, activeFilter]);

  const updateSelection = (newSet: Set<number>) => {
    const newHistory = selectionHistory.slice(0, historyIndex + 1);
    newHistory.push(newSet);
    setSelectionHistory(newHistory);
    setHistoryIndex(newHistory.length - 1);
    setSelectedRowIds(newSet);
  };

  const handleSelectVisible = () => {
    const newSet = new Set<number>(selectedRowIds);
    filteredRows.forEach(r => {
      if (canUpdate(r.changeType)) newSet.add(r.id);
    });
    updateSelection(newSet);
  };

  const handleDeselectVisible = () => {
    const newSet = new Set<number>(selectedRowIds);
    filteredRows.forEach(r => {
      newSet.delete(r.id);
    });
    updateSelection(newSet);
  };

  const toggleRowSelection = (id: number) => {
    const newSet = new Set<number>(selectedRowIds);
    if (newSet.has(id)) {
      newSet.delete(id);
    } else {
      newSet.add(id);
    }
    updateSelection(newSet);
  };

  const handleUndo = () => {
    if (historyIndex > 0) {
      const newIndex = historyIndex - 1;
      setHistoryIndex(newIndex);
      setSelectedRowIds(selectionHistory[newIndex]);
    }
  };

  const handleRedo = () => {
    if (historyIndex < selectionHistory.length - 1) {
      const newIndex = historyIndex + 1;
      setHistoryIndex(newIndex);
      setSelectedRowIds(selectionHistory[newIndex]);
    }
  };

  const getRootSuffix = () => {
    if (!rootFile) return '';
    const firstWord = rootFile.name.split(/[\s_\-\.]+/)[0];
    return firstWord ? `_${firstWord}` : '';
  };

  const downloadDifferenceReport = () => {
    const data = previewRows.map(r => ({
      Row: r.rootRowIndex !== -1 ? r.rootRowIndex : 'N/A',
      'Old NAME': r.oldName,
      'New NAME': r.newName,
      'Old AlternateName': r.oldAlternate,
      'New AlternateName': r.newAlternate,
      'Old referenceID': r.oldReferenceId,
      'New referenceID': r.newReferenceId,
      'Change Type': r.changeType
    }));
    const csv = Papa.unparse(data);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    saveAs(blob, `Current_Difference_Report${getRootSuffix()}.csv`);
  };

  const applyUpdatesAndValidate = async () => {
    if (!workbook) return;
    setIsProcessing(true);
    try {
      const ws = workbook.worksheets[0];
      const { nameCol, possibleNextStepCol } = cols;

      const originalValues = new Map<number, { name: string, refId: string, alternate: string }>();

      previewRows.forEach(pr => {
        if (pr.rootRowIndex !== -1) {
          originalValues.set(pr.rootRowIndex, {
            name: pr.oldName,
            refId: pr.oldReferenceId,
            alternate: pr.oldAlternate
          });
        }
      });

      previewRows.forEach(pr => {
        if (!selectedRowIds.has(pr.id)) return;
        if (pr.rootRowIndex === -1) return;
        if (!pr.matchedNewRow || !pr.parsedStep) return;

        const row = ws.getRow(pr.rootRowIndex);

        if (pr.changeType === 'MATCH' || pr.changeType.includes('Workflow')) {
          row.getCell(nameCol).value = pr.matchedNewRow.workflowStep;
        }
        
        if (pr.changeType.includes('referenceID')) {
          const newStep = updateReferenceIdInStep(
            pr.parsedStep.originalString,
            pr.matchedNewRow.referenceId
          );
          row.getCell(possibleNextStepCol).value = newStep;
        }
      });

      const vRows: any[] = [];
      let workflowUpdates = 0;
      let refIdUpdates = 0;
      let bothUpdates = 0;

      const workflowMap = new Map<string, number[]>();
      const referenceMap = new Map<string, number[]>();
      const workflowAndRefMap = new Map<string, number[]>();
      const rowDetails = new Map<number, { name: string, alternateName: string, referenceId: string }>();
      let totalRowsChecked = 0;

      ws.eachRow((row, rowNumber) => {
        if (rowNumber < 10) return;
        const orig = originalValues.get(rowNumber);
        if (!orig) return;

        const newName = row.getCell(nameCol).text || '';
        const newStepStr = row.getCell(possibleNextStepCol).text || '';
        const parsedNewStep = parsePossibleNextStep(newStepStr);
        const newRefId = parsedNewStep?.referenceId || '';
        const newAlternate = parsedNewStep?.alternateName || '';

        const nameChanged = orig.name.trim() !== newName.trim();
        const refChanged = orig.refId.trim() !== newRefId.trim();

        let updateType = 'No Change';
        if (nameChanged && refChanged) {
          updateType = 'Workflow + referenceID Updated';
          bothUpdates++;
        } else if (nameChanged) {
          updateType = 'Workflow Updated';
          workflowUpdates++;
        } else if (refChanged) {
          updateType = 'referenceID Updated';
          refIdUpdates++;
        }

        if (updateType !== 'No Change') {
          vRows.push({
            rowNumber,
            oldName: orig.name,
            newName,
            oldAlternate: orig.alternate,
            newAlternate,
            oldReferenceId: orig.refId,
            newReferenceId: newRefId,
            updateType
          });
        }

        // Validation logic
        totalRowsChecked++;
        rowDetails.set(rowNumber, { name: newName, alternateName: newAlternate, referenceId: newRefId });

        const normName = newName.trim().toLowerCase().replace(/\s+/g, ' ');
        const normRefId = newRefId.trim().toLowerCase();
        const normBoth = `${normName} | ${normRefId}`;

        if (normName) {
          if (!workflowMap.has(normName)) workflowMap.set(normName, []);
          workflowMap.get(normName)!.push(rowNumber);
        }

        if (normRefId) {
          if (!referenceMap.has(normRefId)) referenceMap.set(normRefId, []);
          referenceMap.get(normRefId)!.push(rowNumber);
        }

        if (normName && normRefId) {
          if (!workflowAndRefMap.has(normBoth)) workflowAndRefMap.set(normBoth, []);
          workflowAndRefMap.get(normBoth)!.push(rowNumber);
        }
      });

      const summary = {
        totalCompared: originalValues.size,
        workflowUpdates,
        refIdUpdates,
        bothUpdates,
        totalUpdated: workflowUpdates + refIdUpdates + bothUpdates
      };

      setVerificationRows(vRows);
      setVerificationSummary(summary);

      const workflowDuplicates: Record<string, number[]> = {};
      const referenceIdDuplicates: Record<string, number[]> = {};
      const workflowAndRefIdDuplicates: Record<string, number[]> = {};
      const detailedRecords: any[] = [];
      const combinedDuplicateRows = new Set<number>();

      let duplicateWorkflowAndRefIds = 0;
      let uniqueWorkflowAndRefIds = 0;
      workflowAndRefMap.forEach((rows, normBoth) => {
        if (rows.length > 1) {
          duplicateWorkflowAndRefIds++;
          workflowAndRefIdDuplicates[normBoth] = rows;
          rows.forEach(r => {
            combinedDuplicateRows.add(r);
            const details = rowDetails.get(r)!;
            detailedRecords.push({
              type: 'Workflow + referenceID Duplicate',
              value: `${details.name} | ${details.referenceId}`,
              rowNumber: r,
              name: details.name,
              alternateName: details.alternateName,
              referenceId: details.referenceId
            });
          });
        } else {
          uniqueWorkflowAndRefIds++;
        }
      });

      let duplicateWorkflowSteps = 0;
      let uniqueWorkflowSteps = 0;
      workflowMap.forEach((rows, normName) => {
        if (rows.length > 1) {
          const isFullyCovered = rows.every(r => combinedDuplicateRows.has(r));
          if (!isFullyCovered) {
            duplicateWorkflowSteps++;
            workflowDuplicates[normName] = rows;
            rows.forEach(r => {
              if (!combinedDuplicateRows.has(r)) {
                const details = rowDetails.get(r)!;
                detailedRecords.push({
                  type: 'Workflow Duplicate',
                  value: details.name,
                  rowNumber: r,
                  name: details.name,
                  alternateName: details.alternateName,
                  referenceId: details.referenceId
                });
              }
            });
          }
        } else {
          uniqueWorkflowSteps++;
        }
      });

      let duplicateReferenceIds = 0;
      let uniqueReferenceIds = 0;
      referenceMap.forEach((rows, normRefId) => {
        if (rows.length > 1) {
          const isFullyCovered = rows.every(r => combinedDuplicateRows.has(r));
          if (!isFullyCovered) {
            duplicateReferenceIds++;
            referenceIdDuplicates[normRefId] = rows;
            rows.forEach(r => {
              if (!combinedDuplicateRows.has(r)) {
                const details = rowDetails.get(r)!;
                detailedRecords.push({
                  type: 'referenceID Duplicate',
                  value: details.referenceId,
                  rowNumber: r,
                  name: details.name,
                  alternateName: details.alternateName,
                  referenceId: details.referenceId
                });
              }
            });
          }
        } else {
          uniqueReferenceIds++;
        }
      });

      const status = (duplicateWorkflowSteps === 0 && duplicateReferenceIds === 0 && duplicateWorkflowAndRefIds === 0) ? 'PASSED' : 'FAILED';

      setValidationResult({
        status,
        totalRowsChecked,
        uniqueWorkflowSteps,
        duplicateWorkflowSteps,
        uniqueReferenceIds,
        duplicateReferenceIds,
        uniqueWorkflowAndRefIds,
        duplicateWorkflowAndRefIds,
        workflowDuplicates,
        referenceIdDuplicates,
        workflowAndRefIdDuplicates,
        detailedRecords
      });

      const buffer = await workbook.xlsx.writeBuffer();
      setUpdatedWorkbookBuffer(buffer);

      // Generate buffer without duplicates
      const rowsToRemove = new Set<number>();
      Object.values(workflowDuplicates).forEach(rows => rows.slice(1).forEach(r => rowsToRemove.add(r)));
      Object.values(referenceIdDuplicates).forEach(rows => rows.slice(1).forEach(r => rowsToRemove.add(r)));
      Object.values(workflowAndRefIdDuplicates).forEach(rows => rows.slice(1).forEach(r => rowsToRemove.add(r)));

      if (rowsToRemove.size > 0) {
        const wbNoDupes = new ExcelJS.Workbook();
        await wbNoDupes.xlsx.load(buffer);
        const wsNoDupes = wbNoDupes.worksheets[0];
        const sortedRowsToRemove = Array.from(rowsToRemove).sort((a, b) => b - a);
        for (const r of sortedRowsToRemove) {
          wsNoDupes.spliceRows(r, 1);
        }
        const bufferNoDupes = await wbNoDupes.xlsx.writeBuffer();
        setUpdatedWorkbookBufferNoDupes(bufferNoDupes);
      } else {
        setUpdatedWorkbookBufferNoDupes(buffer);
      }
    } catch (err: any) {
      setError(err.message || "An error occurred during update.");
    } finally {
      setIsProcessing(false);
    }
  };

  const downloadUpdatedFile = (removeDuplicates: boolean) => {
    const buffer = removeDuplicates ? updatedWorkbookBufferNoDupes : updatedWorkbookBuffer;
    if (!buffer) return;
    let originalName = rootFile?.name || 'ROOT.xlsx';
    const suffix = removeDuplicates ? '_updated_no_duplicates.xlsx' : '_updated.xlsx';
    if (originalName.toLowerCase().endsWith('.xlsx')) {
      originalName = originalName.substring(0, originalName.length - 5) + suffix;
    } else {
      originalName += suffix;
    }
    saveAs(new Blob([buffer]), originalName);
  };

  const downloadDuplicateValidationReport = () => {
    if (!validationResult) return;
    const wb = new ExcelJS.Workbook();
    
    // Sheet 1 - Summary & Details
    const wsSummary = wb.addWorksheet('Validation Summary');
    wsSummary.columns = [
      { header: 'Metric / Type', key: 'metric', width: 30 },
      { header: 'Value', key: 'value', width: 40 },
      { header: 'Row Number', key: 'rowNumber', width: 15 },
      { header: 'NAME', key: 'name', width: 40 },
      { header: 'AlternateName', key: 'alternateName', width: 40 },
      { header: 'referenceID', key: 'referenceId', width: 40 }
    ];
    wsSummary.addRows([
      { metric: 'Total Rows Checked', value: validationResult.totalRowsChecked },
      { metric: 'Unique Workflow Steps', value: validationResult.uniqueWorkflowSteps },
      { metric: 'Duplicate Workflow Steps', value: validationResult.duplicateWorkflowSteps },
      { metric: 'Unique referenceIDs', value: validationResult.uniqueReferenceIds },
      { metric: 'Duplicate referenceIDs', value: validationResult.duplicateReferenceIds },
      { metric: 'Unique Workflow + referenceIDs', value: validationResult.uniqueWorkflowAndRefIds },
      { metric: 'Duplicate Workflow + referenceIDs', value: validationResult.duplicateWorkflowAndRefIds },
      { metric: 'Overall Validation Status', value: validationResult.status }
    ]);

    if (validationResult.detailedRecords.length > 0) {
      wsSummary.addRow({});
      wsSummary.addRow({ metric: '--- DETAILED DUPLICATE RECORDS ---' });
      wsSummary.addRow({
        metric: 'Type',
        value: 'Value',
        rowNumber: 'Row Number',
        name: 'NAME',
        alternateName: 'AlternateName',
        referenceId: 'referenceID'
      });
      
      // Style the header row for details
      const headerRow = wsSummary.lastRow;
      if (headerRow) {
        headerRow.font = { bold: true };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
      }

      validationResult.detailedRecords.forEach(record => {
        wsSummary.addRow({
          metric: record.type,
          value: record.value,
          rowNumber: record.rowNumber,
          name: record.name,
          alternateName: record.alternateName,
          referenceId: record.referenceId
        });
      });
    }

    // Sheet 2 - Duplicate Workflow Steps
    const wsWorkflow = wb.addWorksheet('Duplicate Workflow Steps');
    wsWorkflow.columns = [
      { header: 'Workflow Step', key: 'step', width: 40 },
      { header: 'Occurrence Count', key: 'count', width: 20 },
      { header: 'Row Numbers', key: 'rows', width: 40 }
    ];
    Object.entries(validationResult.workflowDuplicates).forEach(([step, rows]) => {
      wsWorkflow.addRow({ step, count: (rows as number[]).length, rows: (rows as number[]).join(', ') });
    });

    // Sheet 3 - Duplicate referenceIDs
    const wsRef = wb.addWorksheet('Duplicate referenceIDs');
    wsRef.columns = [
      { header: 'referenceID', key: 'refId', width: 40 },
      { header: 'Occurrence Count', key: 'count', width: 20 },
      { header: 'Row Numbers', key: 'rows', width: 40 }
    ];
    Object.entries(validationResult.referenceIdDuplicates).forEach(([refId, rows]) => {
      wsRef.addRow({ refId, count: (rows as number[]).length, rows: (rows as number[]).join(', ') });
    });

    // Sheet 4 - Detailed Duplicate Records
    const wsDetails = wb.addWorksheet('Detailed Duplicate Records');
    wsDetails.columns = [
      { header: 'Type', key: 'type', width: 25 },
      { header: 'Value', key: 'value', width: 40 },
      { header: 'Row Number', key: 'rowNumber', width: 15 },
      { header: 'NAME', key: 'name', width: 40 },
      { header: 'AlternateName', key: 'alternateName', width: 40 },
      { header: 'referenceID', key: 'referenceId', width: 40 }
    ];
    wsDetails.addRows(validationResult.detailedRecords);

    // Sheet 4 - Duplicate Workflow + referenceIDs
    const wsBoth = wb.addWorksheet('Duplicate Workflow+refID');
    wsBoth.columns = [
      { header: 'Workflow + referenceID', key: 'both', width: 40 },
      { header: 'Occurrence Count', key: 'count', width: 20 },
      { header: 'Row Numbers', key: 'rows', width: 40 }
    ];
    Object.entries(validationResult.workflowAndRefIdDuplicates).forEach(([both, rows]) => {
      wsBoth.addRow({ both, count: (rows as number[]).length, rows: (rows as number[]).join(', ') });
    });

    wb.xlsx.writeBuffer().then(buffer => {
      saveAs(new Blob([buffer]), `Duplicate_Validation_Report${getRootSuffix()}.xlsx`);
    });
  };

  const downloadVerificationReport = (format: 'csv' | 'xlsx') => {
    const suffix = getRootSuffix();
    if (format === 'csv') {
      const csv = Papa.unparse(verificationRows);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      saveAs(blob, `Post-Update_Verification_Report${suffix}.csv`);
    } else {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Verification');
      ws.columns = [
        { header: 'Row Number', key: 'rowNumber', width: 15 },
        { header: 'Old NAME', key: 'oldName', width: 30 },
        { header: 'Updated NAME', key: 'newName', width: 30 },
        { header: 'Old AlternateName', key: 'oldAlternate', width: 30 },
        { header: 'Updated AlternateName', key: 'newAlternate', width: 30 },
        { header: 'Old referenceID', key: 'oldReferenceId', width: 30 },
        { header: 'Updated referenceID', key: 'newReferenceId', width: 30 },
        { header: 'Update Type', key: 'updateType', width: 30 }
      ];
      ws.addRows(verificationRows);
      wb.xlsx.writeBuffer().then(buffer => {
        saveAs(new Blob([buffer]), `Post-Update_Verification_Report${suffix}.xlsx`);
      });
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-6">
      <div className="max-w-7xl mx-auto space-y-6">
        <header className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-semibold tracking-tight">Workflow Comparison Tool</h1>
            <p className="text-slate-500 mt-1">Compare and update workflow relationships without breaking formatting.</p>
          </div>
        </header>

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 p-4 rounded-xl flex items-start gap-3">
            <AlertCircle className="w-5 h-5 mt-0.5 flex-shrink-0" />
            <p>{error}</p>
          </div>
        )}

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h2 className="text-lg font-medium mb-4 flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
              1. ROOT XLSX File
            </h2>
            <label className="block w-full cursor-pointer">
              <motion.div 
                initial={false}
                animate={{ 
                  backgroundColor: rootFile ? '#ecfdf5' : '#f8fafc',
                  borderColor: rootFile ? '#10b981' : '#cbd5e1'
                }}
                className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-xl transition-colors"
              >
                <div className="flex flex-col items-center justify-center pt-5 pb-6">
                  {rootFile ? (
                    <motion.div initial={{ scale: 0 }} animate={{ scale: 1 }} className="flex flex-col items-center">
                      <CheckCircle2 className="w-8 h-8 mb-2 text-emerald-500" />
                      <p className="text-sm font-medium text-emerald-700">{rootFile.name}</p>
                      {rootFileRows !== null && (
                        <p className="text-xs text-emerald-600 mt-1">{rootFileRows} rows detected</p>
                      )}
                    </motion.div>
                  ) : (
                    <>
                      <Upload className="w-8 h-8 mb-3 text-slate-400" />
                      <p className="mb-2 text-sm text-slate-500">
                        <span className="font-semibold">Click to upload</span> or drag and drop
                      </p>
                      <p className="text-xs text-slate-500">Master Workflow File (.xlsx)</p>
                    </>
                  )}
                </div>
              </motion.div>
              <input ref={rootFileInputRef} type="file" className="hidden" accept=".xlsx" onChange={handleRootFileChange} />
            </label>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h2 className="text-lg font-medium mb-4 flex items-center gap-2">
              <FileJson className="w-5 h-5 text-indigo-600" />
              2. Workflow Data File
            </h2>
            <label className="block w-full cursor-pointer">
              <motion.div 
                initial={false}
                animate={{ 
                  backgroundColor: newFile ? '#ecfdf5' : '#f8fafc',
                  borderColor: newFile ? '#10b981' : '#cbd5e1'
                }}
                className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-xl transition-colors"
              >
                <div className="flex flex-col items-center justify-center pt-5 pb-6">
                  {newFile ? (
                    <motion.div initial={{ scale: 0 }} animate={{ scale: 1 }} className="flex flex-col items-center">
                      <CheckCircle2 className="w-8 h-8 mb-2 text-emerald-500" />
                      <p className="text-sm font-medium text-emerald-700">{newFile.name}</p>
                      {newFileRows !== null && (
                        <p className="text-xs text-emerald-600 mt-1">{newFileRows} rows detected</p>
                      )}
                    </motion.div>
                  ) : (
                    <>
                      <Upload className="w-8 h-8 mb-3 text-slate-400" />
                      <p className="mb-2 text-sm text-slate-500">
                        <span className="font-semibold">Click to upload</span> or drag and drop
                      </p>
                      <p className="text-xs text-slate-500">CSV or JSON File</p>
                    </>
                  )}
                </div>
              </motion.div>
              <input ref={newFileInputRef} type="file" className="hidden" accept=".csv,.json" onChange={handleNewFileChange} />
            </label>
          </div>
        </div>

        <div className="flex justify-center gap-4">
          <button
            onClick={handleReset}
            disabled={!rootFile && !newFile && previewRows.length === 0}
            className="flex items-center gap-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 px-8 py-3 rounded-full font-medium transition-all disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <RotateCcw className="w-5 h-5" />
            Reset
          </button>
          <button
            onClick={handleProcess}
            disabled={!rootFile || !newFile || isProcessing}
            className="flex items-center gap-2 bg-slate-900 hover:bg-slate-800 text-white px-8 py-3 rounded-full font-medium transition-all disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {isProcessing ? (
              <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
            ) : (
              <Play className="w-5 h-5" />
            )}
            Run Comparison
          </button>
        </div>

        {verificationSummary && (
          <div className="bg-emerald-50 border border-emerald-200 rounded-2xl p-6">
            <h2 className="text-lg font-medium text-emerald-900 mb-4 flex items-center gap-2">
              <CheckSquare className="w-5 h-5" />
              Post-Update Verification Summary
            </h2>
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4 mb-6">
              <div className="bg-white p-4 rounded-xl border border-emerald-100 shadow-sm">
                <p className="text-xs text-emerald-600 font-medium uppercase tracking-wider">Total Compared</p>
                <p className="text-2xl font-semibold text-emerald-900 mt-1">{verificationSummary.totalCompared}</p>
              </div>
              <div className="bg-white p-4 rounded-xl border border-emerald-100 shadow-sm">
                <p className="text-xs text-emerald-600 font-medium uppercase tracking-wider">Workflow Updates</p>
                <p className="text-2xl font-semibold text-emerald-900 mt-1">{verificationSummary.workflowUpdates}</p>
              </div>
              <div className="bg-white p-4 rounded-xl border border-emerald-100 shadow-sm">
                <p className="text-xs text-emerald-600 font-medium uppercase tracking-wider">refID Updates</p>
                <p className="text-2xl font-semibold text-emerald-900 mt-1">{verificationSummary.refIdUpdates}</p>
              </div>
              <div className="bg-white p-4 rounded-xl border border-emerald-100 shadow-sm">
                <p className="text-xs text-emerald-600 font-medium uppercase tracking-wider">Both Updates</p>
                <p className="text-2xl font-semibold text-emerald-900 mt-1">{verificationSummary.bothUpdates}</p>
              </div>
              <div className="bg-emerald-600 p-4 rounded-xl shadow-sm text-white">
                <p className="text-xs text-emerald-100 font-medium uppercase tracking-wider">Total Updated</p>
                <p className="text-2xl font-semibold mt-1">{verificationSummary.totalUpdated}</p>
              </div>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => downloadVerificationReport('csv')}
                className="flex items-center gap-2 bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-200 px-4 py-2 rounded-lg text-sm font-medium transition-colors"
              >
                <FileText className="w-4 h-4" />
                Download Post-Update Verification (CSV)
              </button>
              <button
                onClick={() => downloadVerificationReport('xlsx')}
                className="flex items-center gap-2 bg-white hover:bg-emerald-50 text-emerald-700 border border-emerald-200 px-4 py-2 rounded-lg text-sm font-medium transition-colors"
              >
                <FileSpreadsheet className="w-4 h-4" />
                Download Post-Update Verification (XLSX)
              </button>
            </div>
          </div>
        )}

        {validationResult && (
          <div className={`border rounded-2xl p-6 ${validationResult.status === 'PASSED' ? 'bg-emerald-50 border-emerald-200' : 'bg-red-50 border-red-200'}`}>
            <h2 className={`text-lg font-medium mb-4 flex items-center gap-2 ${validationResult.status === 'PASSED' ? 'text-emerald-900' : 'text-red-900'}`}>
              {validationResult.status === 'PASSED' ? (
                <><CheckCircle2 className="w-5 h-5" /> All validations passed ✅</>
              ) : (
                <><AlertCircle className="w-5 h-5" /> Validation failed ❌</>
              )}
            </h2>
            
            {validationResult.status === 'PASSED' ? (
              <p className="text-emerald-700 mb-6">Workflow Steps and referenceIDs are unique.</p>
            ) : (
              <div className="mb-6">
                <p className="text-red-700 mb-4">Duplicate Workflow Steps or referenceIDs detected.</p>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                  <div className="bg-white p-4 rounded-xl border border-red-100 shadow-sm">
                    <p className="text-xs text-red-600 font-medium uppercase tracking-wider">Duplicate Workflow Steps</p>
                    <p className="text-2xl font-semibold text-red-900 mt-1">{validationResult.duplicateWorkflowSteps}</p>
                  </div>
                  <div className="bg-white p-4 rounded-xl border border-red-100 shadow-sm">
                    <p className="text-xs text-red-600 font-medium uppercase tracking-wider">Duplicate referenceIDs</p>
                    <p className="text-2xl font-semibold text-red-900 mt-1">{validationResult.duplicateReferenceIds}</p>
                  </div>
                  <div className="bg-white p-4 rounded-xl border border-red-100 shadow-sm">
                    <p className="text-xs text-red-600 font-medium uppercase tracking-wider">Duplicate Workflow + refIDs</p>
                    <p className="text-2xl font-semibold text-red-900 mt-1">{validationResult.duplicateWorkflowAndRefIds}</p>
                  </div>
                  <div className="bg-white p-4 rounded-xl border border-red-100 shadow-sm">
                    <p className="text-xs text-red-600 font-medium uppercase tracking-wider">Total Affected Rows</p>
                    <p className="text-2xl font-semibold text-red-900 mt-1">{validationResult.detailedRecords.length}</p>
                  </div>
                </div>

                {validationResult.detailedRecords.length > 0 && (
                  <div className="mt-6 bg-white rounded-xl border border-red-200 overflow-hidden shadow-sm">
                    <div className="overflow-x-auto max-h-80">
                      <table className="w-full text-left border-collapse">
                        <thead className="bg-red-50 sticky top-0 z-10">
                          <tr>
                            <th className="p-3 text-xs font-semibold text-red-800 border-b border-red-100">Type</th>
                            <th className="p-3 text-xs font-semibold text-red-800 border-b border-red-100">Row</th>
                            <th className="p-3 text-xs font-semibold text-red-800 border-b border-red-100">NAME</th>
                            <th className="p-3 text-xs font-semibold text-red-800 border-b border-red-100">AlternateName</th>
                            <th className="p-3 text-xs font-semibold text-red-800 border-b border-red-100">referenceID</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-red-100">
                          {validationResult.detailedRecords.map((record, idx) => (
                            <tr key={idx} className="hover:bg-red-50/50">
                              <td className="p-3 text-sm text-red-900 font-medium whitespace-nowrap">{record.type}</td>
                              <td className="p-3 text-sm text-red-900 whitespace-nowrap">{record.rowNumber}</td>
                              <td className="p-3 text-sm text-red-900 min-w-[200px]">{record.name}</td>
                              <td className="p-3 text-sm text-red-900 min-w-[150px]">{record.alternateName}</td>
                              <td className="p-3 text-sm text-red-900 min-w-[150px]">{record.referenceId}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}

            <div className="flex flex-wrap gap-3">
              <button
                onClick={downloadDuplicateValidationReport}
                className={`flex items-center gap-2 bg-white px-4 py-2 rounded-lg text-sm font-medium transition-colors border ${validationResult.status === 'PASSED' ? 'hover:bg-emerald-50 text-emerald-700 border-emerald-200' : 'hover:bg-red-50 text-red-700 border-red-200'}`}
              >
                <FileSpreadsheet className="w-4 h-4" />
                Download Duplicate Validation Report (XLSX)
              </button>
              
              <button
                onClick={() => downloadUpdatedFile(false)}
                className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-colors text-white ${validationResult.status === 'PASSED' ? 'bg-emerald-600 hover:bg-emerald-700' : 'bg-amber-600 hover:bg-amber-700'}`}
              >
                <Download className="w-4 h-4" />
                {validationResult.status === 'PASSED' ? 'Download Updated XLSX' : 'Download Updated XLSX (With Duplicates)'}
              </button>

              {validationResult.status === 'FAILED' && (
                <button
                  onClick={() => downloadUpdatedFile(true)}
                  className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-colors text-white bg-indigo-600 hover:bg-indigo-700"
                >
                  <Download className="w-4 h-4" />
                  Download Updated XLSX (Without Duplicates)
                </button>
              )}
            </div>
          </div>
        )}

        {previewRows.length > 0 && (
          <div className="bg-indigo-50 border border-indigo-200 rounded-2xl p-6">
            <h2 className="text-lg font-medium text-indigo-900 mb-4 flex items-center gap-2">
              <FileText className="w-5 h-5" />
              Run Comparison Summary
            </h2>
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-4">
              <div className="bg-white p-4 rounded-xl border border-indigo-100 shadow-sm">
                <p className="text-xs text-indigo-600 font-medium uppercase tracking-wider">Total Rows</p>
                <p className="text-2xl font-semibold text-indigo-900 mt-1">{previewRows.length}</p>
              </div>
              {Object.entries(counts).map(([type, count]) => (
                <div key={type} className="bg-white p-4 rounded-xl border border-indigo-100 shadow-sm">
                  <p className="text-xs text-indigo-600 font-medium uppercase tracking-wider truncate" title={type}>{type}</p>
                  <p className="text-2xl font-semibold text-indigo-900 mt-1">{count}</p>
                </div>
              ))}
            </div>
          </div>
        )}

        {previewRows.length > 0 && (
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
            <div className="p-6 border-b border-slate-200 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div>
                <h2 className="text-lg font-medium">Preview Changes</h2>
                <p className="text-sm text-slate-500 mt-1">Select the rows you want to update in the ROOT XLSX file.</p>
              </div>
              <div className="flex gap-3">
                <button
                  onClick={downloadDifferenceReport}
                  className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-lg text-sm font-medium transition-colors"
                >
                  <Download className="w-4 h-4" />
                  Current Difference Report
                </button>
                <button
                  onClick={applyUpdatesAndValidate}
                  disabled={selectedRowIds.size === 0 || isProcessing}
                  className="flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg text-sm font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <CheckSquare className="w-4 h-4" />
                  Apply Updates & Validate
                </button>
              </div>
            </div>

            <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-col gap-4">
              <div className="flex flex-wrap gap-2">
                <button 
                  onClick={() => setActiveFilter('All')} 
                  className={`px-3 py-1.5 border rounded-md text-sm font-medium transition-colors ${activeFilter === 'All' ? 'bg-slate-800 text-white border-slate-800' : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50'}`}
                >
                  All ({previewRows.length})
                </button>
                {Object.entries(counts).map(([type, count]) => (
                  <button 
                    key={type}
                    onClick={() => setActiveFilter(type)} 
                    className={`px-3 py-1.5 border rounded-md text-sm font-medium transition-colors flex items-center gap-2 ${activeFilter === type ? 'bg-slate-800 text-white border-slate-800' : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50'}`}
                  >
                    {type}
                    <span className={`px-1.5 py-0.5 rounded-full text-xs ${activeFilter === type ? 'bg-slate-700 text-slate-200' : 'bg-slate-100 text-slate-500'}`}>
                      {count}
                    </span>
                  </button>
                ))}
              </div>
              
              <div className="flex gap-2 pt-2 border-t border-slate-200">
                <button onClick={handleSelectVisible} className="px-3 py-1.5 bg-white border border-slate-300 rounded-md text-sm font-medium hover:bg-slate-50 text-slate-700">
                  Select All Visible
                </button>
                <button onClick={handleDeselectVisible} className="px-3 py-1.5 bg-white border border-slate-300 rounded-md text-sm font-medium hover:bg-slate-50 text-slate-700">
                  Deselect All Visible
                </button>
                <div className="h-6 w-px bg-slate-300 mx-1 self-center"></div>
                <button 
                  onClick={handleUndo} 
                  disabled={historyIndex === 0}
                  className="px-3 py-1.5 bg-white border border-slate-300 rounded-md text-sm font-medium hover:bg-slate-50 text-slate-700 disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-1"
                >
                  <Undo2 className="w-4 h-4" /> Undo
                </button>
                <button 
                  onClick={handleRedo} 
                  disabled={historyIndex === selectionHistory.length - 1}
                  className="px-3 py-1.5 bg-white border border-slate-300 rounded-md text-sm font-medium hover:bg-slate-50 text-slate-700 disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-1"
                >
                  <Redo2 className="w-4 h-4" /> Redo
                </button>
                <div className="ml-auto text-sm text-slate-500 flex items-center">
                  {selectedRowIds.size} total rows selected
                </div>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-max">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200">
                    <th className="p-3 font-medium text-slate-600 text-sm w-12 text-center">Select</th>
                    <th className="p-3 font-medium text-slate-600 text-sm">Row</th>
                    <th className="p-3 font-medium text-slate-600 text-sm">Change Type</th>
                    <th className="p-3 font-medium text-slate-600 text-sm">Old NAME</th>
                    <th className="p-3 font-medium text-slate-600 text-sm">New NAME</th>
                    <th className="p-3 font-medium text-slate-600 text-sm">Old Alternate</th>
                    <th className="p-3 font-medium text-slate-600 text-sm">New Alternate</th>
                    <th className="p-3 font-medium text-slate-600 text-sm">Old referenceID</th>
                    <th className="p-3 font-medium text-slate-600 text-sm">New referenceID</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredRows.map(row => {
                    const isUpdatable = canUpdate(row.changeType);
                    const isSelected = selectedRowIds.has(row.id);
                    return (
                      <tr key={row.id} className={`hover:bg-slate-50 transition-colors ${isSelected ? 'bg-indigo-50/50' : ''}`}>
                        <td className="p-3 text-center">
                          <button
                            onClick={() => isUpdatable && toggleRowSelection(row.id)}
                            disabled={!isUpdatable}
                            className={`flex items-center justify-center w-full ${!isUpdatable ? 'opacity-30 cursor-not-allowed' : 'cursor-pointer'}`}
                          >
                            {isSelected ? <CheckSquare className="w-5 h-5 text-indigo-600" /> : <Square className="w-5 h-5 text-slate-400" />}
                          </button>
                        </td>
                        <td className="p-3 text-sm text-slate-500">{row.rootRowIndex !== -1 ? row.rootRowIndex : 'N/A'}</td>
                        <td className="p-3 text-sm">
                          <span className={`px-2 py-1 rounded-full text-xs font-medium
                            ${row.changeType === 'MATCH' ? 'bg-emerald-100 text-emerald-700' : ''}
                            ${row.changeType === 'Workflow Change' ? 'bg-blue-100 text-blue-700' : ''}
                            ${row.changeType === 'referenceID Change' ? 'bg-amber-100 text-amber-700' : ''}
                            ${row.changeType === 'Workflow + referenceID Change' ? 'bg-purple-100 text-purple-700' : ''}
                            ${row.changeType === 'Label Change' ? 'bg-slate-100 text-slate-700' : ''}
                            ${row.changeType === 'Relationship Missing' ? 'bg-red-100 text-red-700' : ''}
                            ${row.changeType === 'Invalid Format' ? 'bg-rose-100 text-rose-700' : ''}
                          `}>
                            {row.changeType}
                          </span>
                        </td>
                        <td className="p-3 text-sm text-slate-700 max-w-xs truncate" title={row.oldName}>{row.oldName}</td>
                        <td className="p-3 text-sm text-slate-700 max-w-xs truncate" title={row.newName}>{row.newName}</td>
                        <td className="p-3 text-sm text-slate-700 max-w-xs truncate" title={row.oldAlternate}>{row.oldAlternate}</td>
                        <td className="p-3 text-sm text-slate-700 max-w-xs truncate" title={row.newAlternate}>{row.newAlternate}</td>
                        <td className="p-3 text-sm text-slate-700 max-w-xs truncate font-mono text-xs" title={row.oldReferenceId}>{row.oldReferenceId}</td>
                        <td className="p-3 text-sm text-slate-700 max-w-xs truncate font-mono text-xs" title={row.newReferenceId}>{row.newReferenceId}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
