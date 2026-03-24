import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

const SALARY_BASE = 5000.0;
const SALARY_LIMIT = 24_500.0;

const WORK_HOURS_NORM_2026: Record<number, number> = {
  1: 136, 2: 151, 3: 167, 4: 175,
  5: 143, 6: 167, 7: 184, 8: 168,
  9: 176, 10: 176, 11: 159, 12: 175,
};

const MONTH_NAMES_RU: Record<number, string> = {
  1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
  5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
  9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
};

export interface EmployeeInfo {
  fio: string;
  rawFio: string;
  status: string;
  dismissDate: string | null;
  sickEnd: Date | null;
  row: number;
  existingHours: Record<number, number>;
  totalExistingHours: number;
}

export interface EmployeeResult {
  fio: string;
  status: string;
  totalHours: number;
  salary: number;
  overtime: number;
  dayHours: Record<number, number>;
}

export interface PayrollSession {
  masterBuffer: Buffer | null;
  masterFileName: string | null;
  month: number;
  employees: EmployeeInfo[];
  dayColMap: Record<number, number>;
  headerRow: number;
  uploadedReports: string[];
  reportHours: Record<string, Record<number, number>>;
  results: EmployeeResult[];
  isProcessed: boolean;
  resultBuffer: Buffer | null;
}

export function createSession(): PayrollSession {
  return {
    masterBuffer: null,
    masterFileName: null,
    month: 1,
    employees: [],
    dayColMap: {},
    headerRow: 0,
    uploadedReports: [],
    reportHours: {},
    results: [],
    isProcessed: false,
    resultBuffer: null,
  };
}

let currentSession: PayrollSession = createSession();

export function getSession(): PayrollSession {
  return currentSession;
}

export function resetSession(): void {
  const tmpDir = path.join(process.cwd(), "tmp_payroll");
  if (fs.existsSync(tmpDir)) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
  currentSession = createSession();
}

function extractCleanFio(rawFio: string): string {
  const parts = rawFio.split("-");
  return parts[0].trim().toUpperCase();
}

function parseEmployeeStatus(rawFio: string): {
  status: string;
  dismissDate: Date | null;
  sickEnd: Date | null;
} {
  const upper = rawFio.toUpperCase();
  let dismissDate: Date | null = null;
  let sickEnd: Date | null = null;

  const dismissMatch = upper.match(/УВОЛ[^\d]*(\d{1,2})[./](\d{2})[./](\d{2,4})/);
  if (dismissMatch) {
    const day = parseInt(dismissMatch[1]);
    const mon = parseInt(dismissMatch[2]);
    let yr = parseInt(dismissMatch[3]);
    if (yr < 100) yr += 2000;
    try {
      dismissDate = new Date(yr, mon - 1, day);
    } catch {}
    return { status: "УВОЛЕН", dismissDate, sickEnd };
  }

  const sickMatch = upper.match(/БОЛН[^\d]*(?:ДО\s*)?(\d{1,2})[./](\d{2})[./](\d{2,4})/);
  if (sickMatch) {
    const day = parseInt(sickMatch[1]);
    const mon = parseInt(sickMatch[2]);
    let yr = parseInt(sickMatch[3]);
    if (yr < 100) yr += 2000;
    try {
      sickEnd = new Date(yr, mon - 1, day);
    } catch {}
    return { status: "БОЛЬНИЧНЫЙ", dismissDate, sickEnd };
  }

  if (upper.includes("БОЛН")) return { status: "БОЛЬНИЧНЫЙ", dismissDate, sickEnd };
  if (upper.includes("ОТПУСК") || upper.includes("ОТП")) return { status: "ОТПУСК", dismissDate, sickEnd };

  return { status: "РАБОТАЕТ", dismissDate, sickEnd };
}

export async function parseMasterFile(buffer: Buffer, month: number): Promise<{
  employees: EmployeeInfo[];
  dayColMap: Record<number, number>;
  headerRow: number;
}> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];

  let headerRow = 0;
  const dayColMap: Record<number, number> = {};

  for (let rowIdx = 1; rowIdx <= 10; rowIdx++) {
    const row = ws.getRow(rowIdx);
    const dayVals: number[] = [];
    const dayAllCols: Record<number, number[]> = {};

    for (let col = 1; col <= (row.cellCount || 100); col++) {
      const val = row.getCell(col).value;
      const numVal = typeof val === "number" ? val : null;
      if (numVal !== null && numVal >= 1 && numVal <= 31) {
        dayVals.push(numVal);
        const day = Math.floor(numVal);
        if (!dayAllCols[day]) dayAllCols[day] = [];
        dayAllCols[day].push(col);
      }
    }

    if (dayVals.length >= 20) {
      headerRow = rowIdx;
      for (const [day, cols] of Object.entries(dayAllCols)) {
        const d = parseInt(day);
        dayColMap[d] = cols.length >= 2 ? cols[1] : cols[0];
      }
      break;
    }
  }

  if (headerRow === 0) {
    throw new Error("Не найдена строка заголовка с датами в Мастер-файле");
  }

  const employees: EmployeeInfo[] = [];
  const fioCol = 2;

  for (let rowIdx = headerRow + 2; rowIdx <= ws.rowCount; rowIdx++) {
    const row = ws.getRow(rowIdx);
    const cellFio = row.getCell(fioCol).value;

    if (!cellFio || typeof cellFio !== "string") continue;
    const fioStr = cellFio.trim();
    if (!fioStr) continue;

    const { status, dismissDate, sickEnd } = parseEmployeeStatus(fioStr);
    const fioClean = extractCleanFio(fioStr);

    const existingHours: Record<number, number> = {};
    let totalExisting = 0;
    for (const [dayStr, col] of Object.entries(dayColMap)) {
      const day = parseInt(dayStr);
      const cellVal = row.getCell(col).value;
      if (typeof cellVal === "number" && cellVal > 0) {
        existingHours[day] = cellVal;
        totalExisting += cellVal;
      }
    }

    employees.push({
      fio: fioClean,
      rawFio: fioStr,
      status,
      dismissDate: dismissDate ? dismissDate.toISOString().split("T")[0] : null,
      sickEnd,
      row: rowIdx,
      existingHours,
      totalExistingHours: Math.round(totalExisting * 10) / 10,
    });
  }

  return { employees, dayColMap, headerRow };
}

export async function uploadMaster(buffer: Buffer, month: number, fileName: string) {
  currentSession.masterBuffer = buffer;
  currentSession.masterFileName = fileName;
  currentSession.month = month;
  currentSession.isProcessed = false;
  currentSession.resultBuffer = null;
  currentSession.results = [];

  const parsed = await parseMasterFile(buffer, month);
  currentSession.employees = parsed.employees;
  currentSession.dayColMap = parsed.dayColMap;
  currentSession.headerRow = parsed.headerRow;

  const normHours = WORK_HOURS_NORM_2026[month] || 167;
  const hourCost = Math.round((SALARY_BASE / normHours) * 10000) / 10000;

  return {
    success: true,
    message: `Мастер-файл загружен: ${parsed.employees.length} сотрудников`,
    employeeCount: parsed.employees.length,
    month,
    monthName: MONTH_NAMES_RU[month] || "",
    normHours,
    hourCost,
    employees: parsed.employees.map((e) => ({
      fio: e.fio,
      rawFio: e.rawFio,
      status: e.status,
      dismissDate: e.dismissDate,
      existingHours: e.existingHours,
      totalExistingHours: e.totalExistingHours,
    })),
  };
}

function normalizeFio(fio: string): string {
  return fio.toUpperCase().trim().replace(/\s+/g, " ");
}

function fuzzyMatchFio(
  reportFio: string,
  masterKeys: Record<string, string>,
): string | null {
  const partsReport = reportFio.split(" ");
  if (!partsReport.length) return null;
  const surnameReport = partsReport[0];
  const initialReport = partsReport.length > 1 ? partsReport[1][0] : "";

  let best: string | null = null;
  let bestScore = 0;

  for (const normMaster of Object.keys(masterKeys)) {
    const partsMaster = normMaster.split(" ");
    if (!partsMaster.length) continue;
    const surnameMaster = partsMaster[0];

    let score = 0;
    if (surnameReport === surnameMaster) {
      score = 10;
    } else if (
      surnameReport.includes(surnameMaster) ||
      surnameMaster.includes(surnameReport)
    ) {
      score = 5;
    } else {
      const minLen = Math.min(surnameReport.length, surnameMaster.length, 4);
      if (
        minLen >= 3 &&
        surnameReport.substring(0, minLen) ===
          surnameMaster.substring(0, minLen)
      ) {
        score = 3;
      } else {
        continue;
      }
    }

    if (initialReport && partsMaster.length > 1) {
      if (partsMaster[1][0] === initialReport) score += 3;
    }

    if (score > bestScore) {
      bestScore = score;
      best = normMaster;
    }
  }

  return bestScore >= 5 ? best : null;
}

export async function processReport(
  buffer: Buffer,
  fileName: string,
): Promise<{
  records: number;
  matchedEmployees: string[];
  unmatchedEmployees: string[];
}> {
  const month = currentSession.month;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const ws = wb.worksheets[0];
  const headers: string[] = [];
  const headerRow = ws.getRow(1);
  for (let col = 1; col <= (headerRow.cellCount || 10); col++) {
    const v = headerRow.getCell(col).value;
    headers.push(String(v || "").trim().toUpperCase());
  }

  let dateCol = -1, cashierCol = -1, hoursCol = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].includes("ДАТА") || headers[i].includes("DATE")) dateCol = i + 1;
    if (headers[i].includes("КАССИР") || headers[i].includes("ФИО") || headers[i].includes("CASHIER")) cashierCol = i + 1;
    if (headers[i].includes("НАЧИСЛЕНО") || headers[i].includes("HOURS") || headers[i].includes("CHARGE")) hoursCol = i + 1;
  }

  if (dateCol === -1 || cashierCol === -1 || hoursCol === -1) {
    throw new Error(`Не найдены необходимые столбцы в файле ${fileName}. Нужны: Дата открытия, Кассир, НАЧИСЛЕНО`);
  }

  const fileHours: Record<string, Record<number, number>> = {};
  let records = 0;

  for (let rowIdx = 2; rowIdx <= ws.rowCount; rowIdx++) {
    const row = ws.getRow(rowIdx);
    const rawDate = row.getCell(dateCol).value;
    if (!rawDate) continue;

    let dt: Date | null = null;
    if (rawDate instanceof Date) {
      dt = rawDate;
    } else if (typeof rawDate === "string") {
      dt = new Date(rawDate);
    } else if (typeof rawDate === "number") {
      dt = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
    }

    if (!dt || isNaN(dt.getTime())) continue;
    if (dt.getMonth() + 1 !== month) continue;

    const day = dt.getDate();
    const rawFio = String(row.getCell(cashierCol).value || "").trim().toUpperCase();
    if (!rawFio || rawFio === "NAN" || rawFio.includes("СИС. АДМИНИСТРАТОР")) continue;

    const rawHours = row.getCell(hoursCol).value;
    if (rawHours === null || rawHours === undefined) continue;
    const hours = typeof rawHours === "number" ? rawHours : parseFloat(String(rawHours));
    if (isNaN(hours) || hours <= 0) continue;

    if (!fileHours[rawFio]) fileHours[rawFio] = {};
    fileHours[rawFio][day] = (fileHours[rawFio][day] || 0) + hours;
    records++;
  }

  const masterKeys: Record<string, string> = {};
  for (const emp of currentSession.employees) {
    masterKeys[normalizeFio(emp.fio)] = emp.fio;
  }

  const matched: string[] = [];
  const unmatched: string[] = [];

  for (const [reportFio, dayHours] of Object.entries(fileHours)) {
    const normReport = normalizeFio(reportFio);

    let masterFio: string | null = null;
    if (masterKeys[normReport]) {
      masterFio = masterKeys[normReport];
    } else {
      const fuzzy = fuzzyMatchFio(normReport, masterKeys);
      if (fuzzy) masterFio = masterKeys[fuzzy];
    }

    if (masterFio) {
      matched.push(masterFio);
      if (!currentSession.reportHours[masterFio]) {
        currentSession.reportHours[masterFio] = {};
      }
      for (const [dayStr, hrs] of Object.entries(dayHours)) {
        const d = parseInt(dayStr);
        currentSession.reportHours[masterFio][d] =
          (currentSession.reportHours[masterFio][d] || 0) + hrs;
      }
    } else {
      unmatched.push(reportFio);
    }
  }

  if (!currentSession.uploadedReports.includes(fileName)) {
    currentSession.uploadedReports.push(fileName);
  }
  currentSession.isProcessed = false;
  currentSession.resultBuffer = null;

  return { records, matchedEmployees: matched, unmatchedEmployees: unmatched };
}

export async function processPayroll(): Promise<EmployeeResult[]> {
  const month = currentSession.month;
  const normHours = WORK_HOURS_NORM_2026[month] || 167;
  const hourCost = SALARY_BASE / normHours;

  const lastDay = new Date(2026, month, 0).getDate();
  const results: EmployeeResult[] = [];

  for (const emp of currentSession.employees) {
    let rawHours: Record<number, number> = {
      ...(currentSession.reportHours[emp.fio] || {}),
    };

    if (emp.status === "УВОЛЕН" && emp.dismissDate) {
      const dd = new Date(emp.dismissDate);
      if (dd.getMonth() + 1 === month) {
        if (dd.getDate() < 17) {
          results.push({
            fio: emp.fio,
            status: emp.status,
            totalHours: 0,
            salary: 0,
            overtime: 0,
            dayHours: {},
          });
          continue;
        } else {
          const cutoff = dd.getDate();
          const filtered: Record<number, number> = {};
          for (const [d, h] of Object.entries(rawHours)) {
            if (parseInt(d) <= cutoff) filtered[parseInt(d)] = h;
          }
          rawHours = filtered;
        }
      }
    }

    if (emp.status === "ОТПУСК") {
      results.push({
        fio: emp.fio,
        status: emp.status,
        totalHours: 0,
        salary: 0,
        overtime: 0,
        dayHours: {},
      });
      continue;
    }

    if (emp.status === "БОЛЬНИЧНЫЙ") {
      if (emp.sickEnd) {
        if (emp.sickEnd.getMonth() + 1 === month) {
          const cutoff = emp.sickEnd.getDate();
          const filtered: Record<number, number> = {};
          for (const [d, h] of Object.entries(rawHours)) {
            if (parseInt(d) > cutoff) filtered[parseInt(d)] = h;
          }
          rawHours = filtered;
        } else {
          rawHours = {};
        }
      } else {
        rawHours = {};
      }

      if (Object.keys(rawHours).length === 0) {
        results.push({
          fio: emp.fio,
          status: emp.status,
          totalHours: 0,
          salary: 0,
          overtime: 0,
          dayHours: {},
        });
        continue;
      }
    }

    let totalHours = Object.values(rawHours).reduce((s, h) => s + h, 0);
    let overtime = 0;
    let salary = 0;

    if (totalHours <= normHours) {
      salary = totalHours * hourCost;
    } else {
      overtime = totalHours - normHours;
      salary = normHours * hourCost + overtime * hourCost * 2;
    }

    if (salary > SALARY_LIMIT) {
      if (totalHours <= normHours) {
        const scale = SALARY_LIMIT / salary;
        const newHours: Record<number, number> = {};
        for (const [d, h] of Object.entries(rawHours)) {
          newHours[parseInt(d)] = Math.round(h * scale * 10) / 10;
        }
        rawHours = newHours;
      } else {
        const maxOvertime = (SALARY_LIMIT - SALARY_BASE) / (2 * hourCost);
        const maxTotal = normHours + maxOvertime;
        const scale = maxTotal / totalHours;
        const newHours: Record<number, number> = {};
        for (const [d, h] of Object.entries(rawHours)) {
          newHours[parseInt(d)] = Math.round(h * scale * 10) / 10;
        }
        rawHours = newHours;
      }
      totalHours = Object.values(rawHours).reduce((s, h) => s + h, 0);
      overtime = Math.max(0, totalHours - normHours);
      salary = SALARY_LIMIT;
    }

    results.push({
      fio: emp.fio,
      status: emp.status,
      totalHours: Math.round(totalHours * 10) / 10,
      salary: Math.round(salary * 100) / 100,
      overtime: Math.round(overtime * 10) / 10,
      dayHours: rawHours,
    });
  }

  currentSession.results = results;
  currentSession.isProcessed = true;
  currentSession.resultBuffer = null;

  return results;
}

export async function generateResultFile(): Promise<Buffer> {
  if (currentSession.resultBuffer) {
    return currentSession.resultBuffer;
  }

  if (!currentSession.masterBuffer) {
    throw new Error("Мастер-файл не загружен");
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(currentSession.masterBuffer);
  const ws = wb.worksheets[0];

  for (const result of currentSession.results) {
    const emp = currentSession.employees.find((e) => e.fio === result.fio);
    if (!emp) continue;

    for (const [dayStr, hours] of Object.entries(result.dayHours)) {
      const day = parseInt(dayStr);
      const col = currentSession.dayColMap[day];
      if (!col) continue;

      const cell = ws.getRow(emp.row).getCell(col);
      cell.value = Math.round(hours * 10) / 10;

      if (result.status === "УВОЛЕН") {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" },
        };
      } else if (result.status === "ОТПУСК" || result.status === "БОЛЬНИЧНЫЙ") {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFFF00" },
        };
      }
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  currentSession.resultBuffer = Buffer.from(buf);
  return currentSession.resultBuffer;
}

export function getMonthInfo(month: number) {
  return {
    name: MONTH_NAMES_RU[month] || "",
    normHours: WORK_HOURS_NORM_2026[month] || 167,
    hourCost:
      Math.round((SALARY_BASE / (WORK_HOURS_NORM_2026[month] || 167)) * 10000) /
      10000,
    salaryBase: SALARY_BASE,
    salaryLimit: SALARY_LIMIT,
  };
}

export { MONTH_NAMES_RU, WORK_HOURS_NORM_2026, SALARY_BASE, SALARY_LIMIT };
