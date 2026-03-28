import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

const SALARY_BASE = 5000.0;
const SALARY_LIMIT = 24_500.0;

function getCellValue(cell: ExcelJS.Cell): any {
  const v = cell.value;
  if (v && typeof v === "object" && !Array.isArray(v) && !(v instanceof Date)) {
    if ("result" in v) return (v as any).result;
    if ("sharedFormula" in v) return (v as any).result;
    if ("formula" in v) return (v as any).result ?? null;
  }
  return v;
}

const WORK_HOURS_NORM_2026: Record<number, number> = {
  1: 120, 2: 152, 3: 168, 4: 175,
  5: 151, 6: 167, 7: 184, 8: 168,
  9: 176, 10: 176, 11: 159, 12: 176,
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
  nightPay: number;
  basePay: number;
  overtimePay: number;
  dayHours: Record<number, number>;
  existingHours: Record<number, number>;
  newHours: Record<number, number>;
  employeeSalary: number;
  normHours: number;
  hourCost: number;
  uncappedSalary: number;
}

interface DayColumns {
  col1: number;
  col2: number;
}

export interface PayrollSession {
  masterBuffer: Buffer | null;
  masterFileName: string | null;
  month: number;
  clearHours: boolean;
  employees: EmployeeInfo[];
  dayColPairs: Record<number, DayColumns>;
  headerRow: number;
  uploadedReports: string[];
  reportHours: Record<string, Record<number, number>>;
  storeHours: Record<number, number>;
  results: EmployeeResult[];
  isProcessed: boolean;
  resultBuffer: Buffer | null;
}

export function createSession(): PayrollSession {
  return {
    masterBuffer: null,
    masterFileName: null,
    month: 1,
    clearHours: true,
    employees: [],
    dayColPairs: {},
    headerRow: 0,
    uploadedReports: [],
    reportHours: {},
    storeHours: {},
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

  if (upper.includes("ЗА СВОЙ СЧЕТ") || upper.includes("ЗА СВОЙ СЧЁТ")) {
    return { status: "ЗА СВОЙ СЧЁТ", dismissDate, sickEnd };
  }

  const dismissMatch = upper.match(/УВОЛ[^\d]*(\d{1,2})[./](\d{1,2})[./](\d{2,4})/);
  if (dismissMatch) {
    const day = parseInt(dismissMatch[1]);
    const mon = parseInt(dismissMatch[2]);
    let yr = parseInt(dismissMatch[3]);
    if (yr < 100) yr += 2000;
    try { dismissDate = new Date(yr, mon - 1, day); } catch {}
    return { status: "УВОЛЕН", dismissDate, sickEnd };
  }

  const sickMatch = upper.match(/БОЛ[НЬ]?[^\d]*(?:ДО\s*)?(\d{1,2})[./](\d{1,2})[./](\d{2,4})/);
  if (sickMatch) {
    const day = parseInt(sickMatch[1]);
    const mon = parseInt(sickMatch[2]);
    let yr = parseInt(sickMatch[3]);
    if (yr < 100) yr += 2000;
    try { sickEnd = new Date(yr, mon - 1, day); } catch {}
    return { status: "БОЛЬНИЧНЫЙ", dismissDate, sickEnd };
  }

  if (upper.includes("БОЛН") || upper.includes("-БОЛ")) return { status: "БОЛЬНИЧНЫЙ", dismissDate, sickEnd };
  if (upper.includes("ОТПУСК") || upper.includes("ОТП")) return { status: "ОТПУСК", dismissDate, sickEnd };

  return { status: "РАБОТАЕТ", dismissDate, sickEnd };
}

export async function parseMasterFile(buffer: Buffer, month: number, clearHours: boolean): Promise<{
  employees: EmployeeInfo[];
  dayColPairs: Record<number, DayColumns>;
  headerRow: number;
}> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];

  let headerRow = 0;
  const dayColPairs: Record<number, DayColumns> = {};

  for (let rowIdx = 1; rowIdx <= 10; rowIdx++) {
    const row = ws.getRow(rowIdx);
    const dayAllCols: Record<number, number[]> = {};
    let dayCount = 0;

    for (let col = 1; col <= (row.cellCount || 100); col++) {
      const val = row.getCell(col).value;
      const numVal = typeof val === "number" ? val : null;
      if (numVal !== null && numVal >= 1 && numVal <= 31) {
        const day = Math.floor(numVal);
        if (!dayAllCols[day]) dayAllCols[day] = [];
        dayAllCols[day].push(col);
        dayCount++;
      }
    }

    if (dayCount >= 20) {
      headerRow = rowIdx;
      for (const [day, cols] of Object.entries(dayAllCols)) {
        const d = parseInt(day);
        if (cols.length >= 2) {
          dayColPairs[d] = { col1: cols[0], col2: cols[1] };
        } else {
          dayColPairs[d] = { col1: cols[0], col2: cols[0] };
        }
      }
      break;
    }
  }

  if (headerRow === 0) {
    throw new Error("Не найдена строка заголовка с датами в Мастер-файле");
  }

  const employees: EmployeeInfo[] = [];
  const fioCol = 2;
  let consecutiveEmpty = 0;

  for (let rowIdx = headerRow + 2; rowIdx <= ws.rowCount; rowIdx++) {
    const row = ws.getRow(rowIdx);
    const cellFio = row.getCell(fioCol).value;

    if (!cellFio || typeof cellFio !== "string" || !cellFio.trim()) {
      consecutiveEmpty++;
      if (consecutiveEmpty >= 3) break;
      continue;
    }
    consecutiveEmpty = 0;

    const fioStr = cellFio.trim();
    const lower = fioStr.toLowerCase();
    if (lower.includes("бухгалтер") || lower.includes("руководитель") || lower.includes("главн")) continue;

    const { status, dismissDate, sickEnd } = parseEmployeeStatus(fioStr);
    const fioClean = extractCleanFio(fioStr);
    if (!fioClean || fioClean.length < 2) continue;

    let existingHours: Record<number, number> = {};
    let totalExisting = 0;

    if (!clearHours) {
      for (const [dayStr, pair] of Object.entries(dayColPairs)) {
        const day = parseInt(dayStr);
        let dayTotal = 0;
        const v1 = getCellValue(row.getCell(pair.col1));
        if (typeof v1 === "number" && v1 > 0) dayTotal += v1;
        if (pair.col2 !== pair.col1) {
          const v2 = getCellValue(row.getCell(pair.col2));
          if (typeof v2 === "number" && v2 > 0) dayTotal += v2;
        }
        if (dayTotal > 0) {
          existingHours[day] = dayTotal;
          totalExisting += dayTotal;
        }
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

  return { employees, dayColPairs, headerRow };
}

export async function uploadMaster(buffer: Buffer, month: number, fileName: string, clearHours: boolean = true) {
  currentSession.masterBuffer = buffer;
  currentSession.masterFileName = fileName;
  currentSession.month = month;
  currentSession.clearHours = clearHours;
  currentSession.isProcessed = false;
  currentSession.resultBuffer = null;
  currentSession.results = [];
  currentSession.uploadedReports = [];
  currentSession.reportHours = {};

  const parsed = await parseMasterFile(buffer, month, clearHours);
  currentSession.employees = parsed.employees;
  currentSession.dayColPairs = parsed.dayColPairs;
  currentSession.headerRow = parsed.headerRow;

  const normHours = WORK_HOURS_NORM_2026[month] || 167;

  return {
    success: true,
    message: `Мастер-файл загружен: ${parsed.employees.length} сотрудников` + (clearHours ? ' (часы очищены)' : ''),
    employeeCount: parsed.employees.length,
    month,
    monthName: MONTH_NAMES_RU[month] || "",
    normHours,
    clearHours,
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
  return fio.toUpperCase().trim().replace(/\s+/g, " ").replace(/\./g, "");
}

function extractSurname(fio: string): string {
  return fio.split(/[\s.]+/)[0] || "";
}

function extractInitials(fio: string): string[] {
  const parts = fio.split(/[\s]+/);
  const initials: string[] = [];
  for (let i = 1; i < parts.length; i++) {
    const p = parts[i].replace(/\./g, "");
    if (p.length > 0) {
      if (p.length <= 2) {
        for (const ch of p) initials.push(ch);
      } else {
        initials.push(p[0]);
      }
    }
  }
  return initials;
}

function fuzzyMatchFio(
  reportFio: string,
  masterKeys: Record<string, string>,
): string | null {
  const normReport = normalizeFio(reportFio);
  const surnameReport = extractSurname(normReport);
  const initialsReport = extractInitials(normReport);

  let best: string | null = null;
  let bestScore = 0;

  for (const normMaster of Object.keys(masterKeys)) {
    const surnameMaster = extractSurname(normMaster);

    let score = 0;
    if (surnameReport === surnameMaster) {
      score = 10;
    } else if (
      surnameReport.length >= 3 && surnameMaster.length >= 3 &&
      (surnameReport.includes(surnameMaster) || surnameMaster.includes(surnameReport))
    ) {
      score = 7;
    } else {
      const minLen = Math.min(surnameReport.length, surnameMaster.length, 4);
      if (minLen >= 3 && surnameReport.substring(0, minLen) === surnameMaster.substring(0, minLen)) {
        score = 4;
      } else {
        continue;
      }
    }

    const initialsMaster = extractInitials(normMaster);
    if (initialsReport.length > 0 && initialsMaster.length > 0) {
      if (initialsReport[0] === initialsMaster[0]) score += 3;
      if (initialsReport.length > 1 && initialsMaster.length > 1 && initialsReport[1] === initialsMaster[1]) score += 2;
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
  for (let col = 1; col <= (headerRow.cellCount || 20); col++) {
    const v = headerRow.getCell(col).value;
    headers.push(String(v || "").trim().toUpperCase());
  }

  let dateCol = -1, cashierCol = -1, hoursCol = -1, endTimeCol = -1;
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (dateCol === -1 && (h.includes("ДАТА") || h.includes("DATE"))) dateCol = i + 1;
    if (cashierCol === -1 && (h.includes("КАССИР") || h.includes("ФИО") || h.includes("CASHIER") || h.includes("СОТРУДНИК"))) cashierCol = i + 1;
    if (h.includes("НАЧИСЛЕНО") || h.includes("ОКРУГЛЕНИЕ") || h === "HOURS" || h === "CHARGE") hoursCol = i + 1;
    if (h.includes("ПОСЛЕДН") || h.includes("ЗАКРЫТ") || h.includes("END") || h.includes("ВРЕМЯ ПОСЛЕДН")) endTimeCol = i + 1;
  }

  if (dateCol === -1 || cashierCol === -1 || hoursCol === -1) {
    throw new Error(`Не найдены необходимые столбцы в файле ${fileName}. Ожидаются: Дата открытия, Кассир, НАЧИСЛЕНО. Найдены: ${headers.filter(h => h).join(', ')}`);
  }

  function calculateHoursFromTimestamps(openDate: Date, closeDate: Date): number {
    const diffMs = closeDate.getTime() - openDate.getTime();
    if (diffMs <= 0) return 0;
    const totalMinutes = diffMs / 60000;
    const fullHours = Math.floor(totalMinutes / 60);
    const remainMinutes = totalMinutes % 60;
    const rounded = remainMinutes >= 15 ? fullHours + 1 : fullHours;
    const lunch = rounded > 12 ? 2 : 1;
    return Math.max(0, rounded - lunch);
  }

  const fileHours: Record<string, Record<number, number>> = {};
  const storeHours: Record<number, number> = {};
  let records = 0;

  for (let rowIdx = 2; rowIdx <= ws.rowCount; rowIdx++) {
    const row = ws.getRow(rowIdx);
    const rawDate = getCellValue(row.getCell(dateCol));
    if (!rawDate) continue;

    let dt: Date | null = null;
    if (rawDate instanceof Date) {
      dt = rawDate;
    } else if (typeof rawDate === "string") {
      const parts = rawDate.match(/(\d{1,2})[./](\d{1,2})[./](\d{2,4})/);
      if (parts) {
        let yr = parseInt(parts[3]);
        if (yr < 100) yr += 2000;
        dt = new Date(yr, parseInt(parts[2]) - 1, parseInt(parts[1]));
      } else {
        dt = new Date(rawDate);
      }
    } else if (typeof rawDate === "number") {
      dt = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
    }

    if (!dt || isNaN(dt.getTime())) continue;

    const rawCashier = getCellValue(row.getCell(cashierCol));
    const rawFio = String(rawCashier || "").trim().toUpperCase();
    if (!rawFio || rawFio === "NAN" || rawFio.includes("ИТОГО")) continue;

    let hours: number | null = null;

    const rawHours = getCellValue(row.getCell(hoursCol));
    if (typeof rawHours === "number") {
      hours = rawHours;
    } else if (rawHours !== null && rawHours !== undefined) {
      const strH = String(rawHours).replace(",", ".");
      const timeMatch = strH.match(/^(\d+):(\d+)$/);
      if (timeMatch) {
        hours = parseInt(timeMatch[1]) + parseInt(timeMatch[2]) / 60;
      } else {
        const parsed = parseFloat(strH);
        if (!isNaN(parsed)) hours = parsed;
      }
    }

    if ((hours === null || isNaN(hours as number)) && endTimeCol !== -1) {
      const rawEnd = getCellValue(row.getCell(endTimeCol));
      const openRaw = row.getCell(dateCol).value;
      const closeRaw = row.getCell(endTimeCol).value;
      let openDt: Date | null = null;
      let closeDt: Date | null = null;

      if (openRaw instanceof Date) openDt = openRaw;
      else {
        const v = getCellValue(row.getCell(dateCol));
        if (v instanceof Date) openDt = v;
      }

      if (closeRaw instanceof Date) closeDt = closeRaw;
      else {
        if (rawEnd instanceof Date) closeDt = rawEnd;
      }

      if (openDt && closeDt) {
        hours = calculateHoursFromTimestamps(openDt, closeDt);
      }
    }

    if (hours === null || isNaN(hours) || hours <= 0) continue;

    const openDtForMonth = dt;
    let reportDay: number;
    if (endTimeCol !== -1) {
      const rawEnd = row.getCell(endTimeCol).value;
      let closeDt: Date | null = null;
      if (rawEnd instanceof Date) closeDt = rawEnd;
      if (closeDt && closeDt.getDate() !== openDtForMonth.getDate()) {
        reportDay = closeDt.getDate();
        if (closeDt.getMonth() + 1 !== month) continue;
      } else {
        if (openDtForMonth.getMonth() + 1 !== month) continue;
        reportDay = openDtForMonth.getDate();
      }
    } else {
      if (openDtForMonth.getMonth() + 1 !== month) continue;
      reportDay = openDtForMonth.getDate();
    }

    if (rawFio.includes("СИС") || rawFio.includes("АДМИНИСТРАТОР") || rawFio.includes("SYS")) {
      storeHours[reportDay] = (storeHours[reportDay] || 0) + hours;
      records++;
      continue;
    }

    if (!fileHours[rawFio]) fileHours[rawFio] = {};
    fileHours[rawFio][reportDay] = Math.max(fileHours[rawFio][reportDay] || 0, hours);
    records++;
  }

  if (Object.keys(storeHours).length > 0) {
    if (!currentSession.storeHours) currentSession.storeHours = {};
    for (const [dayStr, hrs] of Object.entries(storeHours)) {
      const d = parseInt(dayStr);
      currentSession.storeHours[d] = (currentSession.storeHours[d] || 0) + hrs;
    }
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
      const fuzzyResult = fuzzyMatchFio(normReport, masterKeys);
      if (fuzzyResult) masterFio = masterKeys[fuzzyResult];
    }

    if (masterFio) {
      matched.push(masterFio);
      if (!currentSession.reportHours[masterFio]) {
        currentSession.reportHours[masterFio] = {};
      }
      for (const [dayStr, hrs] of Object.entries(dayHours)) {
        const d = parseInt(dayStr);
        currentSession.reportHours[masterFio][d] =
          Math.max(currentSession.reportHours[masterFio][d] || 0, hrs);
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

const SHIFT_HOURS = 22;
const TARGET_HOURS_MIN = 350;
const TARGET_HOURS_MAX = 375;

function getDaysInMonth(month: number, year: number = 2026): number {
  return new Date(year, month, 0).getDate();
}

function generateSchedule(
  empIndex: number,
  month: number,
  existingDays: Set<number>,
  existingTotalHours: number,
  maxDay: number = 31,
  minDay: number = 1,
): Record<number, number> {
  const daysInMonth = Math.min(getDaysInMonth(month), maxDay);
  const schedule: Record<number, number> = {};

  const maxFullShifts = Math.floor((TARGET_HOURS_MAX - existingTotalHours) / SHIFT_HOURS);
  if (maxFullShifts <= 0) return schedule;

  const startDay = (empIndex % 2 === 0) ? 1 : 2;

  const candidates: number[] = [];
  for (let d = startDay; d <= daysInMonth; d += 2) {
    if (d >= minDay && !existingDays.has(d)) candidates.push(d);
  }
  if (candidates.length < maxFullShifts) {
    for (let d = (startDay === 1 ? 2 : 1); d <= daysInMonth; d += 2) {
      if (d >= minDay && !existingDays.has(d)) candidates.push(d);
    }
  }

  const shiftsToAdd = Math.min(maxFullShifts, candidates.length);
  for (let i = 0; i < shiftsToAdd; i++) {
    schedule[candidates[i]] = SHIFT_HOURS;
  }

  return schedule;
}

export async function processPayroll(): Promise<EmployeeResult[]> {
  const month = currentSession.month;
  const normHours = WORK_HOURS_NORM_2026[month] || 167;
  const hourCost = SALARY_BASE / normHours;
  const results: EmployeeResult[] = [];

  let empIndex = 0;

  for (const emp of currentSession.employees) {
    const existingHours = { ...emp.existingHours };
    const reportHoursForEmp = currentSession.reportHours[emp.fio] || {};
    const newHours: Record<number, number> = {};

    for (const [dayStr, hrs] of Object.entries(reportHoursForEmp)) {
      const d = parseInt(dayStr);
      if (!existingHours[d]) {
        newHours[d] = hrs;
      } else if (hrs > existingHours[d]) {
        newHours[d] = hrs;
      }
    }

    let mergedHours: Record<number, number> = { ...existingHours };
    for (const [dayStr, hrs] of Object.entries(newHours)) {
      mergedHours[parseInt(dayStr)] = hrs;
    }

    const zeroResult = (status: string): EmployeeResult => ({
      fio: emp.fio, status,
      totalHours: 0, salary: 0, overtime: 0,
      nightPay: 0, basePay: 0, overtimePay: 0, uncappedSalary: 0,
      dayHours: {}, existingHours, newHours: {},
      employeeSalary: SALARY_BASE, normHours, hourCost: Math.round(hourCost * 100) / 100,
    });

    if (emp.status === "ЗА СВОЙ СЧЁТ") {
      results.push(zeroResult(emp.status));
      continue;
    }

    if (emp.status === "ОТПУСК") {
      results.push(zeroResult(emp.status));
      continue;
    }

    let maxDay = getDaysInMonth(month);

    if (emp.status === "УВОЛЕН" && emp.dismissDate) {
      const dd = new Date(emp.dismissDate);
      if (dd.getMonth() + 1 === month) {
        if (dd.getDate() < 17) {
          results.push({ ...zeroResult(emp.status), newHours });
          continue;
        } else {
          maxDay = dd.getDate();
          const filtered: Record<number, number> = {};
          for (const [d, h] of Object.entries(mergedHours)) {
            if (parseInt(d) <= maxDay) filtered[parseInt(d)] = h;
          }
          mergedHours = filtered;
        }
      }
    }

    let minDay = 1;

    if (emp.status === "БОЛЬНИЧНЫЙ") {
      if (emp.sickEnd) {
        if (emp.sickEnd.getMonth() + 1 === month) {
          const cutoff = emp.sickEnd.getDate();
          minDay = cutoff + 1;
          const filtered: Record<number, number> = {};
          for (const [d, h] of Object.entries(mergedHours)) {
            if (parseInt(d) > cutoff) filtered[parseInt(d)] = h;
          }
          mergedHours = filtered;
          if (minDay > getDaysInMonth(month)) {
            results.push(zeroResult(emp.status));
            continue;
          }
        } else {
          const payrollStart = new Date(2026, month - 1, 1);
          const payrollEnd = new Date(2026, month, 0);
          if (emp.sickEnd > payrollEnd) {
            results.push(zeroResult(emp.status));
            continue;
          } else if (emp.sickEnd < payrollStart) {
            // sick leave already ended before this month — treat as working
          }
        }
      } else {
        results.push(zeroResult(emp.status));
        continue;
      }
    }

    const existingDays = new Set(Object.keys(mergedHours).map(Number));
    const existingTotalHours = Object.values(mergedHours).reduce((s, h) => s + h, 0);
    const generated = generateSchedule(empIndex, month, existingDays, existingTotalHours, maxDay, minDay);
    empIndex++;

    for (const [dayStr, hrs] of Object.entries(generated)) {
      const d = parseInt(dayStr);
      if (!mergedHours[d]) {
        mergedHours[d] = hrs;
        newHours[d] = hrs;
      }
    }

    const totalHours = Object.values(mergedHours).reduce((s, h) => s + h, 0);
    const overtime = Math.max(0, totalHours - normHours);

    const basePay = Math.min(totalHours, normHours) * hourCost;
    const overtimePay = overtime * hourCost * 2;
    const nightPay = totalHours * hourCost;
    const uncappedSalary = basePay + overtimePay + nightPay;

    const salary = Math.min(uncappedSalary, SALARY_LIMIT);

    results.push({
      fio: emp.fio,
      status: emp.status,
      totalHours: Math.round(totalHours * 10) / 10,
      salary: Math.round(salary * 100) / 100,
      overtime: Math.round(overtime * 10) / 10,
      basePay: Math.round(basePay * 100) / 100,
      overtimePay: Math.round(overtimePay * 100) / 100,
      nightPay: Math.round(nightPay * 100) / 100,
      uncappedSalary: Math.round(uncappedSalary * 100) / 100,
      dayHours: mergedHours,
      existingHours,
      newHours,
      employeeSalary: SALARY_BASE,
      normHours,
      hourCost: Math.round(hourCost * 100) / 100,
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

  const greenFill = {
    type: "pattern" as const,
    pattern: "solid" as const,
    fgColor: { argb: "FF90EE90" },
  };
  const redFill = {
    type: "pattern" as const,
    pattern: "solid" as const,
    fgColor: { argb: "FFFF6B6B" },
  };
  const yellowFill = {
    type: "pattern" as const,
    pattern: "solid" as const,
    fgColor: { argb: "FFFFFF00" },
  };

  for (const result of currentSession.results) {
    const emp = currentSession.employees.find((e) => e.fio === result.fio);
    if (!emp) continue;

    const row = ws.getRow(emp.row);

    if (currentSession.clearHours) {
      for (const [, pair] of Object.entries(currentSession.dayColPairs)) {
        row.getCell(pair.col1).value = null;
        if (pair.col2 !== pair.col1) {
          row.getCell(pair.col2).value = null;
        }
      }
    }

    const applyFill = (cell: ExcelJS.Cell, fill: ExcelJS.Fill) => {
      cell.style = { ...cell.style, fill };
    };

    if (result.status === "ОТПУСК" || result.status === "ЗА СВОЙ СЧЁТ" || (result.status === "БОЛЬНИЧНЫЙ" && result.totalHours === 0)) {
      for (const [, pair] of Object.entries(currentSession.dayColPairs)) {
        applyFill(row.getCell(pair.col1), yellowFill);
        if (pair.col2 !== pair.col1) applyFill(row.getCell(pair.col2), yellowFill);
      }
      continue;
    }

    if (result.status === "БОЛЬНИЧНЫЙ" && result.totalHours > 0) {
      for (const [dayStr, pair] of Object.entries(currentSession.dayColPairs)) {
        const dayNum = parseInt(dayStr);
        if (!result.dayHours[dayNum]) {
          applyFill(row.getCell(pair.col1), yellowFill);
          if (pair.col2 !== pair.col1) applyFill(row.getCell(pair.col2), yellowFill);
        }
      }
    }

    if (result.status === "УВОЛЕН") {
      for (const [, pair] of Object.entries(currentSession.dayColPairs)) {
        applyFill(row.getCell(pair.col1), redFill);
        if (pair.col2 !== pair.col1) applyFill(row.getCell(pair.col2), redFill);
      }
    }

    for (const [dayStr, hours] of Object.entries(result.dayHours)) {
      const day = parseInt(dayStr);
      const pair = currentSession.dayColPairs[day];
      if (!pair) continue;

      const cell = row.getCell(pair.col2);

      if (currentSession.clearHours) {
        cell.value = Math.round(hours * 10) / 10;
      } else {
        const col1Val = row.getCell(pair.col1).value;
        const col1Num = (typeof col1Val === "number" && col1Val > 0) ? col1Val : 0;
        const existingCol2 = row.getCell(pair.col2).value;
        const col2Num = (typeof existingCol2 === "number" && existingCol2 > 0) ? existingCol2 : 0;
        const delta = Math.round((hours - col1Num - col2Num) * 10) / 10;
        if (delta <= 0) continue;
        cell.value = delta + col2Num;
      }

      if (result.status !== "УВОЛЕН") {
        applyFill(cell, greenFill);
      }
    }
  }

  const resultByRow: Record<number, number> = {};
  for (const result of currentSession.results) {
    const emp = currentSession.employees.find((e) => e.fio === result.fio);
    if (emp) resultByRow[emp.row] = result.totalHours;
  }

  let decEmployeeTotal = 0;
  let janEmployeeTotal = 0;
  for (const emp of currentSession.employees) {
    const boCell = ws.getRow(emp.row).getCell(67).value;
    let decBO = 0;
    if (typeof boCell === "number") {
      decBO = boCell;
    } else if (boCell && typeof boCell === "object" && "result" in boCell) {
      const r = (boCell as { result?: unknown }).result;
      if (typeof r === "number") decBO = r;
    }
    const janBO = resultByRow[emp.row] || 0;
    if (decBO > 0) {
      decEmployeeTotal += decBO;
      janEmployeeTotal += janBO;
    }
  }

  for (const result of currentSession.results) {
    const emp = currentSession.employees.find((e) => e.fio === result.fio);
    if (!emp) continue;
    const empRow = ws.getRow(emp.row);
    if (currentSession.clearHours) {
      empRow.getCell(65).value = 0;
      empRow.getCell(66).value = result.totalHours;
    } else {
      let bmTotal = 0;
      for (const [, pair] of Object.entries(currentSession.dayColPairs)) {
        const v = empRow.getCell(pair.col1).value;
        if (typeof v === "number") bmTotal += v;
      }
      empRow.getCell(65).value = Math.round(bmTotal * 10) / 10;
      empRow.getCell(66).value = Math.round((result.totalHours - bmTotal) * 10) / 10;
    }
    empRow.getCell(67).value = result.totalHours;
  }

  const firstEmpRow = Math.min(...currentSession.employees.map((e) => e.row));
  const lastEmpRow = Math.max(...currentSession.employees.map((e) => e.row));

  for (let r = 1; r <= lastEmpRow + 20; r++) {
    const row = ws.getRow(r);
    for (const col of [65, 66, 67]) {
      const cell = row.getCell(col);
      const v = cell.value;
      if (v && typeof v === "object" && ("formula" in v || "sharedFormula" in v)) {
        const cached = (v as { result?: unknown }).result;
        cell.value = typeof cached === "number" ? cached : 0;
      }
    }
  }

  function colLetterToNum(letters: string): number {
    return letters.split("").reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0);
  }

  function evalFormulaCell(formula: string): number {
    const bmbnRefs = formula.match(/B[MNO]\d+/g);
    if (bmbnRefs) {
      let total = 0;
      for (const ref of bmbnRefs) {
        const refRow = parseInt(ref.replace(/^B[MNO]/, ""));
        const boVal = ws.getRow(refRow).getCell(67).value;
        if (typeof boVal === "number") total += boVal;
      }
      const otherParts = formula.split("+").filter(
        (p) => !p.trim().match(/^B[MNO]\d+$/) && !p.trim().startsWith("#")
      );
      for (const part of otherParts) {
        const m = part.trim().match(/^([A-Z]+)(\d+)$/);
        if (m) {
          const v = ws.getRow(parseInt(m[2])).getCell(colLetterToNum(m[1])).value;
          if (typeof v === "number") total += v;
        }
      }
      return total;
    }
    return 0;
  }

  function evalSumFormula(formula: string): number {
    const m = formula.match(/SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)/);
    if (!m) return 0;
    const col = colLetterToNum(m[1]);
    const startRow = parseInt(m[2]);
    const endRow = parseInt(m[4]);
    let total = 0;
    for (let r = startRow; r <= endRow; r++) {
      const v = ws.getRow(r).getCell(col).value;
      if (typeof v === "number") total += v;
    }
    return total;
  }

  function evalCompositeFormula(formula: string): number {
    const parts = formula.split("+");
    let total = 0;
    for (const part of parts) {
      const m = part.trim().match(/^([A-Z]+)(\d+)$/);
      if (m) {
        const v = ws.getRow(parseInt(m[2])).getCell(colLetterToNum(m[1])).value;
        if (typeof v === "number") total += v;
      }
    }
    return total;
  }

  for (let r = lastEmpRow + 1; r <= lastEmpRow + 20; r++) {
    const row = ws.getRow(r);
    for (const col of [6, 15]) {
      const cellVal = row.getCell(col).value;
      if (!cellVal || typeof cellVal !== "object" || !("formula" in cellVal)) continue;
      const formula = (cellVal as { formula: string }).formula;
      if (formula.match(/B[MN]\d+/)) {
        row.getCell(col).value = Math.round(evalFormulaCell(formula));
      }
    }
  }

  const scaleFactor = decEmployeeTotal > 0 ? janEmployeeTotal / decEmployeeTotal : 1;

  for (let r = lastEmpRow + 1; r <= lastEmpRow + 20; r++) {
    const row = ws.getRow(r);
    const storeName = row.getCell(17).value;
    if (!storeName || typeof storeName !== "string") continue;
    for (const [primary, mirror] of [[22, 23], [26, 27]] as const) {
      const val = row.getCell(primary).value;
      if (typeof val === "number" && val > 0) {
        const scaled = Math.round(val * scaleFactor);
        row.getCell(primary).value = scaled;
        row.getCell(mirror).value = scaled;
      }
    }
  }

  for (let r = lastEmpRow + 1; r <= lastEmpRow + 20; r++) {
    const row = ws.getRow(r);
    for (const col of [6, 15]) {
      const cellVal = row.getCell(col).value;
      if (!cellVal || typeof cellVal !== "object" || !("formula" in cellVal)) continue;
      const formula = (cellVal as { formula: string }).formula;
      if (formula.startsWith("SUM(")) {
        row.getCell(col).value = Math.round(evalSumFormula(formula));
      } else if (!formula.match(/B[MN]\d+/)) {
        row.getCell(col).value = Math.round(evalCompositeFormula(formula));
      }
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  currentSession.resultBuffer = Buffer.from(buf);
  return currentSession.resultBuffer;
}

export function getMonthInfo(month: number) {
  const normHours = WORK_HOURS_NORM_2026[month] || 167;
  return {
    name: MONTH_NAMES_RU[month] || "",
    normHours,
    salaryBase: SALARY_BASE,
    salaryLimit: SALARY_LIMIT,
  };
}

export { MONTH_NAMES_RU, WORK_HOURS_NORM_2026, SALARY_BASE, SALARY_LIMIT };
