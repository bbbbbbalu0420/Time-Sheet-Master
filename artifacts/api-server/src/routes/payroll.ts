import { Router, type IRouter } from "express";
import multer from "multer";
import {
  uploadMaster,
  processReport,
  processPayroll,
  generateResultFile,
  getSession,
  resetSession,
  getMonthInfo,
} from "../lib/payroll-engine";

const router: IRouter = Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

router.post("/payroll/upload-master", upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    if (!file) {
      res.status(400).json({ success: false, message: "Файл не загружен" });
      return;
    }

    const month = parseInt(String(req.body.month || "1"));
    if (month < 1 || month > 12) {
      res.status(400).json({ success: false, message: "Некорректный месяц" });
      return;
    }

    const result = await uploadMaster(file.buffer, month, file.originalname);
    res.json(result);
  } catch (err: any) {
    req.log.error({ err }, "Error uploading master file");
    res.status(500).json({ success: false, message: err.message || "Ошибка загрузки мастер-файла" });
  }
});

router.post("/payroll/upload-reports", upload.array("files", 50), async (req, res) => {
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      res.status(400).json({ success: false, message: "Файлы не загружены" });
      return;
    }

    const session = getSession();
    if (!session.masterBuffer) {
      res.status(400).json({ success: false, message: "Сначала загрузите мастер-файл" });
      return;
    }

    let totalRecords = 0;
    const allMatched: string[] = [];
    const allUnmatched: string[] = [];
    const reportNames: string[] = [];

    for (const file of files) {
      const result = await processReport(file.buffer, file.originalname);
      totalRecords += result.records;
      allMatched.push(...result.matchedEmployees);
      allUnmatched.push(...result.unmatchedEmployees);
      reportNames.push(file.originalname);
    }

    const uniqueMatched = [...new Set(allMatched)];
    const uniqueUnmatched = [...new Set(allUnmatched)];

    res.json({
      success: true,
      message: `Обработано ${files.length} файлов, ${totalRecords} записей`,
      filesProcessed: files.length,
      totalRecords,
      reportNames,
      matchedEmployees: uniqueMatched,
      unmatchedEmployees: uniqueUnmatched,
    });
  } catch (err: any) {
    req.log.error({ err }, "Error uploading reports");
    res.status(500).json({ success: false, message: err.message || "Ошибка загрузки отчётов" });
  }
});

router.post("/payroll/process", async (req, res) => {
  try {
    const session = getSession();
    if (!session.masterBuffer) {
      res.status(400).json({ success: false, message: "Сначала загрузите мастер-файл" });
      return;
    }

    const results = await processPayroll();
    res.json({
      success: true,
      message: `Расчёт выполнен для ${results.length} сотрудников`,
      results,
    });
  } catch (err: any) {
    req.log.error({ err }, "Error processing payroll");
    res.status(500).json({ success: false, message: err.message || "Ошибка обработки" });
  }
});

router.get("/payroll/download", async (req, res) => {
  try {
    const session = getSession();
    if (!session.isProcessed) {
      res.status(400).json({ success: false, message: "Сначала выполните расчёт" });
      return;
    }

    const buffer = await generateResultFile();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="FINAL_MASTER_PAYROLL.xlsx"');
    res.send(buffer);
  } catch (err: any) {
    req.log.error({ err }, "Error generating result file");
    res.status(500).json({ success: false, message: err.message || "Ошибка генерации файла" });
  }
});

router.get("/payroll/status", async (req, res) => {
  try {
    const session = getSession();
    const monthInfo = session.month ? getMonthInfo(session.month) : null;

    res.json({
      hasMaster: !!session.masterBuffer,
      month: session.month || null,
      monthName: monthInfo?.name || null,
      normHours: monthInfo?.normHours || null,
      employeeCount: session.employees.length,
      uploadedReports: session.uploadedReports,
      isProcessed: session.isProcessed,
      employees: session.employees.map((e) => ({
        fio: e.fio,
        rawFio: e.rawFio,
        status: e.status,
        dismissDate: e.dismissDate,
        salary: e.salary,
        existingHours: e.existingHours,
        totalExistingHours: e.totalExistingHours,
      })),
      results: session.results,
    });
  } catch (err: any) {
    req.log.error({ err }, "Error getting payroll status");
    res.status(500).json({ success: false, message: err.message || "Ошибка получения статуса" });
  }
});

router.post("/payroll/reset", async (req, res) => {
  try {
    resetSession();
    res.json({ success: true, message: "Сессия очищена" });
  } catch (err: any) {
    req.log.error({ err }, "Error resetting session");
    res.status(500).json({ success: false, message: err.message || "Ошибка сброса сессии" });
  }
});

export default router;
