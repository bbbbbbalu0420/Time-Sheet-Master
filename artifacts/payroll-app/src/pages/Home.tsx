import React, { useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { 
  Calculator, 
  FileDown, 
  Trash2, 
  ChevronRight,
  AlertTriangle,
  PlayCircle,
  Users
} from "lucide-react";
import { usePayrollState, usePayrollActions } from "@/hooks/use-payroll";
import { Dropzone } from "@/components/ui/Dropzone";
import { StepIndicator } from "@/components/payroll/StepIndicator";
import { ResultsTable } from "@/components/payroll/ResultsTable";
import { formatCurrency } from "@/lib/utils";

const MONTHS = [
  { value: 1, label: "Январь" }, { value: 2, label: "Февраль" },
  { value: 3, label: "Март" }, { value: 4, label: "Апрель" },
  { value: 5, label: "Май" }, { value: 6, label: "Июнь" },
  { value: 7, label: "Июль" }, { value: 8, label: "Август" },
  { value: 9, label: "Сентябрь" }, { value: 10, label: "Октябрь" },
  { value: 11, label: "Ноябрь" }, { value: 12, label: "Декабрь" },
];

export default function Home() {
  const { data: status, isLoading } = usePayrollState();
  const actions = usePayrollActions();

  // Local state for forms
  const [masterFile, setMasterFile] = useState<File | null>(null);
  const [selectedMonth, setSelectedMonth] = useState<number>(new Date().getMonth() + 1);
  const [reportFiles, setReportFiles] = useState<File[]>([]);

  if (isLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-background">
        <div className="w-16 h-16 border-4 border-primary border-t-transparent rounded-full animate-spin"></div>
      </div>
    );
  }

  const currentStep = 
    status?.isProcessed ? 4 : 
    (status?.hasMaster && status.uploadedReports.length > 0) ? 3 :
    status?.hasMaster ? 2 : 1;

  const handleUploadMaster = () => {
    if (!masterFile) return;
    actions.uploadMaster.mutate({ data: { file: masterFile, month: selectedMonth } });
  };

  const handleUploadReports = () => {
    if (reportFiles.length === 0) return;
    actions.uploadReports.mutate({ data: { files: reportFiles } }, {
      onSuccess: () => setReportFiles([])
    });
  };

  const totalCalculatedSalary = status?.results?.reduce((sum, r) => sum + r.salary, 0) || 0;

  return (
    <div className="min-h-screen bg-background relative overflow-hidden text-foreground selection:bg-primary/30">
      {/* Background Image & Effects */}
      <div className="fixed inset-0 z-0">
        <img 
          src={`${import.meta.env.BASE_URL}images/fintech-bg.png`}
          alt="Premium Dark Background"
          className="w-full h-full object-cover opacity-40 mix-blend-overlay"
        />
        <div className="absolute inset-0 bg-background/80 backdrop-blur-3xl"></div>
        <div className="absolute inset-0 bg-grid-pattern opacity-20"></div>
        
        {/* Glow Effects */}
        <div className="absolute top-[-10%] right-[-5%] w-[40%] h-[40%] bg-primary/20 rounded-full blur-[120px]"></div>
        <div className="absolute bottom-[-10%] left-[-10%] w-[50%] h-[50%] bg-blue-500/10 rounded-full blur-[150px]"></div>
      </div>

      <div className="relative z-10 max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-12">
        
        {/* Header */}
        <header className="flex flex-col md:flex-row items-start md:items-center justify-between mb-12 gap-6">
          <div>
            <div className="inline-flex items-center gap-3 bg-white/5 border border-white/10 px-4 py-2 rounded-full mb-4 shadow-lg backdrop-blur-md">
              <Calculator className="w-5 h-5 text-primary" />
              <span className="text-sm font-bold tracking-widest uppercase text-muted-foreground">Система расчета ЗП</span>
            </div>
            <h1 className="text-4xl md:text-5xl font-extrabold bg-clip-text text-transparent bg-gradient-to-r from-white to-white/60 drop-shadow-sm">
              ММ Расчет Графика
            </h1>
          </div>

          {status?.hasMaster && (
            <button 
              onClick={() => actions.resetSession.mutate()}
              className="flex items-center gap-2 px-5 py-2.5 rounded-xl bg-rose-500/10 text-rose-400 border border-rose-500/20 hover:bg-rose-500/20 transition-all font-semibold shadow-lg backdrop-blur-md"
            >
              <Trash2 className="w-4 h-4" />
              Сбросить сессию
            </button>
          )}
        </header>

        {/* Step Indicator */}
        <div className="mb-16 mt-8 max-w-3xl mx-auto">
          <StepIndicator currentStep={Math.min(currentStep, 3)} />
        </div>

        {/* Dynamic Content */}
        <div className="space-y-8">
          <AnimatePresence mode="popLayout">
            
            {/* STEP 1: Master File */}
            {!status?.hasMaster && (
              <motion.div
                key="step-1"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95 }}
                className="glass-panel rounded-3xl p-8 md:p-12 max-w-3xl mx-auto"
              >
                <div className="text-center mb-8">
                  <h2 className="text-2xl font-bold mb-3">Загрузите Мастер-файл</h2>
                  <p className="text-muted-foreground">Шаблон графика с сотрудниками и статусами (ТРУД, ОТПУСК и др.)</p>
                </div>

                <div className="space-y-6">
                  <div>
                    <label className="block text-sm font-semibold mb-2 text-white/80">Выберите месяц для расчета (2026 год)</label>
                    <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-6 gap-2">
                      {MONTHS.map(m => (
                        <button
                          key={m.value}
                          onClick={() => setSelectedMonth(m.value)}
                          className={`py-2 px-1 text-sm font-medium rounded-xl border transition-all ${
                            selectedMonth === m.value 
                            ? 'bg-primary text-primary-foreground border-primary shadow-lg shadow-primary/25' 
                            : 'bg-white/5 border-white/10 text-muted-foreground hover:bg-white/10'
                          }`}
                        >
                          {m.label.substring(0, 3)}
                        </button>
                      ))}
                    </div>
                  </div>

                  <Dropzone 
                    accept={{ 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }}
                    maxFiles={1}
                    value={masterFile ? [masterFile] : []}
                    onDropFiles={(files) => setMasterFile(files[0])}
                    onRemove={() => setMasterFile(null)}
                    title="Мастер-таблица (.xlsx)"
                  />

                  <button
                    onClick={handleUploadMaster}
                    disabled={!masterFile || actions.uploadMaster.isPending}
                    className="w-full flex items-center justify-center gap-2 py-4 rounded-xl bg-gradient-to-r from-primary to-emerald-400 text-primary-foreground font-bold text-lg shadow-[0_0_20px_rgba(16,185,129,0.3)] hover:shadow-[0_0_30px_rgba(16,185,129,0.5)] disabled:opacity-50 disabled:cursor-not-allowed transition-all active:scale-[0.98]"
                  >
                    {actions.uploadMaster.isPending ? "Загрузка..." : "Продолжить"}
                    <ChevronRight className="w-5 h-5" />
                  </button>
                </div>
              </motion.div>
            )}

            {/* STEP 2: Reports */}
            {status?.hasMaster && (
              <motion.div
                key="step-2"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                className="grid lg:grid-cols-3 gap-8"
              >
                {/* Left Col: Context & Status */}
                <div className="lg:col-span-1 space-y-6">
                  <div className="glass-panel rounded-3xl p-6">
                    <h3 className="text-lg font-bold text-white/90 mb-4 flex items-center gap-2">
                      <Users className="w-5 h-5 text-primary" />
                      Текущая сессия
                    </h3>
                    <div className="space-y-4">
                      <div className="flex justify-between items-center pb-4 border-b border-white/10">
                        <span className="text-muted-foreground text-sm">Месяц</span>
                        <span className="font-bold text-primary">{status.monthName}</span>
                      </div>
                      <div className="flex justify-between items-center pb-4 border-b border-white/10">
                        <span className="text-muted-foreground text-sm">Норма часов</span>
                        <span className="font-bold font-mono">{status.normHours} ч.</span>
                      </div>
                      <div className="flex justify-between items-center pb-4 border-b border-white/10">
                        <span className="text-muted-foreground text-sm">Сотрудников</span>
                        <span className="font-bold font-mono">{status.employeeCount}</span>
                      </div>
                      <div className="flex justify-between items-center">
                        <span className="text-muted-foreground text-sm">Загружено отчетов</span>
                        <span className="font-bold font-mono text-white bg-white/10 px-2 py-1 rounded-md">
                          {status.uploadedReports.length}
                        </span>
                      </div>
                    </div>
                  </div>

                  {status.uploadedReports.length > 0 && (
                    <div className="glass-panel rounded-3xl p-6">
                      <h3 className="text-sm font-bold text-muted-foreground mb-3 uppercase tracking-wider">Файлы отчетов</h3>
                      <ul className="space-y-2">
                        {status.uploadedReports.map((r, i) => (
                          <li key={i} className="text-sm text-white/80 flex items-center gap-2 before:w-1.5 before:h-1.5 before:rounded-full before:bg-primary">
                            {r}
                          </li>
                        ))}
                      </ul>
                    </div>
                  )}
                </div>

                {/* Right Col: Actions */}
                <div className="lg:col-span-2 space-y-6">
                  
                  {/* Results preview if processed */}
                  {status.isProcessed && status.results.length > 0 && (
                    <div className="glass-panel rounded-3xl p-6 sm:p-8 border-primary/30 shadow-[0_0_40px_rgba(16,185,129,0.1)]">
                      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-6 mb-8">
                        <div>
                          <h2 className="text-2xl font-bold mb-1">Итоги расчета</h2>
                          <p className="text-muted-foreground">Итоговый ФОТ: <span className="text-foreground font-mono font-bold">{formatCurrency(totalCalculatedSalary)}</span></p>
                        </div>
                        <button
                          onClick={actions.handleDownload}
                          className="flex items-center gap-2 px-6 py-3 rounded-xl bg-primary text-primary-foreground font-bold shadow-lg shadow-primary/25 hover:shadow-xl hover:shadow-primary/40 hover:-translate-y-0.5 active:translate-y-0 transition-all"
                        >
                          <FileDown className="w-5 h-5" />
                          Скачать результат
                        </button>
                      </div>

                      <ResultsTable results={status.results} />
                    </div>
                  )}

                  {/* Upload Reports Zone */}
                  <div className="glass-panel rounded-3xl p-6 sm:p-8">
                    <h2 className="text-xl font-bold mb-2">
                      {status.uploadedReports.length === 0 ? "Загрузите отчеты часов" : "Добавить еще отчеты"}
                    </h2>
                    <p className="text-muted-foreground mb-6">
                      Загрузите файлы <code>RESULT_PAYROLL_*.xlsx</code> для начисления часов.
                    </p>
                    
                    <Dropzone 
                      accept={{ 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xls'], 'text/csv': ['.csv'] }}
                      value={reportFiles}
                      onDropFiles={(files) => setReportFiles(prev => [...prev, ...files])}
                      onRemove={(f) => setReportFiles(prev => prev.filter(x => x !== f))}
                      title="Отчеты из системы (.xlsx, .csv)"
                      description="Можно выбрать сразу несколько файлов"
                    />

                    {reportFiles.length > 0 && (
                      <div className="mt-6 flex justify-end">
                        <button
                          onClick={handleUploadReports}
                          disabled={actions.uploadReports.isPending}
                          className="flex items-center gap-2 px-6 py-3 rounded-xl bg-white/10 hover:bg-white/20 text-white font-bold transition-all border border-white/10"
                        >
                          {actions.uploadReports.isPending ? "Загрузка..." : "Применить отчеты"}
                        </button>
                      </div>
                    )}
                  </div>

                  {/* Process Trigger */}
                  {status.uploadedReports.length > 0 && !status.isProcessed && (
                    <motion.div 
                      initial={{ opacity: 0, scale: 0.95 }}
                      animate={{ opacity: 1, scale: 1 }}
                      className="glass-panel rounded-3xl p-8 border-warning/30 bg-warning/5"
                    >
                      <div className="flex flex-col sm:flex-row items-center justify-between gap-6">
                        <div>
                          <h3 className="text-xl font-bold text-white mb-2 flex items-center gap-2">
                            <AlertTriangle className="w-5 h-5 text-warning" />
                            Данные готовы к расчету
                          </h3>
                          <p className="text-muted-foreground text-sm max-w-md">
                            Загружено отчетов: {status.uploadedReports.length}. Нажмите кнопку для применения бизнес-правил, учета отпусков и подсчета зарплаты.
                          </p>
                        </div>
                        <button
                          onClick={() => actions.processPayroll.mutate()}
                          disabled={actions.processPayroll.isPending}
                          className="flex-shrink-0 flex items-center gap-2 px-8 py-4 rounded-xl bg-gradient-to-r from-warning to-amber-500 text-warning-foreground font-bold text-lg shadow-[0_0_20px_rgba(245,158,11,0.3)] hover:shadow-[0_0_30px_rgba(245,158,11,0.5)] transition-all active:scale-[0.98]"
                        >
                          <PlayCircle className="w-6 h-6" />
                          {actions.processPayroll.isPending ? "Считаем..." : "Рассчитать ЗП"}
                        </button>
                      </div>
                    </motion.div>
                  )}

                  {status.isProcessed && (
                    <div className="text-center pt-4 pb-8">
                       <button
                          onClick={() => actions.processPayroll.mutate()}
                          disabled={actions.processPayroll.isPending}
                          className="text-sm text-muted-foreground hover:text-white underline underline-offset-4 transition-colors"
                        >
                          Пересчитать (если добавлены новые отчеты)
                        </button>
                    </div>
                  )}
                  
                </div>
              </motion.div>
            )}
            
          </AnimatePresence>
        </div>
      </div>
    </div>
  );
}
