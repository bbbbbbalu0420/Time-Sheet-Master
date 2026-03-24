import { useQueryClient } from "@tanstack/react-query";
import { 
  useGetPayrollStatus,
  useUploadMaster,
  useUploadReports,
  useProcessPayroll,
  useResetSession,
  getGetPayrollStatusQueryKey,
  downloadResult
} from "@workspace/api-client-react";
import { useToast } from "@/hooks/use-toast";

export function usePayrollState() {
  return useGetPayrollStatus({
    query: {
      retry: false,
      staleTime: 1000,
    }
  });
}

export function usePayrollActions() {
  const queryClient = useQueryClient();
  const { toast } = useToast();
  
  const invalidateStatus = () => {
    queryClient.invalidateQueries({ queryKey: getGetPayrollStatusQueryKey() });
  };

  const uploadMasterMutation = useUploadMaster({
    mutation: {
      onSuccess: (data) => {
        invalidateStatus();
        toast({
          title: "Мастер-файл загружен",
          description: `Распознано ${data.employeeCount} сотрудников.`,
        });
      },
      onError: (error: any) => {
        toast({
          variant: "destructive",
          title: "Ошибка загрузки",
          description: error?.message || "Не удалось загрузить мастер-файл",
        });
      }
    }
  });

  const uploadReportsMutation = useUploadReports({
    mutation: {
      onSuccess: (data) => {
        invalidateStatus();
        toast({
          title: "Отчеты загружены",
          description: `Обработано ${data.filesProcessed} файлов, найдено ${data.totalRecords} записей.`,
        });
      },
      onError: (error: any) => {
        toast({
          variant: "destructive",
          title: "Ошибка загрузки",
          description: error?.message || "Не удалось загрузить отчеты",
        });
      }
    }
  });

  const processMutation = useProcessPayroll({
    mutation: {
      onSuccess: () => {
        invalidateStatus();
        toast({
          title: "Расчет завершен",
          description: "Зарплаты успешно рассчитаны.",
        });
      },
      onError: (error: any) => {
        toast({
          variant: "destructive",
          title: "Ошибка расчета",
          description: error?.message || "Не удалось выполнить расчет",
        });
      }
    }
  });

  const resetMutation = useResetSession({
    mutation: {
      onSuccess: () => {
        invalidateStatus();
        toast({
          title: "Сессия сброшена",
          description: "Все данные очищены, можно начать заново.",
        });
      }
    }
  });

  const handleDownload = async () => {
    try {
      const blob = await downloadResult();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "FINAL_MASTER_PAYROLL.xlsx";
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      toast({
        title: "Скачивание началось",
        description: "Ваш файл FINAL_MASTER_PAYROLL.xlsx готов.",
      });
    } catch (error) {
      toast({
        variant: "destructive",
        title: "Ошибка скачивания",
        description: "Не удалось скачать результат.",
      });
    }
  };

  return {
    uploadMaster: uploadMasterMutation,
    uploadReports: uploadReportsMutation,
    processPayroll: processMutation,
    resetSession: resetMutation,
    handleDownload
  };
}
