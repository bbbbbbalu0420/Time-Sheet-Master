import React, { useState } from "react";
import {
  useReactTable,
  getCoreRowModel,
  getSortedRowModel,
  flexRender,
  createColumnHelper,
} from "@tanstack/react-table";
import { formatCurrency, cn } from "@/lib/utils";
import { ArrowUpDown, AlertCircle } from "lucide-react";

interface EmployeeResult {
  fio: string;
  status: string;
  totalHours: number;
  salary: number;
  overtime: number;
  dayHours: Record<string, number>;
  nightPay?: number;
  basePay?: number;
  overtimePay?: number;
  uncappedSalary?: number;
  employeeSalary?: number;
  normHours?: number;
  hourCost?: number;
}

interface ResultsTableProps {
  results: EmployeeResult[];
}

const columnHelper = createColumnHelper<EmployeeResult>();

const getStatusColor = (status: string) => {
  if (status === 'РАБОТАЕТ') return 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20';
  if (status === 'ОТПУСК') return 'bg-yellow-500/10 text-yellow-400 border-yellow-500/20';
  if (status === 'БОЛЬНИЧНЫЙ') return 'bg-yellow-500/10 text-yellow-400 border-yellow-500/20';
  if (status.includes('УВОЛЕН')) return 'bg-rose-500/10 text-rose-400 border-rose-500/20';
  if (status.includes('СВОЙ СЧ')) return 'bg-violet-500/10 text-violet-400 border-violet-500/20';
  return 'bg-white/5 text-muted-foreground border-white/10';
};

export function ResultsTable({ results }: ResultsTableProps) {
  const [sorting, setSorting] = useState([]);

  const columns = [
    columnHelper.accessor("fio", {
      header: "ФИО",
      cell: (info) => <span className="font-semibold text-foreground text-sm">{info.getValue()}</span>,
    }),
    columnHelper.accessor("status", {
      header: "Статус",
      cell: (info) => {
        const status = info.getValue();
        return (
          <span className={cn("px-2 py-0.5 text-[11px] font-bold rounded-full border whitespace-nowrap", getStatusColor(status))}>
            {status}
          </span>
        );
      },
    }),
    columnHelper.accessor("totalHours", {
      header: "Часы",
      cell: (info) => {
        const val = info.getValue();
        if (val <= 0) return <div className="text-right text-muted-foreground font-mono text-sm">0</div>;
        return <div className="text-right font-mono text-sm font-medium">{val.toFixed(1)}</div>;
      },
    }),
    columnHelper.accessor("overtime", {
      header: "Сверхур.",
      cell: (info) => {
        const val = info.getValue();
        if (val <= 0) return <div className="text-right text-muted-foreground font-mono text-sm">—</div>;
        return <div className="text-right text-amber-400 font-mono text-sm font-bold">+{val.toFixed(1)}</div>;
      },
    }),
    columnHelper.display({
      id: "nightPay",
      header: "Ночные",
      cell: (info) => {
        const val = info.row.original.nightPay;
        if (!val || val <= 0) return <div className="text-right text-muted-foreground font-mono text-xs">—</div>;
        return <div className="text-right font-mono text-xs text-blue-400">{formatCurrency(val)}</div>;
      },
    }),
    columnHelper.accessor("salary", {
      header: "К выплате",
      cell: (info) => {
        const val = info.getValue();
        const uncapped = info.row.original.uncappedSalary || 0;
        const isCapped = val >= 24500 && uncapped > val;
        return (
          <div className="flex items-center justify-end gap-1.5">
            {isCapped && <AlertCircle className="w-3.5 h-3.5 text-rose-400 flex-shrink-0" />}
            <span className={cn(
              "font-mono font-bold text-base whitespace-nowrap",
              isCapped ? "text-rose-400" : val > 0 ? "text-emerald-400" : "text-muted-foreground"
            )}>
              {formatCurrency(val)}
            </span>
          </div>
        );
      },
    }),
  ];

  const table = useReactTable({
    data: results,
    columns,
    state: { sorting },
    onSortingChange: setSorting,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
  });

  const totalSalary = results.reduce((s, r) => s + r.salary, 0);
  const totalHours = results.reduce((s, r) => s + r.totalHours, 0);
  const totalOvertime = results.reduce((s, r) => s + r.overtime, 0);
  const totalNight = results.reduce((s, r) => s + (r.nightPay || 0), 0);

  return (
    <div className="w-full overflow-hidden rounded-2xl border border-white/10 bg-card/40 backdrop-blur-md shadow-2xl">
      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse min-w-[650px]">
          <thead>
            {table.getHeaderGroups().map((headerGroup) => (
              <tr key={headerGroup.id} className="border-b border-white/10 bg-white/5">
                {headerGroup.headers.map((header) => (
                  <th
                    key={header.id}
                    className="px-3 py-3 text-xs font-semibold text-muted-foreground tracking-wider cursor-pointer hover:text-foreground transition-colors group whitespace-nowrap"
                    onClick={header.column.getToggleSortingHandler()}
                  >
                    <div className={cn("flex items-center gap-1", header.column.id === 'fio' || header.column.id === 'status' ? 'justify-start' : 'justify-end')}>
                      {flexRender(header.column.columnDef.header, header.getContext())}
                      <ArrowUpDown className="w-3 h-3 opacity-0 group-hover:opacity-50 transition-opacity" />
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {table.getRowModel().rows.map((row) => (
              <tr
                key={row.id}
                className="border-b border-white/5 hover:bg-white/[0.02] transition-colors"
              >
                {row.getVisibleCells().map((cell) => (
                  <td key={cell.id} className="px-3 py-2.5 align-middle">
                    {flexRender(cell.column.columnDef.cell, cell.getContext())}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className="border-t-2 border-white/20 bg-white/5">
              <td className="px-3 py-3 text-sm font-bold text-foreground" colSpan={2}>ИТОГО</td>
              <td className="px-3 py-3 text-right font-mono text-sm font-bold">{totalHours.toFixed(1)}</td>
              <td className="px-3 py-3 text-right font-mono text-sm font-bold text-amber-400">{totalOvertime > 0 ? `+${totalOvertime.toFixed(1)}` : '—'}</td>
              <td className="px-3 py-3 text-right font-mono text-xs font-bold text-blue-400">{totalNight > 0 ? formatCurrency(totalNight) : '—'}</td>
              <td className="px-3 py-3 text-right font-mono text-base font-bold text-emerald-400">{formatCurrency(totalSalary)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}
