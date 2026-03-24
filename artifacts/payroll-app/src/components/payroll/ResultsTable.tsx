import React, { useState } from "react";
import {
  useReactTable,
  getCoreRowModel,
  getSortedRowModel,
  flexRender,
  createColumnHelper,
} from "@tanstack/react-table";
import { formatCurrency, cn } from "@/lib/utils";
import type { EmployeeResult } from "@workspace/api-client-react";
import { ArrowUpDown, AlertCircle } from "lucide-react";

interface ResultsTableProps {
  results: EmployeeResult[];
}

const columnHelper = createColumnHelper<EmployeeResult>();

const getStatusColor = (status: string) => {
  if (status === 'РАБОТАЕТ') return 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20';
  if (status === 'ОТПУСК') return 'bg-blue-500/10 text-blue-400 border-blue-500/20';
  if (status === 'БОЛЬНИЧНЫЙ') return 'bg-amber-500/10 text-amber-400 border-amber-500/20';
  if (status.includes('УВОЛЕН')) return 'bg-rose-500/10 text-rose-400 border-rose-500/20';
  return 'bg-white/5 text-muted-foreground border-white/10';
};

export function ResultsTable({ results }: ResultsTableProps) {
  const [sorting, setSorting] = useState([]);

  const columns = [
    columnHelper.accessor("fio", {
      header: "Сотрудник (ФИО)",
      cell: (info) => <span className="font-semibold text-foreground">{info.getValue()}</span>,
    }),
    columnHelper.accessor("status", {
      header: "Статус",
      cell: (info) => {
        const status = info.getValue();
        return (
          <span className={cn("px-3 py-1 text-xs font-bold rounded-full border", getStatusColor(status))}>
            {status}
          </span>
        );
      },
    }),
    columnHelper.accessor("totalHours", {
      header: "Часы",
      cell: (info) => (
        <div className="font-mono text-right font-medium">
          {info.getValue().toFixed(1)} ч.
        </div>
      ),
    }),
    columnHelper.accessor("overtime", {
      header: "Переработка",
      cell: (info) => {
        const val = info.getValue();
        if (val <= 0) return <div className="text-right text-muted-foreground font-mono">-</div>;
        return (
          <div className="flex items-center justify-end gap-2 text-warning font-mono font-bold">
            +{val.toFixed(1)} ч.
          </div>
        );
      },
    }),
    columnHelper.accessor("salary", {
      header: "К выплате",
      cell: (info) => {
        const val = info.getValue();
        const isLimit = val >= 24500;
        return (
          <div className="flex items-center justify-end gap-2">
            {isLimit && (
              <AlertCircle className="w-4 h-4 text-rose-400" title="Достигнут лимит 24 500 руб." />
            )}
            <span className={cn(
              "font-mono font-bold text-lg tracking-tight",
              isLimit ? "text-rose-400" : "text-emerald-400"
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

  return (
    <div className="w-full overflow-hidden rounded-2xl border border-white/10 bg-card/40 backdrop-blur-md shadow-2xl">
      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse">
          <thead>
            {table.getHeaderGroups().map((headerGroup) => (
              <tr key={headerGroup.id} className="border-b border-white/10 bg-white/5">
                {headerGroup.headers.map((header) => (
                  <th 
                    key={header.id} 
                    className="p-4 text-sm font-semibold text-muted-foreground tracking-wider cursor-pointer hover:text-foreground transition-colors group"
                    onClick={header.column.getToggleSortingHandler()}
                  >
                    <div className={cn("flex items-center gap-2", header.column.id === 'fio' || header.column.id === 'status' ? 'justify-start' : 'justify-end')}>
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
                  <td key={cell.id} className="p-4 align-middle">
                    {flexRender(cell.column.columnDef.cell, cell.getContext())}
                  </td>
                ))}
              </tr>
            ))}
            {results.length === 0 && (
              <tr>
                <td colSpan={columns.length} className="p-12 text-center text-muted-foreground">
                  Нет данных для отображения
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
