import { AlertTriangle } from "lucide-react";
import type { Page, SlidePlan, Table as TableT } from "@/lib/api";
import { Card, CardContent, CardHeader } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";

interface PlanReviewProps {
  plan: SlidePlan;
  warnings: string[];
  onChange: (plan: SlidePlan) => void;
}

export function PlanReview({ plan, warnings, onChange }: PlanReviewProps) {
  function patchPage(index: number, patch: Partial<Page>) {
    const pages = plan.pages.map((p, i) =>
      i === index ? { ...p, ...patch } : p,
    );
    onChange({ pages });
  }

  return (
    <div className="space-y-5">
      {warnings.length > 0 && (
        <div className="flex gap-3 rounded-md border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-800">
          <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" />
          <ul className="list-disc space-y-1 pl-4">
            {warnings.map((w, i) => (
              <li key={i}>{w}</li>
            ))}
          </ul>
        </div>
      )}

      {plan.pages.map((page, i) => (
        <Card key={i}>
          <CardHeader className="flex items-center justify-between">
            <span className="text-sm font-semibold text-brand-blue">
              Slide {i + 1}
            </span>
            {page.tables.length > 0 && (
              <span className="text-xs text-brand-slate">
                {page.tables.length} table{page.tables.length > 1 ? "s" : ""}
              </span>
            )}
          </CardHeader>
          <CardContent className="space-y-3">
            <div>
              <label className="mb-1 block text-xs font-medium text-brand-slate">
                Heading (leave blank for no heading)
              </label>
              <Input
                value={page.title}
                placeholder="Untitled slide"
                onChange={(e) => patchPage(i, { title: e.target.value })}
              />
            </div>
            <div>
              <label className="mb-1 block text-xs font-medium text-brand-slate">
                Body
              </label>
              <Textarea
                value={page.body}
                rows={Math.min(10, Math.max(3, page.body.split("\n").length))}
                placeholder="Slide body text…"
                onChange={(e) => patchPage(i, { body: e.target.value })}
              />
            </div>
            {page.tables.map((t, ti) => (
              <TablePreview key={ti} table={t} />
            ))}
          </CardContent>
        </Card>
      ))}
    </div>
  );
}

// Tables are shown read-only in v1 — editing structured cells comes later.
function TablePreview({ table }: { table: TableT }) {
  if (table.rows.length === 0) return null;
  const [head, ...body] = table.rows;
  return (
    <div className="overflow-x-auto rounded-md border">
      <table className="w-full border-collapse text-sm">
        <thead className="bg-slate-50">
          <tr>
            {head.map((c, i) => (
              <th
                key={i}
                className="border-b px-3 py-1.5 text-left font-medium text-brand-blue"
              >
                {c}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {body.map((row, ri) => (
            <tr key={ri}>
              {row.map((c, ci) => (
                <td key={ci} className="border-b px-3 py-1.5 text-brand-slate">
                  {c}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
