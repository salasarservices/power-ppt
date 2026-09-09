import { useState } from "react";
import { useMutation, useQuery } from "@tanstack/react-query";
import { Download, RotateCcw, Loader2 } from "lucide-react";
import {
  analyzePptx,
  generateDeck,
  getHealth,
  type AnalyzeResponse,
  type SlidePlan,
} from "@/lib/api";
import { Uploader } from "@/components/Uploader";
import { PlanReview } from "@/components/PlanReview";
import { Button } from "@/components/ui/button";
import { useToast } from "@/components/ui/toast";

const LOGO = "https://ik.imagekit.io/salasarservices/Salasar-Logo-new.png";

function errMessage(e: unknown, fallback: string): string {
  const detail = (e as { response?: { data?: { detail?: string } } })?.response
    ?.data?.detail;
  return typeof detail === "string" ? detail : fallback;
}

export default function App() {
  const toast = useToast();
  const health = useQuery({ queryKey: ["health"], queryFn: getHealth });

  const [result, setResult] = useState<AnalyzeResponse | null>(null);
  const [plan, setPlan] = useState<SlidePlan | null>(null);
  const [sourceName, setSourceName] = useState<string>("");

  const analyze = useMutation({
    mutationFn: (file: File) => analyzePptx(file),
    onSuccess: (data, file) => {
      setResult(data);
      setPlan(data.plan);
      setSourceName(file.name.replace(/\.pptx$/i, ""));
    },
    onError: (e) =>
      toast({ kind: "error", message: errMessage(e, "Could not analyse that deck.") }),
  });

  const generate = useMutation({
    mutationFn: (p: SlidePlan) => generateDeck(p),
    onSuccess: (blob) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${sourceName || "deck"} — standardised.pptx`;
      a.click();
      URL.revokeObjectURL(url);
      toast({ kind: "success", message: "Standardised deck downloaded." });
    },
    onError: (e) =>
      toast({ kind: "error", message: errMessage(e, "Could not generate the deck.") }),
  });

  function reset() {
    setResult(null);
    setPlan(null);
    setSourceName("");
  }

  return (
    <div className="flex min-h-screen flex-col">
      <header className="flex items-center justify-between border-b bg-white px-6 py-4">
        <div>
          <h1 className="text-xl font-semibold text-brand-blue">
            Power<span className="text-brand-green">PPT</span>
          </h1>
          <p className="text-sm text-brand-slate">
            Reformat any deck into the authorised Salasar template.
          </p>
        </div>
        <img src={LOGO} alt="Salasar Services" className="h-10 w-auto" />
      </header>

      <main className="mx-auto w-full max-w-5xl flex-1 px-6 py-8">
        {!result ? (
          <>
            <Uploader onFile={(f) => analyze.mutate(f)} busy={analyze.isPending} />
            <p className="mt-4 text-center text-xs text-brand-slate">
              {health.isPending
                ? "Connecting to the service…"
                : health.isError
                  ? "Service unreachable — start the FastAPI backend on :8077."
                  : `Service ok · brand template ${health.data?.template_version}`}
            </p>
          </>
        ) : (
          <>
            <div className="mb-5 flex flex-wrap items-center justify-between gap-3">
              <div>
                <h2 className="text-lg font-semibold text-brand-blue">
                  Review — {result.slides} slide{result.slides === 1 ? "" : "s"}
                </h2>
                <p className="text-sm text-brand-slate">
                  Edit headings and body, then generate the branded deck. A named
                  person must review the output before it reaches a client.
                </p>
              </div>
              <div className="flex gap-2">
                <Button variant="outline" onClick={reset} disabled={generate.isPending}>
                  <RotateCcw className="h-4 w-4" /> Start over
                </Button>
                <Button
                  onClick={() => plan && generate.mutate(plan)}
                  disabled={generate.isPending || !plan}
                >
                  {generate.isPending ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <Download className="h-4 w-4" />
                  )}
                  Generate deck
                </Button>
              </div>
            </div>
            {plan && (
              <PlanReview plan={plan} warnings={result.warnings} onChange={setPlan} />
            )}
          </>
        )}
      </main>

      <footer className="mt-auto">
        <div className="mx-auto max-w-5xl px-6 pb-2 text-center text-xs text-brand-slate">
          Salasar Services (Insurance Brokers) Pvt. Ltd. · IRDAI Licence No. 143
          <br />
          Partnering a Secured Future
        </div>
        <div className="h-3 rounded-t-md bg-brand-bar" />
      </footer>
    </div>
  );
}
