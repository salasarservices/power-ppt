import { useState } from "react";
import { useMutation, useQuery } from "@tanstack/react-query";
import {
  analyzePptx,
  getHealth,
  layoutPlan,
  type Deck,
} from "@/lib/api";
import { Uploader } from "@/components/Uploader";
import { DeckEditor } from "@/components/DeckEditor";
import { useToast } from "@/components/ui/toast";

const LOGO = "https://ik.imagekit.io/salasarservices/Salasar-Logo-new.png";

function errMessage(e: unknown, fallback: string): string {
  const detail = (e as { response?: { data?: { detail?: string } } })?.response
    ?.data?.detail;
  return typeof detail === "string" ? detail : fallback;
}

interface Loaded {
  deck: Deck;
  warnings: string[];
  sourceName: string;
}

export default function App() {
  const toast = useToast();
  const health = useQuery({ queryKey: ["health"], queryFn: getHealth });
  const [loaded, setLoaded] = useState<Loaded | null>(null);

  const open = useMutation({
    mutationFn: async (file: File) => {
      const analysis = await analyzePptx(file);
      const deck = await layoutPlan(analysis.plan);
      return {
        deck,
        warnings: analysis.warnings,
        sourceName: file.name.replace(/\.pptx$/i, ""),
      } satisfies Loaded;
    },
    onSuccess: setLoaded,
    onError: (e) =>
      toast({ kind: "error", message: errMessage(e, "Could not open that deck.") }),
  });

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

      <main className="mx-auto w-full max-w-6xl flex-1 px-6 py-8">
        {!loaded ? (
          <>
            <Uploader onFile={(f) => open.mutate(f)} busy={open.isPending} />
            <p className="mt-4 text-center text-xs text-brand-slate">
              {health.isPending
                ? "Connecting to the service…"
                : health.isError
                  ? "Service unreachable — start the FastAPI backend on :8077."
                  : `Service ok · brand template ${health.data?.template_version}`}
            </p>
          </>
        ) : (
          <DeckEditor
            deck={loaded.deck}
            warnings={loaded.warnings}
            sourceName={loaded.sourceName}
            onReset={() => setLoaded(null)}
          />
        )}
      </main>

      <footer className="mt-auto">
        <div className="mx-auto max-w-6xl px-6 pb-2 text-center text-xs text-brand-slate">
          Salasar Services (Insurance Brokers) Pvt. Ltd. · IRDAI Licence No. 143
          <br />
          Partnering a Secured Future
        </div>
        <div className="h-3 rounded-t-md bg-brand-bar" />
      </footer>
    </div>
  );
}
