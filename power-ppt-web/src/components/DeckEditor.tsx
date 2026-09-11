import { useCallback, useEffect, useState } from "react";
import { useMutation } from "@tanstack/react-query";
import { ChevronLeft, ChevronRight, Download, Loader2, RotateCcw } from "lucide-react";
import { renderDeck, type Deck } from "@/lib/api";
import { SlideCanvas } from "@/components/SlideCanvas";
import { Button } from "@/components/ui/button";
import { useToast } from "@/components/ui/toast";

interface DeckEditorProps {
  deck: Deck;
  warnings: string[];
  sourceName: string;
  onReset: () => void;
}

export function DeckEditor({ deck, warnings, sourceName, onReset }: DeckEditorProps) {
  const toast = useToast();
  const [idx, setIdx] = useState(0);
  const count = deck.slides.length;
  const clamped = Math.min(idx, count - 1);

  const go = useCallback(
    (d: number) => setIdx((i) => Math.max(0, Math.min(count - 1, i + d))),
    [count],
  );

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "ArrowLeft") go(-1);
      if (e.key === "ArrowRight") go(1);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [go]);

  const render = useMutation({
    mutationFn: () => renderDeck(deck),
    onSuccess: (blob) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${sourceName || "deck"} — standardised.pptx`;
      a.click();
      URL.revokeObjectURL(url);
      toast({ kind: "success", message: "Standardised deck downloaded." });
    },
    onError: () =>
      toast({ kind: "error", message: "Could not generate the deck." }),
  });

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-lg font-semibold text-brand-blue">
            Review &amp; edit — {count} slide{count === 1 ? "" : "s"}
          </h2>
          <p className="text-sm text-brand-slate">
            Navigate with the arrows or ← →. A named person must review the output
            before it reaches a client.
          </p>
        </div>
        <div className="flex gap-2">
          <Button variant="outline" onClick={onReset} disabled={render.isPending}>
            <RotateCcw className="h-4 w-4" /> Start over
          </Button>
          <Button onClick={() => render.mutate()} disabled={render.isPending}>
            {render.isPending ? (
              <Loader2 className="h-4 w-4 animate-spin" />
            ) : (
              <Download className="h-4 w-4" />
            )}
            Generate deck
          </Button>
        </div>
      </div>

      {warnings.length > 0 && (
        <ul className="list-disc rounded-md border border-amber-300 bg-amber-50 px-6 py-3 text-sm text-amber-800">
          {warnings.map((w, i) => (
            <li key={i}>{w}</li>
          ))}
        </ul>
      )}

      {/* canvas + side arrows */}
      <div className="flex items-center gap-3">
        <Button
          variant="outline"
          className="h-10 w-10 shrink-0 rounded-full p-0"
          onClick={() => go(-1)}
          disabled={clamped === 0}
          aria-label="Previous slide"
        >
          <ChevronLeft className="h-5 w-5" />
        </Button>

        <div className="flex-1">
          {deck.slides[clamped] && <SlideCanvas slide={deck.slides[clamped]} />}
        </div>

        <Button
          variant="outline"
          className="h-10 w-10 shrink-0 rounded-full p-0"
          onClick={() => go(1)}
          disabled={clamped === count - 1}
          aria-label="Next slide"
        >
          <ChevronRight className="h-5 w-5" />
        </Button>
      </div>

      <p className="text-center text-sm text-brand-slate">
        Slide {clamped + 1} of {count}
      </p>
    </div>
  );
}
