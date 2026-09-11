import { useEffect, useRef, useState } from "react";
import {
  BRAND,
  HEADING,
  SLIDE,
  type DeckSlide,
  type Placement,
  type PptImage,
} from "@/lib/api";

const LOGO = "https://ik.imagekit.io/salasarservices/Salasar-Logo-new.png";
const LOGO_BOX = { left: 10.45, top: -0.28, width: 2.63, height: 1.86 };
const BAR = { insetFrac: 0.049, heightFrac: 0.045 };

const pctX = (inch: number) => `${(inch / SLIDE.wIn) * 100}%`;
const pctY = (inch: number) => `${(inch / SLIDE.hIn) * 100}%`;
const fontCqw = (pt: number) => `${((pt / 72) / SLIDE.wIn) * 100}cqw`;

export interface EditHandlers {
  editingIndex: number | null;
  onStartEdit: (i: number) => void;
  onChange: (i: number, p: Placement) => void;
  onDelete: (i: number) => void;
  onStopEdit: () => void;
}

function fileToImage(file: File): Promise<PptImage> {
  return new Promise((resolve) => {
    const r = new FileReader();
    r.onload = () => {
      const dataUrl = r.result as string;
      resolve({ data: dataUrl.split(",")[1], content_type: file.type || "image/png" });
    };
    r.readAsDataURL(file);
  });
}

function HeadingText({ title }: { title: string }) {
  const t = title.toUpperCase();
  const i = t.indexOf("-");
  const primary = i >= 0 ? t.slice(0, i + 1) : t;
  const qualifier = i >= 0 ? t.slice(i + 1) : "";
  return (
    <span style={{ fontFamily: "Poppins, sans-serif", fontWeight: 600, fontSize: fontCqw(BRAND.headingPt), lineHeight: 1.1 }}>
      <span style={{ color: BRAND.blue }}>{primary}</span>
      {qualifier && <span style={{ color: BRAND.green }}>{qualifier}</span>}
    </span>
  );
}

// Small floating chrome toolbar anchored above an object being edited.
function EditBar({ left, top, children }: { left: number; top: number; children: React.ReactNode }) {
  return (
    <div
      onDoubleClick={(e) => e.stopPropagation()}
      style={{ position: "absolute", left: pctX(left), top: pctY(top), transform: "translateY(-118%)", display: "flex", gap: 4, zIndex: 5 }}
    >
      {children}
    </div>
  );
}

const barBtn: React.CSSProperties = {
  fontSize: 10, lineHeight: 1, padding: "3px 6px", borderRadius: 4,
  border: "1px solid #CBD5E1", background: "#fff", color: BRAND.slate, cursor: "pointer",
};
const doneBtn: React.CSSProperties = { ...barBtn, background: BRAND.blue, color: "#fff", border: "none" };

function BodyEditor({ p, index, edit }: { p: Placement; index: number; edit: EditHandlers }) {
  const ref = useRef<HTMLTextAreaElement>(null);
  useEffect(() => ref.current?.focus(), []);
  return (
    <>
      <EditBar left={p.left} top={p.top}>
        <button style={doneBtn} onClick={edit.onStopEdit}>Done</button>
      </EditBar>
      <textarea
        ref={ref}
        className="pp-editing"
        value={p.text ?? ""}
        onChange={(e) => edit.onChange(index, { ...p, text: e.target.value })}
        onKeyDown={(e) => e.key === "Escape" && edit.onStopEdit()}
        style={{
          position: "absolute", left: pctX(p.left), top: pctY(p.top), width: pctX(p.width),
          height: pctY(p.height ?? 1), color: BRAND.slate, fontFamily: "Poppins, sans-serif",
          fontSize: fontCqw(BRAND.bodyPt), lineHeight: 1.6, border: "none", resize: "none",
          background: "rgba(0,112,192,0.04)", padding: 0, outline: "none",
        }}
      />
    </>
  );
}

function TableEditor({ p, index, edit }: { p: Placement; index: number; edit: EditHandlers }) {
  const rows = p.table?.rows ?? [];
  const commit = (nr: string[][]) =>
    edit.onChange(index, { ...p, table: { header: nr[0] ?? [], rows: nr } });
  const setCell = (r: number, c: number, v: string) => {
    const nr = rows.map((row) => [...row]);
    nr[r][c] = v;
    commit(nr);
  };
  const ncols = rows[0]?.length ?? 1;
  return (
    <>
      <EditBar left={p.left} top={p.top}>
        <button style={barBtn} onClick={() => commit([...rows, Array(ncols).fill("")])}>+Row</button>
        <button style={barBtn} onClick={() => rows.length > 1 && commit(rows.slice(0, -1))}>−Row</button>
        <button style={barBtn} onClick={() => commit(rows.map((r) => [...r, ""]))}>+Col</button>
        <button style={barBtn} onClick={() => ncols > 1 && commit(rows.map((r) => r.slice(0, -1)))}>−Col</button>
        <button style={doneBtn} onClick={edit.onStopEdit}>Done</button>
      </EditBar>
      <table className="pp-editing" style={{ position: "absolute", left: pctX(p.left), top: pctY(p.top), width: pctX(p.width), borderCollapse: "collapse", fontFamily: "Poppins, sans-serif", tableLayout: "fixed" }}>
        <tbody>
          {rows.map((row, ri) => (
            <tr key={ri}>
              {row.map((cell, ci) => (
                <td key={ci} style={{ border: "1px solid #E2E8F0", padding: 0, background: ri === 0 ? BRAND.blue : "#fff" }}>
                  <input
                    value={cell}
                    onChange={(e) => setCell(ri, ci, e.target.value)}
                    style={{ width: "100%", border: "none", outline: "none", padding: "0.4cqw 0.6cqw", fontFamily: "Poppins, sans-serif", fontSize: fontCqw(ri === 0 ? 11 : 10), fontWeight: ri === 0 ? 600 : 400, color: ri === 0 ? "#fff" : BRAND.slate, background: "transparent" }}
                  />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </>
  );
}

function ImageEditor({ p, index, edit }: { p: Placement; index: number; edit: EditHandlers }) {
  return (
    <>
      <EditBar left={p.left} top={p.top}>
        <label style={barBtn}>
          Replace
          <input type="file" accept="image/*" hidden onChange={async (e) => {
            const f = e.target.files?.[0];
            if (f) edit.onChange(index, { ...p, image: await fileToImage(f) });
          }} />
        </label>
        <button style={{ ...barBtn, color: "#DC2626" }} onClick={() => edit.onDelete(index)}>Remove</button>
        <button style={doneBtn} onClick={edit.onStopEdit}>Done</button>
      </EditBar>
      {p.image && (
        <img className="pp-editing" src={`data:${p.image.content_type};base64,${p.image.data}`} alt=""
          style={{ position: "absolute", left: pctX(p.left), top: pctY(p.top), width: pctX(p.width), height: "auto", objectFit: "contain" }} />
      )}
    </>
  );
}

function StaticPlacement({ p, index, edit }: { p: Placement; index: number; edit?: EditHandlers }) {
  const editable = edit ? "pp-editable" : "";
  const onDouble = edit ? () => edit.onStartEdit(index) : undefined;
  const base: React.CSSProperties = { position: "absolute", left: pctX(p.left), top: pctY(p.top), width: pctX(p.width) };
  const title = edit ? "Double-click to edit" : undefined;

  if (p.kind === "body") {
    return (
      <div className={editable} title={title} onDoubleClick={onDouble}
        style={{ ...base, color: BRAND.slate, fontFamily: "Poppins, sans-serif", fontSize: fontCqw(BRAND.bodyPt), lineHeight: 1.6, whiteSpace: "pre-wrap" }}>
        {p.text}
      </div>
    );
  }
  if (p.kind === "table" && p.table) {
    return (
      <table className={editable} title={title} onDoubleClick={onDouble}
        style={{ ...base, borderCollapse: "collapse", fontFamily: "Poppins, sans-serif", tableLayout: "fixed" }}>
        <tbody>
          {(p.table.rows ?? []).map((row, ri) => (
            <tr key={ri}>
              {row.map((cell, ci) => (
                <td key={ci} style={{ border: "1px solid #E2E8F0", padding: "0.4cqw 0.6cqw", fontSize: fontCqw(ri === 0 ? 11 : 10), fontWeight: ri === 0 ? 600 : 400, color: ri === 0 ? "#fff" : BRAND.slate, background: ri === 0 ? BRAND.blue : "#fff" }}>
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }
  if (p.kind === "image" && p.image) {
    return (
      <img className={editable} title={title} onDoubleClick={onDouble}
        src={`data:${p.image.content_type};base64,${p.image.data}`} alt=""
        style={{ ...base, height: "auto", objectFit: "contain" }} />
    );
  }
  return null;
}

function PlacementView({ p, index, edit }: { p: Placement; index: number; edit?: EditHandlers }) {
  if (edit && edit.editingIndex === index) {
    if (p.kind === "body") return <BodyEditor p={p} index={index} edit={edit} />;
    if (p.kind === "table") return <TableEditor p={p} index={index} edit={edit} />;
    if (p.kind === "image") return <ImageEditor p={p} index={index} edit={edit} />;
  }
  return <StaticPlacement p={p} index={index} edit={edit} />;
}

export function SlideCanvas({ slide, edit }: { slide: DeckSlide; edit?: EditHandlers }) {
  const [imgOk, setImgOk] = useState(true);
  return (
    <div style={{ position: "relative", width: "100%", aspectRatio: `${SLIDE.wIn} / ${SLIDE.hIn}`, background: "#FFFFFF", containerType: "size", boxShadow: "0 1px 8px rgba(0,0,0,0.12)", overflow: "hidden", borderRadius: 4 }}>
      {imgOk && (
        <img src={LOGO} alt="Salasar" onError={() => setImgOk(false)}
          style={{ position: "absolute", left: pctX(LOGO_BOX.left), top: pctY(LOGO_BOX.top), width: pctX(LOGO_BOX.width), height: "auto" }} />
      )}

      {slide.title && (
        <div style={{ position: "absolute", left: pctX(HEADING.left), top: pctY(HEADING.top), width: pctX(HEADING.width) }}>
          <HeadingText title={slide.title} />
        </div>
      )}

      {slide.placements.map((p, i) => (
        <PlacementView key={i} p={p} index={i} edit={edit} />
      ))}

      <div style={{ position: "absolute", left: pctX(SLIDE.wIn * BAR.insetFrac), width: pctX(SLIDE.wIn * (1 - 2 * BAR.insetFrac)), bottom: 0, height: pctY(SLIDE.hIn * BAR.heightFrac), background: `linear-gradient(90deg, ${BRAND.green}, ${BRAND.blue})`, borderTopLeftRadius: "0.4cqw", borderTopRightRadius: "0.4cqw" }} />
    </div>
  );
}
