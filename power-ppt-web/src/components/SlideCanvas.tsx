import { BRAND, HEADING, SLIDE, type DeckSlide, type Placement } from "@/lib/api";

const LOGO = "https://ik.imagekit.io/salasarservices/Salasar-Logo-new.png";

// Template geometry (inches) mirrored from the brand engine.
const LOGO_BOX = { left: 10.45, top: -0.28, width: 2.63, height: 1.86 };
const BAR = { insetFrac: 0.049, heightFrac: 0.045 };

const pctX = (inch: number) => `${(inch / SLIDE.wIn) * 100}%`;
const pctY = (inch: number) => `${(inch / SLIDE.hIn) * 100}%`;
// Font size relative to the container width so text scales with the canvas.
const fontCqw = (pt: number) => `${((pt / 72) / SLIDE.wIn) * 100}cqw`;

/** Heading brand split: primary term (up to & incl. the hyphen) blue, qualifier green. */
function HeadingText({ title }: { title: string }) {
  const t = title.toUpperCase();
  const i = t.indexOf("-");
  const primary = i >= 0 ? t.slice(0, i + 1) : t;
  const qualifier = i >= 0 ? t.slice(i + 1) : "";
  return (
    <span
      style={{
        fontFamily: "Poppins, sans-serif",
        fontWeight: 600,
        fontSize: fontCqw(BRAND.headingPt),
        lineHeight: 1.1,
      }}
    >
      <span style={{ color: BRAND.blue }}>{primary}</span>
      {qualifier && <span style={{ color: BRAND.green }}>{qualifier}</span>}
    </span>
  );
}

function PlacementView({ p }: { p: Placement }) {
  const style: React.CSSProperties = {
    position: "absolute",
    left: pctX(p.left),
    top: pctY(p.top),
    width: pctX(p.width),
  };

  if (p.kind === "body") {
    return (
      <div
        style={{
          ...style,
          color: BRAND.slate,
          fontFamily: "Poppins, sans-serif",
          fontSize: fontCqw(BRAND.bodyPt),
          lineHeight: 1.6,
          whiteSpace: "pre-wrap",
        }}
      >
        {p.text}
      </div>
    );
  }

  if (p.kind === "table" && p.table) {
    const rows = p.table.rows ?? [];
    return (
      <table
        style={{
          ...style,
          borderCollapse: "collapse",
          fontFamily: "Poppins, sans-serif",
          tableLayout: "fixed",
        }}
      >
        <tbody>
          {rows.map((row, ri) => (
            <tr key={ri}>
              {row.map((cell, ci) => (
                <td
                  key={ci}
                  style={{
                    border: "1px solid #E2E8F0",
                    padding: "0.4cqw 0.6cqw",
                    fontSize: fontCqw(ri === 0 ? 11 : 10),
                    fontWeight: ri === 0 ? 600 : 400,
                    color: ri === 0 ? "#FFFFFF" : BRAND.slate,
                    background: ri === 0 ? BRAND.blue : "#FFFFFF",
                  }}
                >
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
      <img
        src={`data:${p.image.content_type};base64,${p.image.data}`}
        alt=""
        style={{ ...style, height: "auto", objectFit: "contain" }}
      />
    );
  }
  return null;
}

export function SlideCanvas({ slide }: { slide: DeckSlide }) {
  return (
    <div
      style={{
        position: "relative",
        width: "100%",
        aspectRatio: `${SLIDE.wIn} / ${SLIDE.hIn}`,
        background: "#FFFFFF",
        containerType: "size",
        boxShadow: "0 1px 8px rgba(0,0,0,0.12)",
        overflow: "hidden",
        borderRadius: 4,
      }}
    >
      {/* logo top-right */}
      <img
        src={LOGO}
        alt="Salasar"
        style={{
          position: "absolute",
          left: pctX(LOGO_BOX.left),
          top: pctY(LOGO_BOX.top),
          width: pctX(LOGO_BOX.width),
          height: "auto",
        }}
      />

      {/* heading */}
      {slide.title && (
        <div style={{ position: "absolute", left: pctX(HEADING.left), top: pctY(HEADING.top), width: pctX(HEADING.width) }}>
          <HeadingText title={slide.title} />
        </div>
      )}

      {/* content placements */}
      {slide.placements.map((p, i) => (
        <PlacementView key={i} p={p} />
      ))}

      {/* footer bar */}
      <div
        style={{
          position: "absolute",
          left: pctX(SLIDE.wIn * BAR.insetFrac),
          width: pctX(SLIDE.wIn * (1 - 2 * BAR.insetFrac)),
          bottom: 0,
          height: pctY(SLIDE.hIn * BAR.heightFrac),
          background: `linear-gradient(90deg, ${BRAND.green}, ${BRAND.blue})`,
          borderTopLeftRadius: "0.4cqw",
          borderTopRightRadius: "0.4cqw",
        }}
      />
    </div>
  );
}
