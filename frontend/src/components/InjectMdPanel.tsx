type Props = {
  templateHwpx: string | null;
  active: boolean;
  onClear: () => void;
  onSelect: () => void;
};

export default function InjectTargetPanel({ templateHwpx, active, onClear, onSelect }: Props) {
  const name = templateHwpx ? templateHwpx.split(/[\\/]/).pop() : null;
  const lower = (templateHwpx || "").toLowerCase();
  const isPptx = lower.endsWith(".pptx");
  const isHwpx = lower.endsWith(".hwpx");
  const tagClass = isPptx ? "ext-pptx" : isHwpx ? "ext-hwpx" : "";
  const tagLabel = isPptx ? "pptx" : isHwpx ? "hwpx" : "?";
  const title = isPptx ? "🎨 주입 문서 (템플릿 PPTX)" : "🎯 주입 문서 (템플릿 HWPX)";
  return (
    <div className="panel-section" style={{ borderTop: "2px solid var(--accent)" }}>
      <div className="panel-section-title" style={{ color: "var(--accent)" }}>
        {title}
      </div>
      {templateHwpx ? (
        <div
          className={`md-item ${active ? "selected" : ""}`}
          onClick={onSelect}
          title={templateHwpx}
          style={{ paddingLeft: 10 }}
        >
          <span className={`ext-tag ${tagClass}`}>{tagLabel}</span>
          <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis" }}>{name}</span>
          <button
            onClick={(e) => {
              e.stopPropagation();
              onClear();
            }}
            style={{ padding: "0 6px", fontSize: 10 }}
            title="템플릿 해제"
          >
            ×
          </button>
        </div>
      ) : (
        <div style={{ padding: "8px 10px", fontSize: 11, color: "var(--fg-dim)", lineHeight: 1.5 }}>
          HWPX / PPTX 파일 우클릭 →<br />
          <b style={{ color: "var(--accent)" }}>"★ 이 HWPX/PPTX를 템플릿으로 지정"</b>
        </div>
      )}
    </div>
  );
}
