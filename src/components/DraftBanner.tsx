
interface DraftBannerProps {
  handleDiscardDraft: () => void | Promise<void>;
  restoredDraftAt: number | null;
  setRestoredDraftAt: React.Dispatch<React.SetStateAction<number | null>>;
}

export function DraftBanner({ handleDiscardDraft, restoredDraftAt, setRestoredDraftAt }: DraftBannerProps) {
  return (
    <>
      {restoredDraftAt && (
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '10px 20px',
            backgroundColor: 'var(--pastel-blue-bg)',
            borderBottom: '1px solid var(--border-hairline)',
            color: 'var(--pastel-blue-fg)',
            fontSize: '14px',
            flexShrink: 0,
            gap: '12px',
          }}
        >
          <span>
            Restored your unsaved draft from {new Date(restoredDraftAt).toLocaleString()}. Press
            Save to keep it.
          </span>
          <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
            <button type="button" className="btn btn-sm" onClick={handleDiscardDraft}>
              Discard draft
            </button>
            <button type="button" className="btn btn-sm" onClick={() => setRestoredDraftAt(null)}>
              Dismiss
            </button>
          </div>
        </div>
      )}

    </>
  );
}
