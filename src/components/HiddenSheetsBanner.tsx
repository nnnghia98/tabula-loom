import type { SheetData } from '../types';
import { IconEye, IconWarning } from '../Icons';

interface HiddenSheetsBannerProps {
  dismissedHiddenNotice: boolean;
  hiddenSheets: SheetData[];
  setDismissedHiddenNotice: React.Dispatch<React.SetStateAction<boolean>>;
  setShowHiddenSheets: React.Dispatch<React.SetStateAction<boolean>>;
  showHiddenSheets: boolean;
}

export function HiddenSheetsBanner({ dismissedHiddenNotice, hiddenSheets, setDismissedHiddenNotice, setShowHiddenSheets, showHiddenSheets }: HiddenSheetsBannerProps) {
  return (
    <>
      {hiddenSheets.length > 0 && !dismissedHiddenNotice && (
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '10px 20px',
            backgroundColor: 'var(--pastel-yellow-bg)',
            borderBottom: '1px solid #E5D4A0',
            color: 'var(--pastel-yellow-fg)',
            fontSize: '14px',
            flexShrink: 0,
            gap: '12px',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
            <IconWarning size={16} />
            <span>
              <strong>
                This Excel file contains {hiddenSheets.length} hidden{' '}
                {hiddenSheets.length === 1 ? 'sheet' : 'sheets'}:
              </strong>{' '}
              {hiddenSheets.map((hs) => hs.name).join(', ')}
            </span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
            <button
              type="button"
              className="btn btn-sm"
              style={{
                backgroundColor: '#FFFFFF',
                borderColor: '#D6C080',
                color: '#6E4800',
              }}
              onClick={() => setShowHiddenSheets((prev) => !prev)}
            >
              <IconEye visible={showHiddenSheets} size={14} />
              {showHiddenSheets ? 'Hide Hidden Sheets' : 'Show Hidden Sheets'}
            </button>
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              style={{ color: '#6E4800' }}
              onClick={() => setDismissedHiddenNotice(true)}
            >
              Dismiss
            </button>
          </div>
        </div>
      )}
    </>
  );
}
