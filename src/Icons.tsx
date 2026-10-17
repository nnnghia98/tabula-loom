import React from 'react';
import { FieldDataType } from './types';

export function IconUpload({ size = 15 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M2.5 10.5v2a1 1 0 0 0 1 1h9a1 1 0 0 0 1-1v-2" />
      <path d="M8 2.5v8" />
      <path d="M5 5.5 8 2.5l3 3" />
    </svg>
  );
}

export function IconDownload({ size = 15 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M2.5 10.5v2a1 1 0 0 0 1 1h9a1 1 0 0 0 1-1v-2" />
      <path d="M8 2.5v8" />
      <path d="M5 7.5 8 10.5l3-3" />
    </svg>
  );
}

export function IconPlus({ size = 14 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round">
      <path d="M8 3v10M3 8h10" />
    </svg>
  );
}

export function IconTrash({ size = 14 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M3 4.5h10" />
      <path d="M6.5 2.5h3" />
      <path d="M4.5 4.5v8.5a1 1 0 0 0 1 1h5a1 1 0 0 0 1-1V4.5" />
    </svg>
  );
}

export function IconEye({ visible, size = 14 }: { visible: boolean; size?: number }) {
  if (visible) {
    return (
      <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
        <path d="M1.5 8s2.5-4.5 6.5-4.5S14.5 8 14.5 8s-2.5 4.5-6.5 4.5S1.5 8 1.5 8Z" />
        <circle cx="8" cy="8" r="2" />
      </svg>
    );
  }
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M2 2l12 12" />
      <path d="M6.5 6.5a2 2 0 0 0 2.8 2.8" />
      <path d="M3.5 5.5C2.3 6.6 1.5 8 1.5 8s2.5 4.5 6.5 4.5c1.4 0 2.6-.5 3.6-1.2" />
      <path d="M10.5 4.1C9.7 3.7 8.9 3.5 8 3.5c-.6 0-1.1.1-1.6.2" />
    </svg>
  );
}

export function IconSearch({ size = 14 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
      <circle cx="7" cy="7" r="4.5" />
      <path d="m10.5 10.5 3 3" />
    </svg>
  );
}

export function IconSliders({ size = 14 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
      <path d="M2.5 4.5h11M2.5 8h11M2.5 11.5h11" />
      <circle cx="5.5" cy="4.5" r="1.5" fill="var(--bg-surface)" />
      <circle cx="10.5" cy="8" r="1.5" fill="var(--bg-surface)" />
      <circle cx="6.5" cy="11.5" r="1.5" fill="var(--bg-surface)" />
    </svg>
  );
}

export function IconTable({ size = 14 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <rect x="2" y="2.5" width="12" height="11" rx="1.5" />
      <path d="M2 6.5h12M2 10h12M6.5 6.5v7" />
    </svg>
  );
}

export function IconLink({ size = 13 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M6.5 9.5a3 3 0 0 0 4.2 0l2-2a3 3 0 0 0-4.2-4.2l-1.1 1" />
      <path d="M9.5 6.5a3 3 0 0 0-4.2 0l-2 2a3 3 0 0 0 4.2 4.2l1.1-1" />
    </svg>
  );
}

export function IconWarning({ size = 13 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
      <path d="M8 2.2 1.6 13.2h12.8L8 2.2Z" />
      <path d="M8 6.5v3.2" />
      <circle cx="8" cy="11.5" r="0.6" fill="currentColor" />
    </svg>
  );
}

export function IconCheck({ size = 13 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <path d="M3 8.5 6.5 12 13 4.5" />
    </svg>
  );
}

export function IconDatabase({ size = 14 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <ellipse cx="8" cy="3.5" rx="5.5" ry="2" />
      <path d="M2.5 3.5v9c0 1.1 2.5 2 5.5 2s5.5-.9 5.5-2v-9" />
      <path d="M2.5 8c0 1.1 2.5 2 5.5 2s5.5-.9 5.5-2" />
    </svg>
  );
}

export function FieldTypeGlyph({ type }: { type: FieldDataType }) {
  switch (type) {
    case 'text':
      return <span className="type-glyph type-text">Aa</span>;
    case 'number':
      return <span className="type-glyph type-number">#</span>;
    case 'date':
      return <span className="type-glyph type-date">Cal</span>;
    case 'select':
      return <span className="type-glyph type-select">Opt</span>;
    case 'boolean':
      return <span className="type-glyph type-boolean">T/F</span>;
    case 'formula':
      return <span className="type-glyph type-formula">fx</span>;
    case 'lookup':
      return <span className="type-glyph type-lookup">Ref</span>;
  }
}
