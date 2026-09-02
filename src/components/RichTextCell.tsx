import React, { useRef, useEffect, useCallback } from "react";

interface RichTextCellProps {
  value: string;
  onChange: (value: string) => void;
  onFocus?: () => void;
  onBlur?: () => void;
  onKeyDown?: (e: React.KeyboardEvent<HTMLDivElement>) => void;
  onPaste?: (e: React.ClipboardEvent<HTMLDivElement>) => void;
  onClick?: (e: React.MouseEvent<HTMLDivElement>) => void;
  style?: React.CSSProperties;
  className?: string;
  dataRow?: number;
  dataCol?: number;
  placeholder?: string;
  readOnly?: boolean;
}

export const RichTextCell: React.FC<RichTextCellProps> = ({
  value,
  onChange,
  onFocus,
  onBlur,
  onKeyDown,
  onPaste,
  onClick,
  style,
  className,
  dataRow,
  dataCol,
  placeholder,
  readOnly = false,
}) => {
  const editorRef = useRef<HTMLDivElement>(null);
  const isComposingRef = useRef(false);

  // Synchronize external value with innerHTML when element is not active or initially loaded
  useEffect(() => {
    if (editorRef.current) {
      const currentHTML = editorRef.current.innerHTML;
      const normalizedValue = value || "";
      if (currentHTML !== normalizedValue && document.activeElement !== editorRef.current) {
        editorRef.current.innerHTML = normalizedValue;
      }
    }
  }, [value]);

  const handleInput = useCallback(() => {
    if (editorRef.current && !isComposingRef.current) {
      const html = editorRef.current.innerHTML;
      onChange(html);
    }
  }, [onChange]);

  return (
    <div
      ref={editorRef}
      contentEditable={!readOnly}
      suppressContentEditableWarning
      data-row={dataRow}
      data-col={dataCol}
      onInput={handleInput}
      onFocus={onFocus}
      onBlur={() => {
        if (editorRef.current) {
          onChange(editorRef.current.innerHTML);
        }
        onBlur?.();
      }}
      onClick={onClick}
      onKeyDown={onKeyDown}
      onPaste={onPaste}
      onCompositionStart={() => {
        isComposingRef.current = true;
      }}
      onCompositionEnd={() => {
        isComposingRef.current = false;
        handleInput();
      }}
      style={{
        outline: "none",
        minHeight: "1.2em",
        wordBreak: "break-word",
        whiteSpace: "pre-wrap",
        ...style,
      }}
      className={className}
      placeholder={placeholder}
    />
  );
};

export default RichTextCell;
