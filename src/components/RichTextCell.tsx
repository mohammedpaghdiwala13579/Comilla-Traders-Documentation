import React, { useRef, useEffect, useCallback } from "react";
import { saveCurrentSelection } from "../utils/textFormatter";

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
      saveCurrentSelection();
    }
  }, [onChange]);

  // Handle typing to ensure formatting resets for subsequent words in the same cell
  const handleCellKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    onKeyDown?.(e);
    if (e.defaultPrevented) return;

    // When space is pressed at the boundary of a styled span (highlighted or colored text),
    // break out into an unstyled text node so the next word resets to clean default text
    if (e.key === " " && !e.shiftKey && !e.ctrlKey && !e.metaKey) {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && sel.isCollapsed) {
        const range = sel.getRangeAt(0);
        const node = range.startContainer;
        const styledSpan = (node instanceof HTMLElement ? node : node.parentElement)?.closest("span[style], b, i, u, mark, font");

        if (styledSpan && styledSpan !== editorRef.current && editorRef.current?.contains(styledSpan)) {
          const isAtEnd = node.nodeType === Node.TEXT_NODE && range.startOffset === (node.textContent?.length || 0);
          if (isAtEnd) {
            e.preventDefault();
            // Insert space after the styled span
            const spaceText = document.createTextNode("\u00A0");
            if (styledSpan.nextSibling) {
              styledSpan.parentNode?.insertBefore(spaceText, styledSpan.nextSibling);
            } else {
              styledSpan.parentNode?.appendChild(spaceText);
            }

            const newRange = document.createRange();
            newRange.setStartAfter(spaceText);
            newRange.collapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);

            handleInput();
            return;
          }
        }
      }
    }
  };

  const handleCellPaste = (e: React.ClipboardEvent<HTMLDivElement>) => {
    if (onPaste) {
      onPaste(e);
      if (e.defaultPrevented) return;
    }
    // Paste plain text to avoid foreign styling contamination
    e.preventDefault();
    const text = e.clipboardData.getData("text/plain");
    document.execCommand("insertText", false, text);
    handleInput();
  };

  return (
    <div
      ref={editorRef}
      contentEditable={!readOnly}
      suppressContentEditableWarning
      data-row={dataRow}
      data-col={dataCol}
      onInput={handleInput}
      onFocus={(e) => {
        saveCurrentSelection();
        onFocus?.();
      }}
      onBlur={() => {
        if (editorRef.current) {
          onChange(editorRef.current.innerHTML);
        }
        onBlur?.();
      }}
      onClick={(e) => {
        saveCurrentSelection();
        onClick?.(e);
      }}
      onMouseUp={() => {
        saveCurrentSelection();
      }}
      onKeyUp={() => {
        saveCurrentSelection();
      }}
      onKeyDown={handleCellKeyDown}
      onPaste={handleCellPaste}
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
