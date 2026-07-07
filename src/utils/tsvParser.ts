// TSV state-machine parser to handle double quotes, internal cell newlines, and tabs
export function parseTSV(text: string): string[][] {
  const result: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          // Escaped double quote: "" -> "
          cell += '"';
          i++; // Skip next quote
        } else {
          // End of quoted cell
          inQuotes = false;
        }
      } else {
        cell += char;
      }
    } else {
      if (char === '"') {
        // Start of quoted cell
        inQuotes = true;
      } else if (char === '\t') {
        // End of column
        row.push(cell);
        cell = "";
      } else if (char === '\r') {
        // Carriage return: handle trailing \n
        if (nextChar === '\n') {
          row.push(cell);
          result.push(row);
          row = [];
          cell = "";
          i++; // Skip \n
        } else {
          row.push(cell);
          result.push(row);
          row = [];
          cell = "";
        }
      } else if (char === '\n') {
        // Newline
        row.push(cell);
        result.push(row);
        row = [];
        cell = "";
      } else {
        cell += char;
      }
    }
  }

  if (cell !== "" || row.length > 0) {
    row.push(cell);
    result.push(row);
  }

  // Clean up empty trailing row which is common when copying from Excel
  if (
    result.length > 1 &&
    result[result.length - 1].length === 1 &&
    result[result.length - 1][0] === ""
  ) {
    result.pop();
  }

  return result;
}
