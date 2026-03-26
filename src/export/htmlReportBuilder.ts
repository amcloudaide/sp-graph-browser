export function downloadHtmlReport(
  data: unknown,
  nodeLabel: string,
  breadcrumb: string[]
): void {
  const title = `SP Graph Browser — ${nodeLabel}`;
  const pathStr = breadcrumb.join(" > ");
  const timestamp = new Date().toISOString();

  let bodyContent: string;
  if (Array.isArray(data) && data.length > 0 && typeof data[0] === "object") {
    const columns = Object.keys(data[0] as Record<string, unknown>);
    const headerRow = columns.map((c) => `<th>${esc(c)}</th>`).join("");
    const rows = data
      .map(
        (item) =>
          "<tr>" +
          columns
            .map((c) => `<td>${esc(String((item as Record<string, unknown>)[c] ?? ""))}</td>`)
            .join("") +
          "</tr>"
      )
      .join("\n");
    bodyContent = `<table><thead><tr>${headerRow}</tr></thead><tbody>${rows}</tbody></table>`;
  } else {
    bodyContent = `<pre>${esc(JSON.stringify(data, null, 2))}</pre>`;
  }

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>${esc(title)}</title>
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; margin: 2rem; color: #333; }
  h1 { font-size: 1.4rem; }
  .meta { color: #666; font-size: 0.85rem; margin-bottom: 1.5rem; }
  table { border-collapse: collapse; width: 100%; font-size: 0.85rem; }
  th, td { border: 1px solid #ddd; padding: 6px 10px; text-align: left; }
  th { background: #f5f5f5; font-weight: 600; }
  tr:nth-child(even) { background: #fafafa; }
  pre { background: #f5f5f5; padding: 1rem; border-radius: 4px; overflow: auto; font-size: 0.8rem; }
</style>
</head>
<body>
<h1>${esc(title)}</h1>
<div class="meta">${esc(pathStr)} — exported ${timestamp}</div>
${bodyContent}
</body>
</html>`;

  const blob = new Blob([html], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${nodeLabel.replace(/[^a-zA-Z0-9]/g, "-")}-report.html`;
  a.click();
  URL.revokeObjectURL(url);
}

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
