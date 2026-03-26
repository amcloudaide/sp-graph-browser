export function toCsv(data: Record<string, unknown>[]): string {
  if (data.length === 0) return "";
  const columns = Object.keys(data[0]);
  const header = columns.join(",");
  const rows = data.map((row) =>
    columns.map((col) => {
      const val = String(row[col] ?? "");
      if (val.includes(",") || val.includes('"') || val.includes("\n")) {
        return `"${val.replace(/"/g, '""')}"`;
      }
      return val;
    }).join(",")
  );
  return [header, ...rows].join("\n");
}

export function downloadCsv(data: Record<string, unknown>[], filename: string): void {
  const csv = toCsv(data);
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${filename}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}
