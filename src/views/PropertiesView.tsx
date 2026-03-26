import type { ReactNode } from "react";
import {
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Badge,
  Link,
} from "@fluentui/react-components";

interface PropertiesViewProps {
  data: unknown;
}

function formatValue(key: string, value: unknown): ReactNode {
  if (value === null || value === undefined) return <span style={{ color: "#666" }}>—</span>;
  if (typeof value === "boolean") {
    return <Badge appearance="filled" color={value ? "success" : "danger"}>{String(value)}</Badge>;
  }
  if (typeof value === "string") {
    if (value.startsWith("http://") || value.startsWith("https://")) {
      return <Link href={value} target="_blank">{value}</Link>;
    }
    // Highlight sharing capabilities
    if (key.toLowerCase().includes("sharing") && value.includes("External")) {
      return <Badge appearance="filled" color="warning">{value}</Badge>;
    }
  }
  if (typeof value === "object") return JSON.stringify(value);
  return String(value);
}

export function PropertiesView({ data }: PropertiesViewProps) {
  if (!data || typeof data !== "object") return <p>No data</p>;

  // If data is an array, show count and suggest table view
  if (Array.isArray(data)) {
    return <p>{data.length} items — switch to Table view for details</p>;
  }

  const entries = Object.entries(data as Record<string, unknown>).filter(
    ([key]) => !key.startsWith("@odata") && !key.startsWith("_")
  );

  return (
    <Table size="small" aria-label="Properties">
      <TableHeader>
        <TableRow>
          <TableHeaderCell style={{ width: "30%" }}>Property</TableHeaderCell>
          <TableHeaderCell>Value</TableHeaderCell>
        </TableRow>
      </TableHeader>
      <TableBody>
        {entries.map(([key, value]) => (
          <TableRow key={key}>
            <TableCell style={{ color: "var(--colorNeutralForeground3)", fontFamily: "monospace", fontSize: 12 }}>
              {key}
            </TableCell>
            <TableCell style={{ fontFamily: "monospace", fontSize: 12 }}>
              {formatValue(key, value)}
            </TableCell>
          </TableRow>
        ))}
      </TableBody>
    </Table>
  );
}
