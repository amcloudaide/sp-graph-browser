import { useState } from "react";
import { Button } from "@fluentui/react-components";
import { Copy20Regular, Checkmark20Regular } from "@fluentui/react-icons";

interface JsonViewProps {
  data: unknown;
}

export function JsonView({ data }: JsonViewProps) {
  const [copied, setCopied] = useState(false);
  const json = JSON.stringify(data, null, 2);

  const handleCopy = async () => {
    await navigator.clipboard.writeText(json);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div style={{ position: "relative" }}>
      <Button
        icon={copied ? <Checkmark20Regular /> : <Copy20Regular />}
        appearance="subtle"
        size="small"
        onClick={handleCopy}
        style={{ position: "absolute", top: 8, right: 8 }}
      >
        {copied ? "Copied" : "Copy"}
      </Button>
      <pre style={{
        fontSize: 12,
        fontFamily: "monospace",
        whiteSpace: "pre-wrap",
        wordBreak: "break-word",
        padding: 16,
        backgroundColor: "var(--colorNeutralBackground3)",
        borderRadius: 4,
        overflow: "auto",
        maxHeight: "calc(100vh - 120px)",
      }}>
        {json}
      </pre>
    </div>
  );
}
