import { useState } from "react";
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogTrigger,
  Button,
  Input,
  Label,
  Select,
  SpinButton,
} from "@fluentui/react-components";
import { Settings20Regular } from "@fluentui/react-icons";
import type { AppSettings, ViewMode } from "../types";

interface SettingsDialogProps {
  settings: AppSettings;
  onSave: (settings: AppSettings) => void;
}

export function SettingsDialog({ settings, onSave }: SettingsDialogProps) {
  const [draft, setDraft] = useState(settings);
  const [open, setOpen] = useState(false);

  const handleOpen = () => {
    setDraft(settings);
    setOpen(true);
  };

  const handleSave = () => {
    onSave(draft);
    localStorage.setItem("sp-graph-browser-settings", JSON.stringify(draft));
    setOpen(false);
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => data.open ? handleOpen() : setOpen(false)}>
      <DialogTrigger>
        <Button icon={<Settings20Regular />} appearance="subtle" size="small" />
      </DialogTrigger>
      <DialogSurface>
        <DialogTitle>Settings</DialogTitle>
        <DialogBody>
          <div style={{ display: "flex", flexDirection: "column", gap: 16, padding: "16px 0" }}>
            <div>
              <Label htmlFor="ttl">Cache TTL (minutes)</Label>
              <SpinButton
                id="ttl"
                value={draft.cacheTtlMinutes}
                min={1}
                max={1440}
                onChange={(_, data) => setDraft({ ...draft, cacheTtlMinutes: data.value ?? 30 })}
              />
            </div>
            <div>
              <Label htmlFor="theme">Theme</Label>
              <Select
                id="theme"
                value={draft.theme}
                onChange={(_, data) => setDraft({ ...draft, theme: data.value as AppSettings["theme"] })}
              >
                <option value="system">System</option>
                <option value="light">Light</option>
                <option value="dark">Dark</option>
              </Select>
            </div>
            <div>
              <Label htmlFor="defaultView">Default View</Label>
              <Select
                id="defaultView"
                value={draft.defaultViewMode}
                onChange={(_, data) => setDraft({ ...draft, defaultViewMode: data.value as ViewMode })}
              >
                <option value="properties">Properties</option>
                <option value="json">Raw JSON</option>
                <option value="table">Table</option>
              </Select>
            </div>
            <div>
              <Label htmlFor="clientId">Custom Client ID (optional)</Label>
              <Input
                id="clientId"
                value={draft.customClientId ?? ""}
                onChange={(_, data) => setDraft({ ...draft, customClientId: data.value || null })}
                placeholder="Leave blank for default"
              />
            </div>
          </div>
        </DialogBody>
        <DialogActions>
          <DialogTrigger>
            <Button appearance="secondary">Cancel</Button>
          </DialogTrigger>
          <Button appearance="primary" onClick={handleSave}>Save</Button>
        </DialogActions>
      </DialogSurface>
    </Dialog>
  );
}
