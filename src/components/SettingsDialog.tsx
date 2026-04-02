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
  Checkbox,
  Body1,
} from "@fluentui/react-components";
import { Settings20Regular } from "@fluentui/react-icons";
import { useAuth } from "../auth/AuthProvider";
import { graphScopesFiles } from "../auth/msalConfig";
import type { AppSettings, ViewMode } from "../types";

interface SettingsDialogProps {
  settings: AppSettings;
  onSave: (settings: AppSettings) => void;
}

export function SettingsDialog({ settings, onSave }: SettingsDialogProps) {
  const { requestAdditionalScopes } = useAuth();
  const [draft, setDraft] = useState(settings);
  const [open, setOpen] = useState(false);
  const [consentPending, setConsentPending] = useState(false);

  const handleOpen = () => {
    setDraft(settings);
    setOpen(true);
  };

  const handleSave = async () => {
    // If files access was just enabled, request incremental consent
    if (draft.enableFilesAccess && !settings.enableFilesAccess) {
      setConsentPending(true);
      try {
        await requestAdditionalScopes(graphScopesFiles);
      } catch (e) {
        console.warn("Consent for Files.Read.All failed:", e);
        // Revert the toggle if consent was denied
        draft.enableFilesAccess = false;
      } finally {
        setConsentPending(false);
      }
    }
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
            <div>
              <Label htmlFor="proxyUrl">Proxy URL (optional)</Label>
              <Input
                id="proxyUrl"
                value={draft.proxyUrl ?? ""}
                onChange={(_, data) => setDraft({ ...draft, proxyUrl: data.value || null })}
                placeholder="https://your-function.azurewebsites.net"
              />
            </div>
            <div>
              <Label htmlFor="blobSasUrl">Blob Storage SAS URL (optional)</Label>
              <Input
                id="blobSasUrl"
                value={draft.blobSasUrl ?? ""}
                onChange={(_, data) => setDraft({ ...draft, blobSasUrl: data.value || null })}
                placeholder="https://storageaccount.blob.core.windows.net/container?sv=..."
              />
              <Body1 style={{ fontSize: 11, color: "var(--colorNeutralForeground3)", marginTop: 4 }}>
                SAS URL for your m365reports blob container. Enables Analytics mode with full permissions data.
              </Body1>
            </div>
            <div style={{ borderTop: "1px solid var(--colorNeutralStroke2)", paddingTop: 12 }}>
              <Label style={{ fontWeight: 600, marginBottom: 8, display: "block" }}>Additional Permissions</Label>
              <Checkbox
                checked={draft.enableFilesAccess}
                onChange={(_, data) => setDraft({ ...draft, enableFilesAccess: !!data.checked })}
                label="Enable Files.Read.All (sharing links, drive permissions)"
              />
              <Body1 style={{ fontSize: 11, color: "var(--colorNeutralForeground3)", marginTop: 4, marginLeft: 28 }}>
                Requires the Files.Read.All delegated permission on your app registration.
                Enabling this will trigger a consent popup.
              </Body1>
            </div>
          </div>
        </DialogBody>
        <DialogActions>
          <DialogTrigger>
            <Button appearance="secondary">Cancel</Button>
          </DialogTrigger>
          <Button appearance="primary" onClick={handleSave} disabled={consentPending}>
            {consentPending ? "Requesting consent..." : "Save"}
          </Button>
        </DialogActions>
      </DialogSurface>
    </Dialog>
  );
}
