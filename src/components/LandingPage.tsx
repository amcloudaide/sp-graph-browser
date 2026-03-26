import { useState } from "react";
import {
  Button,
  Card,
  Input,
  Label,
  Title1,
  Body1,
  Divider,
} from "@fluentui/react-components";
import { useAuth } from "../auth/AuthProvider";

export function LandingPage() {
  const { login } = useAuth();
  const [showCustom, setShowCustom] = useState(false);
  const [customClientId, setCustomClientId] = useState("");

  return (
    <div style={{
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      height: "100vh",
      flexDirection: "column",
      gap: 24,
    }}>
      <Title1>SP Graph Browser</Title1>
      <Body1>Browse your SharePoint Online tenant structure via Microsoft Graph</Body1>

      <Card style={{ padding: 24, maxWidth: 400, width: "100%" }}>
        <Button appearance="primary" size="large" onClick={login} style={{ width: "100%" }}>
          Connect to your tenant
        </Button>
        <Body1 style={{ textAlign: "center", marginTop: 8, color: "#888" }}>
          Signs in with your Microsoft 365 account (delegated, read-only)
        </Body1>
      </Card>

      <Divider style={{ maxWidth: 400, width: "100%" }}>or</Divider>

      <Card style={{ padding: 24, maxWidth: 400, width: "100%" }}>
        {!showCustom ? (
          <Button appearance="subtle" onClick={() => setShowCustom(true)} style={{ width: "100%" }}>
            Use custom app registration
          </Button>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Label htmlFor="clientId">Client ID</Label>
            <Input
              id="clientId"
              value={customClientId}
              onChange={(_, data) => setCustomClientId(data.value)}
              placeholder="00000000-0000-0000-0000-000000000000"
            />
            <Button appearance="primary" onClick={login} disabled={!customClientId}>
              Connect
            </Button>
          </div>
        )}
      </Card>
    </div>
  );
}
