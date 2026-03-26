import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webDarkTheme } from "@fluentui/react-components";
import { AuthProvider } from "./auth/AuthProvider";
import App from "./App";

// TODO Task 11: replace webDarkTheme with dynamic theme derived from AppSettings.theme
createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <FluentProvider theme={webDarkTheme}>
      <AuthProvider>
        <App />
      </AuthProvider>
    </FluentProvider>
  </StrictMode>
);
