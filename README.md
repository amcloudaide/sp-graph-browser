# SP Graph Browser

Browse SharePoint Online structure via Microsoft Graph API.

Inspired by [SharePoint Client Browser (SPCB)](https://github.com/nickvdyck/SPCB).

## Features

- Tree browser for SharePoint sites, lists, content types, columns, views, permissions, term store, and recycle bin
- Three view modes: Properties, Raw JSON, and sortable/filterable Table
- IndexedDB caching with configurable TTL for offline-friendly browsing
- Export to JSON, CSV, or styled HTML report
- PWA support (installable, works offline with cached data)
- Multi-tenant authentication with Bring Your Own App (BYOA) support

## Quick Start

```bash
npm install
npm run dev
```

## Required Graph Permissions

The app requires the following Microsoft Graph delegated permissions:

- `Sites.Read.All` -- browse sites, lists, and list items
- `TermStore.Read.All` -- browse managed metadata term store
- `User.Read` -- sign in and read user profile

## App Registration

1. Register a new app in [Entra ID](https://entra.microsoft.com/) > App registrations
2. Set **Supported account types** to "Accounts in any organizational directory" (multi-tenant)
3. Add a **Single-page application** redirect URI: `http://localhost:5173` (dev) and your production URL
4. Under **API permissions**, add the delegated permissions listed above and grant admin consent
5. Copy the **Application (client) ID** and either set it in Settings within the app, or update `src/auth/msalConfig.ts`

## Proxy Setup (Optional)

The "All Sites" node uses `Sites.Read.All` **application** permission to call
`GET /sites?search=*`, which cannot be acquired by a browser SPA (delegated-only).
An Azure Function proxy calls the Graph API with client credentials on behalf of
the SPA.

### Prerequisites

- Azure subscription
- [Azure Functions Core Tools v4](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local): `npm install -g azure-functions-core-tools@4`
- Entra ID app registration with `Sites.Read.All` **application** permission + admin consent

### Deploy

```bash
# Build the proxy
cd proxy && npm install && npm run build

# Create the Function App
az functionapp create \
  --resource-group <rg> \
  --consumption-plan-location <region> \
  --runtime node --runtime-version 20 \
  --functions-version 4 \
  --name <app-name> \
  --storage-account <storage>

# Set environment variables
az functionapp config appsettings set \
  --name <app-name> --resource-group <rg> \
  --settings \
    GRAPH_CLIENT_ID=<client-id> \
    GRAPH_CLIENT_SECRET=<client-secret> \
    GRAPH_TENANT_ID=<tenant-id> \
    ALLOWED_ORIGINS=https://amcloudaide.github.io

# Deploy
cd proxy && func azure functionapp publish <app-name>
```

### Configure in SPA

Open the app, click the **Settings** gear icon, and paste the proxy URL:

```
https://<app-name>.azurewebsites.net
```

### Security

Set `ALLOWED_TENANT_IDS` (comma-separated) to restrict which tenants may call
the proxy. When unset, any authenticated tenant can use it.

## License

MIT
