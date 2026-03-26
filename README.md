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

## License

MIT
