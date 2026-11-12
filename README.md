# Calibre Library

Browse and download your Calibre ebooks from personal OneDrive. The browser reads the SQLite catalog after Microsoft sign-in. The public GitHub Pages site serves only the viewer code. Keep your database, covers and ebooks in OneDrive.

## Publish on GitHub Pages

1. Create a GitHub repository for this code. GitHub Free requires a public repository for Pages. In **Settings → Pages**, select **GitHub Actions** as the source.
2. In the [Microsoft Entra admin center](https://entra.microsoft.com), create an app registration supporting **Personal Microsoft accounts**. Add a **Single-page application** redirect URI for the exact Pages URL, including the trailing slash, such as `https://USERNAME.github.io/REPOSITORY/`. For local development, also register `http://localhost:5173/`.
3. Under **API permissions**, add Microsoft Graph delegated **Files.Read**. Microsoft grants read access to the signed-in user's files. The viewer uses that access for the configured Calibre folder.
4. In GitHub **Settings → Secrets and variables → Actions → Variables**, set `MICROSOFT_CLIENT_ID` to the Microsoft **Application (client) ID**. This ID is public. A client secret is unnecessary. Optionally set `CALIBRE_FOLDER` to a OneDrive-relative folder path if your folder differs from `Calibre Library`.
5. Push the code to `main` or rerun the **Publish Calibre viewer** workflow after setting the variables. Open the URL reported by the deployment, choose **Sign in with Microsoft**, and approve read access.
6. Search for a book, open its details, and download an available format. Use **Refresh** after Calibre changes have synced to OneDrive.

## Develop locally

Use Node.js 22.12 or later. Copy `.env.example` to `.env.local`, set the Microsoft application ID and run:

```sh
npm ci
npm test
npm run dev
```

Open the exact localhost URL registered with Microsoft. `npm run build` produces the deployable static files in `dist/`. The viewer reuses Microsoft sign-in across tabs in the same browser profile and reads the catalog again on refresh. OneDrive is accessed with read-only permissions.
