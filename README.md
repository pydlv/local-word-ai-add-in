# Local Writer Word Add-in

Local Writer is a Microsoft Word task pane add-in that uses a local OpenAI-compatible inference server. It has no Azure OpenAI dependency and does not send prompts or document text to a hosted AI service.

## Features

- Insert mode: writes about one paragraph at the cursor.
- Replace mode: rewrites the selected text using the prompt.
- Rolling context window around the cursor or selected text.
- Configurable model, endpoint, context size, and output token limit.
- Local settings persistence through the add-in webview's `localStorage`.
- GitHub Pages workflow for static hosting.

## How It Works

The add-in UI is a static web app. During development it runs at `https://localhost:3000`. In production it can be hosted on GitHub Pages or another HTTPS static host.

For local development, the add-in calls:

```text
/local-ai/v1/...
```

The webpack dev server proxies that to:

```text
http://127.0.0.1:1234/v1/...
```

When hosted from GitHub Pages, the add-in calls your local inference server directly:

```text
http://127.0.0.1:1234/v1/...
```

Your inference server must allow CORS from the hosted add-in origin.

## Local LLM Requirements

The server should expose one or more of these OpenAI-compatible endpoints:

```text
GET  /v1/models
POST /v1/chat/completions
POST /v1/responses
POST /v1/completions
```

The add-in tries generation endpoints in this order:

```text
/v1/chat/completions
/v1/responses
/v1/completions
```

For GitHub Pages hosting, enable CORS on the local LLM server. At minimum:

```text
Access-Control-Allow-Origin: https://YOUR_GITHUB_USER.github.io
Access-Control-Allow-Headers: Content-Type
Access-Control-Allow-Methods: GET, POST, OPTIONS
```

If the project is hosted at a repository path, the origin is still only the host, for example `https://YOUR_GITHUB_USER.github.io`.

## Development

Install dependencies:

```powershell
npm install
```

Start Word with the add-in sideloaded:

```powershell
npm start
```

Use Windows PowerShell or a Windows VS Code terminal for desktop Word sideloading. The default local inference target is:

```text
http://127.0.0.1:1234
```

To use a different local LLM server while developing:

```powershell
$env:LOCAL_AI_TARGET="http://127.0.0.1:YOUR_PORT"
npm start
```

Stop and unload the development add-in:

```powershell
npm stop
```

## Production Build

Build static assets:

```powershell
npm run build
```

The output is written to:

```text
dist/
```

Production builds rewrite manifest URLs from `https://localhost:3000/` to `ADDIN_BASE_URL`.

Example:

```powershell
$env:ADDIN_BASE_URL="https://YOUR_GITHUB_USER.github.io/YOUR_REPO/"
npm run build
```

## Publish to GitHub Pages

This repo includes a GitHub Actions workflow:

```text
.github/workflows/publish-pages.yml
```

To use it:

1. Push this folder as the root of a GitHub repository.
2. In the repository settings, enable Pages with source set to GitHub Actions.
3. Push to `main` or `master`, or run the workflow manually.

The workflow computes the default Pages URL:

```text
https://OWNER.github.io/REPOSITORY/
```

You can override it with a repository variable:

```text
ADDIN_BASE_URL
```

or with the manual workflow input `addin_base_url`.

After publishing, use the generated `manifest.xml` from the Pages site:

```text
https://OWNER.github.io/REPOSITORY/manifest.xml
```

Install/sideload that manifest in Word.

## Using the Add-in

1. Open the Local Writer task pane in Word.
2. Set the endpoint.
   - Local dev default: `/local-ai`
   - GitHub Pages/static hosting default: `http://127.0.0.1:1234`
3. Click `Models`, or type the model id manually.
4. Choose `Insert` or `Replace`.
5. Enter a prompt.
6. Click `Insert Paragraph` or `Replace Selection`.

In Replace mode, select text in the Word document before running the action.

## Offline and Privacy Notes

The add-in does not include the original sample telemetry pixel and does not call Azure OpenAI.

Active network calls are limited to:

- The add-in host, such as GitHub Pages or `https://localhost:3000`.
- The local inference server, usually `http://127.0.0.1:1234`.
- Microsoft's Office JavaScript runtime script, loaded from `https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js`.

If you need a completely air-gapped deployment, you will need to validate how your Office installation provides `office.js` and whether it can be served locally in your environment.

## Important Files

```text
manifest.xml
src/taskpane/components/Home.tsx
src/taskpane/css/Home.css
src/taskpane/taskpane.html
webpack.config.js
.github/workflows/publish-pages.yml
```
