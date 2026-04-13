# M365 Attachment Reader MCP Local

A local `stdio` MCP server for Claude Desktop that reads Outlook emails and deeply parses email attachments through the Microsoft Graph API.

> Status: Functional for personal single-user local use with Claude Desktop.

---

## Why This Exists

Claude's built-in Microsoft 365 connector can list emails, read message bodies, and check calendars. But it cannot read the actual content inside email attachments.

That means when you say "What does the PDF in my latest email say?", Claude can see the attachment metadata, but not the text, tables, images, or nested documents inside it.

This project fills that gap — running entirely on your local machine over `stdio`, with no public endpoints or tunnels required.

---

## What It Does

This server runs as a local MCP process started by Claude Desktop. It:

1. Authenticates with Microsoft 365 via device code flow
2. Lists Outlook emails and their attachments through Microsoft Graph
3. Downloads and parses attachment contents locally
4. Returns structured text and image blocks directly to Claude Desktop

### Supported Formats

| Format | What Gets Extracted |
|---|---|
| PDF | Full text content |
| Scanned PDF | OCR text, plus optional rendered page images |
| DOCX | Text and embedded images |
| DOC | Text content |
| PPTX / PPTM / PPSX / POTX | Slide text, notes, and embedded images |
| PPT | Best-effort legacy text extraction |
| XLSX / XLS / CSV | All sheets converted to CSV |
| JPG / JPEG / PNG / GIF / WEBP / BMP / TIFF | Returned as MCP image blocks for visual analysis |
| ZIP / RAR / 7Z | Archive contents recursively parsed file by file |
| MSG | Subject, sender, body, and embedded attachments |
| TXT / MD / JSON / XML / HTML | Raw text |
| Outlook `itemAttachment` | Text content |

### MCP Tools

| Tool | Description |
|---|---|
| `health_check` | Check if the server is alive |
| `begin_auth` | Start device code login flow |
| `auth_status` | Check authentication status |
| `list_recent_messages` | List recent Outlook emails |
| `list_email_attachments` | List attachments for a specific email |
| `read_email_attachment` | Download, parse, and return attachment content |

---

## Real-World Use Cases

### Retail / Sales Operations
> "Pull the last 5 Daily Dashboard emails, read the Excel attachments, and analyze the sales trend across all store locations over the past week."

### Finance / Accounting
> "Find the latest email from our vendor with 'Invoice' in the subject, read the PDF attachment, and extract the total amount, due date, and line items."

### Legal / Contract Review
> "Open the most recent email from legal@partner.com, read the Word or PowerPoint attachment, and summarize the key terms."

### HR / Recruiting
> "Find emails from recruiting@company.com with attachments, read each resume PDF, and create a comparison table of candidates."

---

## Prerequisites

- Windows 10/11, macOS, or Linux
- [Node.js 20 or later](https://nodejs.org)
- Claude Desktop
- A Microsoft 365 / Outlook account
- A Microsoft Entra app registration (see Step 1 below)

---

## Setup

### 1. Create a Microsoft Entra App Registration

Go to [Microsoft Entra admin center](https://entra.microsoft.com) → **App registrations** → **New registration**.

- **Name:** anything you like, e.g. `m365-mcp-local`
- **Supported account types:** Accounts in any organizational directory and personal Microsoft accounts

Then:

1. Copy the **Application (client) ID** from the Overview page
2. Go to **Authentication** → enable **Allow public client flows** → **Save**
3. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → add `User.Read` and `Mail.Read` → **Grant admin consent**

### 2. Clone and Install

```bash
git clone https://github.com/Zacccck/Claude-MCP-Read-Email-Attachments.git
cd Claude-MCP-Read-Email-Attachments
npm install
```

### 3. Configure Environment Variables

Copy the example file:

```bash
cp .env.example .env
```

Edit `.env` and fill in your client ID:

```env
M365_CLIENT_ID=your-application-client-id-here
M365_TENANT_ID=common
M365_AUTO_OPEN_BROWSER=true
```

Variable reference:

| Variable | Required | Description |
|---|---|---|
| `M365_CLIENT_ID` | ✅ Yes | Your Entra app's Application (client) ID |
| `M365_TENANT_ID` | No | Default `common` works for most accounts |
| `M365_AUTO_OPEN_BROWSER` | No | Set `true` to auto-open Microsoft login page |
| `M365_MCP_DATA_DIR` | No | Custom path for auth cache; auto-detected if omitted |

### 4. Find Your Node.js Path

You'll need the full path to `node.exe` (Windows) or `node` (macOS/Linux) in the next step.

```powershell
# Windows
where.exe node

# macOS / Linux
which node
```

Example output: `C:\Program Files\nodejs\node.exe`

### 5. Open Claude Desktop's Config File

Locate and open the config file for your platform:

| Platform | Path |
|---|---|
| Windows (standard) | `%APPDATA%\Claude\claude_desktop_config.json` |
| Windows (Store) | `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json` |
| macOS | `~/Library/Application Support/Claude/claude_desktop_config.json` |

If the file does not exist yet, create it.

### 6. Add the Server to Claude Desktop

Add the following entry to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "m365-attachment-reader-local": {
      "command": "C:\\Program Files\\nodejs\\node.exe",
      "args": [
        "C:\\path\\to\\Claude-MCP-Read-Email-Attachments\\server.mjs"
      ],
      "env": {
        "M365_CLIENT_ID": "your-client-id",
        "M365_TENANT_ID": "common",
        "M365_AUTO_OPEN_BROWSER": "true"
      }
    }
  }
}
```

> **Tips:**
> - Use the full absolute path from Step 4 for `command`.
> - Replace `args[0]` with the actual path to `server.mjs` on your machine.
> - If you already have other MCP servers in the config, merge this entry into the existing `mcpServers` object — do not overwrite the whole file.

### 7. Restart Claude Desktop

Completely quit Claude Desktop and reopen it. Claude Desktop starts the MCP server automatically — you do not need to run `node server.mjs` manually.

### 8. Authenticate with Microsoft 365

In Claude Desktop, type:

```text
Please call begin_auth
```

A browser window will open (or you'll receive a login URL + device code). Complete the Microsoft login flow, then verify:

```text
Please call auth_status
```

You should see your Microsoft account listed as authenticated.

### 9. Verify It Works

Run a quick health check:

```text
Please call health_check
```

Then try a real request:

```text
Show me my recent Outlook emails with attachments
```

```text
Summarize the contents of the attachments from the latest email
```

---

## Suggested Claude Prompts

```text
Please call begin_auth
```

```text
Please call auth_status
```

```text
Show me my recent Outlook emails with attachments
```

```text
Summarize the contents of the attachments from the email
```

```text
Find the latest invoice email and extract the total amount, due date, and line items from the PDF attachment
```

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Claude cannot find MCP tools | Restart Claude Desktop completely. Check that `command` and `args` paths in the config are correct and absolute. |
| `invalid_grant` error | Re-check your Entra app: supported account types, public client flow enabled, `User.Read` and `Mail.Read` permissions granted. |
| Device code not showing | Make sure `begin_auth` was called successfully. Do not manually enter a code. |
| Want to switch Microsoft accounts | Restart Claude Desktop and call `begin_auth` again in a private browser window. |
| Debug log location | `<M365_MCP_DATA_DIR>\debug.log` — defaults to a subdirectory auto-created next to `server.mjs`. |

---

## Manual Development Run

For debugging outside Claude Desktop, start the server manually:

```powershell
cd Claude-MCP-Read-Email-Attachments
node .\server.mjs
```

> **Note:** Do not type into that terminal. It is a `stdio` MCP process and expects an MCP client on standard input/output.

---

## Docker

A `Dockerfile` is included for containerized testing:

```powershell
docker build -t m365-attachment-reader-mcp-local .
docker run --rm -i `
  -e M365_CLIENT_ID=your-client-id `
  -e M365_TENANT_ID=common `
  -e M365_AUTO_OPEN_BROWSER=false `
  m365-attachment-reader-mcp-local
```

> The container still runs as a `stdio` server. For everyday Claude Desktop use, the direct `node` approach in Step 6 is simpler.

---

## Project Structure

```text
Claude-MCP-Read-Email-Attachments/
├── server.mjs
├── package.json
├── manifest.json
├── server.json
├── glama.json
├── Dockerfile
├── .env.example
├── .gitignore
├── LICENSE
└── README.md
```

---

## Limitations

- **Single-user only** — one server instance supports one Microsoft account at a time
- **Auth state is in-memory** — restarting the server requires re-authenticating
- **You must create your own Entra app** and supply your own client ID
- **Very large images** may be downscaled or skipped to stay within Claude Desktop payload limits
- **Legacy `.xls`** parsing is best-effort and less reliable than `.xlsx`
- **Not suitable** for public or multi-user hosting

---

## License

MIT
