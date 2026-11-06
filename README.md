# Gemini Copilot for Microsoft PowerPoint 🪄📊

A modern Microsoft PowerPoint Add-in powered by the **Google Gemini API** from **Google AI Studio**. It functions as an intelligent taskpane presentation copilot—just like the Claude add-in for PowerPoint—enabling you to generate complete slides, rewrite bullet points, craft presenter notes, and interact with your presentation in real-time.

---

## 🌟 Key Features (Claude-like Experience)

| Feature | Description |
| :--- | :--- |
| 💬 **Copilot Chat** | Freeform presentation copilot. Ask for slide ideas, outline topics, brainstorm strategies, and click **"Insert to Slide"** with one tap. |
| 🪄 **Instant Slide Creator** | Input a topic or raw notes, pick a layout (3-Card Summary, Executive Bullets, Problem/Solution, Timeline), and let Gemini construct and insert a formatted slide directly into PowerPoint. |
| ✏️ **Smart Refine / Rewrite** | Highlight any text box on your slide and polish it instantly: *Make Concise*, *Executive Tone*, *Convert to Bullets*, *Make Persuasive*, or *Fix Grammar*. |
| 🎤 **Speaker Notes Generator** | Reads your active slide and automatically writes natural, persuasive talking points into presenter notes. |
| 🔍 **Slide Health Critique** | Evaluates slide density, visual hierarchy, clarity, and recommends actionable fixes. |
| 🔑 **Bring Your Own Key (BYOK)** | Uses your personal Gemini API key from Google AI Studio. Stored strictly in local browser storage—no third-party servers. |
| ⚡ **Multi-Model Support** | Switch between `gemini-2.5-flash` (fast & intelligent), `gemini-2.5-pro` (deep reasoning & analysis), and `gemini-2.0-flash`. |

---

## 🚀 Quick Start (Local Setup)

### 1. Prerequisites
- **Node.js** (v18+)
- **Google AI Studio API Key**: Get a free key at [aistudio.google.com](https://aistudio.google.com/app/apikey)

### 2. Start the Local Server
```bash
# Navigate to add-in directory
cd C:\Users\shiva\.gemini\antigravity\scratch\gemini-powerpoint-addin

# Start the dev server over HTTPS (port 3000)
npm run dev
```

The Vite dev server will start at `https://localhost:3000`.
> **Note on Local HTTPS:** Office Add-ins require HTTPS. Vite automatically configures a development SSL certificate using `@vitejs/plugin-basic-ssl`. When you first visit `https://localhost:3000` in Edge/Chrome, click **Advanced -> Proceed to localhost** to accept the local dev certificate once.

---

## 💻 How to Install & Run (Sideloading in PowerPoint)

You do **NOT** need to publish to the Microsoft Store to use this add-in! You can sideload it in minutes using either of these two methods:

### Method A: PowerPoint on the Web (Easiest - 1 Minute)
1. Open your browser and go to [PowerPoint Online (office.com/launch/powerpoint)](https://www.office.com/launch/powerpoint).
2. Create a new presentation or open an existing one.
3. In the top ribbon, click **Insert** > **Add-ins** (or **Get Add-ins**).
4. In the dialog, click **Manage My Add-ins** (or **My Add-ins** tab at top).
5. Click **Upload My Add-in** (blue link in top-right or bottom).
6. Browse and select your `manifest.xml` file located at:
   `C:\Users\shiva\.gemini\antigravity\scratch\gemini-powerpoint-addin\manifest.xml`
7. Click **Upload**.
8. A new **Gemini AI** button will appear in the **Home** tab! Click **Gemini Copilot** to open the taskpane.

---

### Method B: PowerPoint for Windows Desktop (Shared Folder Catalog)
1. Create a local folder to store your manifest, for example: `C:\OfficeManifests`.
2. Copy `manifest.xml` into `C:\OfficeManifests`.
3. Right-click the `C:\OfficeManifests` folder > **Properties** > **Sharing** tab > **Advanced Sharing...**
4. Check **Share this folder**, click **Permissions**, ensure your user account has **Read** permission, and click **OK**.
5. Note the network path (UNC path), e.g., `\\YOUR-PC-NAME\OfficeManifests`.
6. Open **PowerPoint Desktop**:
   - Go to **File** > **Options** > **Trust Center** > **Trust Center Settings...**
   - Click **Trusted Add-in Catalogs** in the left sidebar.
   - In the **Catalog Url** box, paste your UNC path (e.g. `\\YOUR-PC-NAME\OfficeManifests`), and click **Add Catalog**.
   - Check the **Show in Menu** checkbox next to the added URL, then click **OK**.
7. Restart PowerPoint.
8. Go to **Insert** > **Get Add-ins** > **SHARED FOLDER** tab > Select **Gemini Copilot for PowerPoint** > Click **Add**.

---

## ❓ Frequently Asked Questions

### 1. "Do I need to publish this in the add-in store?"
**NO, you do not need to publish.**

There are three ways to use an Office Add-in:
1. **Personal Sideloading (Free, 0 review, instant)**: Described above. Works for you locally or anyone you give the manifest and URL to.
2. **Centralized Deployment (For your company/team)**: If your company has a Microsoft 365 Business or Enterprise tenant, an IT admin can upload the manifest in the **Microsoft 365 Admin Center** (`admin.microsoft.com` > Settings > Integrated Apps). This automatically deploys the add-in to all employees or specific teams—**no Microsoft Store submission required**.
3. **Microsoft AppSource (Public Store)**: Only required if you want anyone in the world to be able to find and install your add-in from the official Microsoft Office Add-in Store.

---

### 2. "How do I publish this in the Microsoft Add-in Store (AppSource)?"
If you decide you want to publish it publicly so anyone can install it, follow these steps:

#### Step 1: Host Your Web App Publicly
The add-in taskpane is a static web app. You must host it on a public HTTPS domain with 99.9% uptime.
1. Build the production files:
   ```bash
   npm run build
   ```
2. Deploy the `dist/` directory to any free static host:
   - **Vercel**: `vercel deploy --prod`
   - **GitHub Pages**: push `dist/` to a `gh-pages` branch
   - **Cloudflare Pages** or **Netlify**: drag and drop the `dist/` folder
3. Update `manifest.prod.xml`:
   Replace all instances of `https://YOUR_HOSTED_DOMAIN.com` with your production URL (e.g., `https://gemini-powerpoint.vercel.app`).
   Save as `manifest.xml`.

#### Step 2: Validate the Manifest
Ensure your manifest complies with Microsoft's strict schema:
```bash
npm run validate
```
Fix any warnings or errors reported.

#### Step 3: Register in Microsoft Partner Center
1. Sign in to the [Microsoft Partner Center](https://partner.microsoft.com/dashboard).
2. Enroll in the **Commercial Marketplace** program (requires a one-time developer registration fee: ~$19 USD for individuals, ~$99 USD for companies).
3. Complete business/identity verification.

#### Step 4: Create the Add-in Submission
In Partner Center:
1. Go to **Marketplace offers** > **New offer** > **Office Add-in**.
2. Fill in the required metadata:
   - **Offer Alias** & **Title**: "Gemini Copilot for PowerPoint"
   - **Manifest**: Upload your validated `manifest.xml`.
   - **Store Listings**: English (US) description, features, keywords.
   - **Icons**: 256x256 icon and 1280x720 promotional screenshots.
   - **Mandatory Legal URLs**:
     - Privacy Policy URL (explaining how API keys are handled)
     - Terms of Use URL
     - Support contact URL or email.
3. In **Lead management / Notes for certification**, provide test instructions (e.g., mention that the add-in requires a Google AI Studio API key and explain where to enter it).

#### Step 5: Submit for Certification
1. Click **Review and publish**.
2. The Microsoft validation team will test your add-in across PowerPoint on Windows, Mac, and PowerPoint on the Web.
3. Review typically takes **3 to 7 business days**. Once approved, it appears in Microsoft AppSource and within PowerPoint's "Get Add-ins" store.

---

## 🔒 Security & Privacy Architecture
- **Direct API Calls**: All calls to the Gemini API (`https://generativelanguage.googleapis.com`) originate directly from the client's browser/taskpane sandbox.
- **No Intermediary Server**: Your presentation content and API keys are never relayed through any third-party proxy or database.
- **Client-Side Storage**: Your API key is stored securely in `window.localStorage` inside the sandboxed Office WebView2 instance on your machine.
