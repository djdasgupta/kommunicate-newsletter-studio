# Kommunicate Newsletter Studio - Mapping Core

This release adds manual document-to-placeholder mapping, protected overrides, document-first newsletter creation, source provenance, field-level history, undo and local draft autosave.

# Kommunicate Newsletter Studio

A browser-based internal tool for turning structured Excel or Word content into a branded Lexplosion regulatory newsletter.

## What the app supports

- Upload `.xlsx` or `.docx` source files.
- Edit newsletter content in the browser.
- Generate an interactive HTML newsletter.
- Change primary and accent colours live.
- Download the final newsletter as HTML.
- Download an editable Word `.docx` review copy.
- Send the rendered newsletter directly to one or more reviewers.
- Create a shareable interactive review link after deployment.

## Repository structure

```text
kommunicate-newsletter-studio/
├── index.html
├── config.js
├── config.example.js
├── sample_newsletter_input.xlsx
├── assets/
│   ├── css/styles.css
│   └── js/app.js
├── netlify/
│   └── functions/send-newsletter.mjs
├── netlify.toml
└── README.md
```

The UI, styling, and application logic are separated so team members can maintain the app through normal GitHub pull requests.

## Recommended deployment model

Use **GitHub as the source repository** and connect the repository to **Netlify** for continuous deployment. Every merge to the production branch will automatically publish the updated app. Netlify also runs the serverless email function used by the review-email button.

GitHub Pages can host the visual application, but it cannot run the included serverless email function. For one-click email delivery, deploy through Netlify or replace the email function with an approved internal API.

## First-time setup

### 1. Create the GitHub repository

```bash
git init
git add .
git commit -m "Initial Kommunicate Newsletter Studio"
git branch -M main
git remote add origin <YOUR_GITHUB_REPOSITORY_URL>
git push -u origin main
```

Grant the appropriate team members repository access. Future changes should be made in branches and merged through pull requests.

### 2. Deploy from GitHub to Netlify

1. In Netlify, choose **Add new site** and import an existing project.
2. Select GitHub and choose this repository.
3. No build command is required.
4. Publish directory: `.`
5. Functions directory: `netlify/functions` (also defined in `netlify.toml`).
6. Deploy the site.

### 3. Configure direct email delivery

The app sends review emails through the included Netlify Function and the Resend email API. Create a Resend account, verify a sending domain, and create an API key.

Add these environment variables in Netlify under the site configuration. Ensure they are available to Functions:

```text
RESEND_API_KEY=re_xxxxxxxxxxxxxxxxx
NEWSLETTER_FROM=Kommunicate Reviews <newsletter@updates.yourdomain.com>
```

`NEWSLETTER_FROM` must use a domain verified in Resend. A reply-capable address is recommended.

Redeploy the site after adding or changing environment variables.

### 4. Configure the public app URL

After Netlify assigns the production URL, edit `config.js`:

```javascript
window.APP_CONFIG = {
  emailEndpoint: '/.netlify/functions/send-newsletter',
  publicAppUrl: 'https://your-site-name.netlify.app'
};
```

Commit and push that change. The app also detects its current hosted URL, so `publicAppUrl` may remain blank unless you use a custom domain or a different canonical URL.

## Using the app

1. Upload the structured Excel or Word file, or choose **Load demo**.
2. Review and edit the extracted content.
3. Adjust the colour controls where required.
4. Select **Generate**.
5. Choose one of the output actions:
   - **Download HTML** for the final browser/email-ready file.
   - **Download Word** for an editable `.docx` review copy.
   - **Copy review link** for the hosted interactive version.
   - **Email newsletter for review** to send the rendered newsletter to one or more recipients.

## Review recipients

The old “staging email” field has been replaced with **Review recipients**. Enter one or more addresses separated by commas, semicolons, or new lines. The app sends the same generated newsletter to all recipients in one request.

The old “hosted app URL” field has been removed from the user interface. The URL is now controlled centrally through `config.js` or detected automatically from the deployed site. This prevents individual users from having to understand or maintain deployment settings.

## Local development

For the static UI only, start any local web server:

```bash
python3 -m http.server 8080
```

Open `http://localhost:8080`.

To test the email function locally, install the Netlify CLI, link the local folder to the Netlify site, and run:

```bash
netlify dev
```

The local environment must have `RESEND_API_KEY` and `NEWSLETTER_FROM` available through Netlify CLI environment configuration.

## Maintaining the application

- UI markup: `index.html`
- Visual design: `assets/css/styles.css`
- Parsing, generation, Word export, and interactions: `assets/js/app.js`
- Email delivery: `netlify/functions/send-newsletter.mjs`
- Runtime configuration: `config.js`

Do not place API keys in `config.js`, JavaScript files, or GitHub. Secrets must remain in Netlify environment variables.

## Word export notes

The Word download contains the issue metadata, lead story, regulatory updates, full analysis fields where available, expert insights, product content, and Lexplosion footer text. It is intentionally structured for manual editing rather than attempting to reproduce every email-layout detail.

## Troubleshooting email delivery

- **Email service is not configured:** the Netlify environment variables are missing or not available to Functions.
- **Provider rejected the request:** verify the Resend API key and sender domain.
- **Works locally but not in production:** redeploy after changing environment variables.
- **Review link missing:** set `publicAppUrl` in `config.js` or open the deployed app over HTTPS rather than directly from a local file.

## Review workflow and Word export

### Branded Word download

The **Download Word** button now converts the rendered newsletter HTML into a Word document rather than rebuilding the content as plain paragraphs. This retains the newsletter's primary and accent colours, masthead, typography hierarchy, card layout and footer much more closely.

The result is intended for editing in Microsoft Word desktop. Minor spacing differences may still occur because HTML and Word use different layout engines.

### Create review link

After generating a newsletter, click **Create review link**. A dialog displays the complete interactive link and provides:

- **Copy link** to place the URL on the clipboard.
- **Open link** to test the review experience in a new browser tab.
- A visible URL field for manual copying when browser clipboard permission is blocked.

The link contains the current newsletter content and selected colours. Anyone receiving it can open the same generated edition in the hosted app.

### Email for review

Click **Email for review** after generating a newsletter. A dialog asks for:

- One or more recipient email addresses.
- Email subject.
- Review message.

Multiple addresses can be separated by commas, semicolons or new lines. The app sends the rendered newsletter and interactive review link through the Netlify serverless function. `RESEND_API_KEY` and `NEWSLETTER_FROM` must be configured in Netlify environment variables.

## Flexible design editor (Phase 1)

This version adds a governed visual editor while retaining email-safe HTML output.

Users can now:
- change heading and body fonts;
- change the lead-story headline colour and size;
- change body text size and newsletter content width;
- upload or replace the Lexplosion masthead logo;
- upload replacement logos for Komrisk AI, Komtrakt, Komtrol and Komtrol Plus;
- restore default logos;
- save a preferred design in the current browser;
- reset the design to the governed default;
- carry the selected design into HTML, Word and review-link outputs.

Logo files are stored in the browser as part of the active design state. Use PNG, JPG, WebP or SVG files under 2 MB. For a later multi-user production version, logos and templates should be stored in managed cloud storage rather than browser local storage.

## Planned next phase

The next increment will introduce unstructured ingestion for PDF, Word and pasted text, followed by a secure Netlify Function that sends extracted content and a user-written instruction to an LLM and returns schema-validated newsletter content.

## Phase 2: multi-format source ingestion

This version adds a new source-input layer before the future LLM workflow.

### Input modes

1. **Structured file**: existing `.xlsx` and structured `.docx` mapping.
2. **Documents**: upload one or more `.pdf`, `.docx`, or `.txt` files.
3. **Paste text**: paste unstructured source content directly.

### What the document mode does

- Extracts selectable text locally in the browser.
- Lists each source with file type, size, page/section count, and word count.
- Lets users review and edit extracted text.
- Supports include/exclude, remove, and reorder controls.
- Combines all included sources into a single reviewable text package.
- Downloads the combined source package as a `.txt` file.

### Current limitation

This phase does not send source content to an LLM and does not automatically populate the newsletter. Scanned or image-only PDFs may produce no selectable text and will require OCR in a later phase.

### Recommended test

1. Open `index.html` from a hosted environment or local static server.
2. Select **Documents**.
3. Upload a normal text-based PDF, DOCX, and TXT file.
4. Review each extracted source.
5. Exclude and reorder sources.
6. Select **Combine sources** and verify the combined text.
7. Select **Paste text**, add sample content, and confirm it appears in the source list.

## Clean primary-first workflow

This build makes the standard Excel template the default entry point. Supporting PDF, Word, Excel and text sources are available under the collapsed **Add supporting content** section. Advanced design, editing and draft-recovery controls are also collapsed to reduce sidebar congestion.

## AI source selection safeguard (v1.1)

The standard Excel workbook is retained for extraction review but is excluded from AI analysis by default. In the right-hand Extracted sources panel, select **Use for AI** only on the supporting PDF, Word, Excel or text sources that the model should analyse. Changing the AI source selection clears unapplied suggestions to prevent stale results.

## Release Candidate 1

This package is the recommended internal deployment candidate. It adds clearer manual-versus-AI workflow guidance, portable project files, source-provenance reporting and an in-app release-readiness check.

Before deployment, read `DEPLOYMENT_AND_TEAM_HANDOFF.md` and run:

```bash
node scripts/release-check.mjs
```

The release check validates required files, duplicate element IDs, JavaScript-to-HTML element references, missing-key behavior and a mocked structured AI response.
