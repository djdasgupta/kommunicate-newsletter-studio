# Deployment and team handoff

## Release scope

This release candidate combines the approved Excel-first workflow, optional supporting-source ingestion, manual document mapping, AI-assisted suggestions, source-level safeguards, design controls, exports, review links and email delivery.

It also adds project export/import, a field-level source report and a release-readiness check. These features allow a colleague to continue work without sharing browser storage.

## Netlify environment variables

Configure these under Site configuration > Environment variables and make them available to Functions:

- `OPENAI_API_KEY`: server-side OpenAI API key used by `analyze-sources.mjs`.
- `OPENAI_MODEL`: optional model override. If omitted, the function uses its configured default.
- `RESEND_API_KEY`: server-side Resend API key.
- `NEWSLETTER_FROM`: verified sender, for example `Kommunicate <newsletter@yourdomain.com>`.

Never add these values to GitHub, `config.js` or browser JavaScript.

## GitHub and Netlify release

1. Keep the current production commit available for rollback.
2. Upload the complete contents of this folder to the existing GitHub repository, preserving `assets/` and `netlify/functions/`.
3. Commit to a release branch rather than `main` for the first deployment.
4. In Netlify, create a deploy preview from that branch.
5. Confirm that the Functions directory is `netlify/functions` and the publish directory is `.`.
6. Add or verify the four environment variables above.
7. Run the acceptance test below on the deploy preview.
8. Merge the release branch into `main` only after acceptance.

## Team handoff workflow

A user can select **Project and release tools > Export project file**. The JSON project contains newsletter content, extracted text, mappings, source provenance, AI review state and design settings.

A colleague can open the hosted app and select **Import project file**. This replaces their current working session after confirmation.

The project file may contain extracted regulatory source text. Store and share it only through approved internal channels.

## Minimum acceptance test

1. Upload `sample_newsletter_input.xlsx` and confirm the newsletter is generated.
2. Upload a PDF as supporting content and confirm text extraction.
3. Map one highlighted passage manually into a new update.
4. Select only that PDF for AI and generate one suggestion.
5. Confirm the suggestion cites the selected source and does not use the Excel workbook.
6. Apply the suggestion to a new update and test undo on a separate manual mapping.
7. Change fonts, colours and width, then refresh the preview.
8. Export HTML and Word.
9. Export the project JSON, refresh the browser, and import the project.
10. Download the source report and confirm field origins are listed.
11. Run **Release readiness check** and resolve any failures.
12. Create a review link and send a test email from the Netlify preview.

## Rollback

If production validation fails, restore the previous GitHub commit or use Netlify's published-deploy rollback. Do not delete the previous production deployment until the team has completed acceptance.
