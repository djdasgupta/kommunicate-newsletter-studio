# AI-Assisted Mapping Release Checklist

## Automated checks completed

- JavaScript syntax validation for `assets/js/app.js`
- JavaScript syntax validation for both Netlify functions
- HTML duplicate-ID check
- JavaScript-to-HTML element reference check
- Required project-file check
- Netlify function input and missing-key error-path tests
- ZIP integrity test

## One consolidated user acceptance test

1. Open the app and confirm no previous draft loads automatically.
2. Upload `sample_newsletter_input.xlsx` and confirm three regulatory updates appear.
3. Add a PDF, Word file, text file, or pasted source under **Add supporting content**.
4. Confirm the source appears in **Extracted sources** and can be reviewed.
5. Generate AI suggestions. Local files use labelled preview mode; Netlify uses the configured API.
6. Review one suggestion and confirm all field values and source evidence are visible.
7. Map it to a new update and apply it.
8. Review another suggestion against an existing Excel-generated update.
9. Confirm current values are shown before selecting replace or append.
10. Reject one suggestion and regenerate another.
11. Confirm undo restores the last field-level change.
12. Confirm HTML, Word, review-link and email controls still operate.

## Production configuration

- `OPENAI_API_KEY`
- Optional `OPENAI_MODEL`
- `RESEND_API_KEY`
- `NEWSLETTER_FROM`
- `config.js` public URL updated after final Netlify deployment
