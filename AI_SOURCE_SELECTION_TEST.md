# AI Source Selection Test

1. Load `sample_newsletter_input.xlsx` using the primary Excel control.
2. Add `5169.pdf` as a supporting document.
3. In Extracted sources, confirm:
   - the Excel workbook says **Not selected for AI analysis**;
   - the PDF says **Selected for AI analysis**.
4. Confirm the AI card names only `5169.pdf` as the selected source.
5. Use the instruction: `Create one concise update from this RBI press release. Preserve the consultation deadline and do not invent an effective date.`
6. Select **Generate suggestions**.
7. In local preview mode, verify the suggestion refers to RBI draft directions on loan recovery and recovery agents, identifies India and RBI, captures March 6, 2026 as the feedback deadline, and leaves effective date blank.
8. Verify no suggestion refers to `sample_newsletter_input.xlsx`, the existing newsletter introduction, or unrelated newsletter content.
