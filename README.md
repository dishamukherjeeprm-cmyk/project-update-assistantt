# Project Update Assistant

Project Update Assistant is a Streamlit prototype for internal consulting teams. Upload the current requirements document and the latest meeting notes, inspect the extracted content, compare the documents, review evidence-backed findings, and export only the findings a human has approved.

## Run locally

The workspace uses Python 3.11 and keeps the Python dependencies in `pyproject.toml`.

```bash
streamlit run main.py --server.port 5000
```

The Replit workflow uses the same Streamlit app on the workspace preview port.

## OpenAI API key

The app reads `OPENAI_API_KEY` from the environment. Add it through the Replit Secrets panel; never paste it into `main.py`, `README.md`, or an uploaded document. Optionally set `OPENAI_MODEL` to choose another compatible chat-completions model.

When `OPENAI_API_KEY` is not available, the app uses a transparent, evidence-based local comparison so the workflow can still be demonstrated. It does not pretend that this fallback is a model-generated result.

## Supported files

- Requirements: `.xlsx`, `.docx`, `.pdf`, `.txt`
- Meeting notes/transcripts: `.pdf`, `.docx`, `.txt`

Excel sheets and Word tables are preserved as extracted sections. PDF text is extracted with `pypdf`.

## Data-handling limitations

- Only upload anonymised or approved documents.
- Files are read into memory for the active Streamlit session and are not permanently stored by the application.
- The OpenAI comparison sends extracted document text to the configured OpenAI API when the secret is available. Follow your organisation's data-processing approval requirements.
- The app does not update source documents, send Slack messages, create Asana tasks, or make client-facing changes automatically.
- Exports contain only findings that a reviewer explicitly approves in the interface.

## Human review

This is a prototype requiring human review. Meeting language is treated as evidence, not as an automatic client decision. Owners, deadlines, requirements, and technical solutions are never invented; missing information is labeled for validation.