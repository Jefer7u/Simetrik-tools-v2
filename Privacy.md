# Privacy Policy & Data Handling

## No data is stored, transmitted, or logged

This tool processes JSON files **entirely in memory** within your browser session.

- Uploaded files are **never written to disk**, saved to a database, or sent to any third-party server.
- All processing happens inside a temporary Streamlit session. When you close the tab or the session expires, all data is discarded automatically.
- No analytics, tracking pixels, or external scripts are loaded.
- No account or registration is required.

This is verifiable by inspecting the source code in this repository. There are no `open()` write calls, no database connections, and no outbound HTTP requests to external services.

---

## Disclaimer of liability

This software is provided **"as is"**, without warranty of any kind, express or implied.

The generated Excel and Word documents are produced **algorithmically** from the data contained in the uploaded JSON. The accuracy, completeness, or fitness for any particular purpose of the output is not guaranteed.

**The authors and contributors of this tool:**

- Are not responsible for any decisions made based on the generated documentation.
- Do not guarantee that the output faithfully represents the configuration of any reconciliation platform, including Simetrik or any other third-party product.
- Are not liable for any direct, indirect, incidental, or consequential damages arising from the use or inability to use this tool or its outputs.
- Do not store, access, or process the content of uploaded files beyond the active browser session.

Users are solely responsible for:

- Verifying the accuracy of the generated documentation against the actual source configuration.
- Ensuring that any files uploaded do not contain information that they are not authorized to process through third-party tools.
- Complying with their organization's data handling and confidentiality policies before uploading any JSON exports.

---

## Third-party services

This tool is hosted on **Streamlit Community Cloud**. Streamlit's own privacy policy applies to the hosting infrastructure. This tool does not control or take responsibility for Streamlit's data handling practices. Please review [Streamlit's privacy policy](https://streamlit.io/privacy-policy) before use.

---

## Open source

This tool is released under the [MIT License](./LICENSE). You are free to self-host it in a fully controlled environment if your organization requires it.
