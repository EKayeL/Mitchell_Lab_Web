# Publish the Mitchell Lab web portal: no Terminal required

This is a **deployment package**. Users get a URL, not a ZIP, Python install, or terminal command.

## Publish once, in your browser

1. Extract `Mitchell_Lab_Web_Deployment_No_Terminal.zip` on your computer.
2. In https://github.com/new, create a **private** repository such as `mitchell-lab-web`. Do not upload the old master ZIP, experimental datasets, desktop app, or credential files.
3. In the new repository, select **Add file → Upload files**. Drag in the **contents of the extracted deployment folder**, so `render.yaml`, `Dockerfile`, and `streamlit_app.py` sit at the **repository root**, alongside `analysis_core.py` and the `assets/` folder. If Finder hides `.streamlit`, use Cmd+Shift+. to reveal it, or upload it separately. Commit the changes. Do not upload the ZIP itself as the repository's only file.
4. Open https://dashboard.render.com/ and sign in. Choose **New → Blueprint**. Connect GitHub and select the private repository. Render will read the included `render.yaml`.
5. When Render asks for `APP_PASSWORD`, create a long unique password **in Render**, not in GitHub. Review the compute plan and any charges before approving the deployment. The starter config uses Render's `free` tier for initial smoke testing; 512MB RAM may be inadequate for heavy R analyses or simultaneous users. Upgrade to a plan with adequate RAM if needed.
6. Approve deployment, review build logs, and copy the actual `https://...onrender.com` web-service URL. Share that link and the password privately with authorized lab colleagues after checking successful representative data runs.

## What lab members do (no software installations)

1. Visit the link in Chrome, Firefox, Safari, etc.
2. Enter the lab access password.
3. Choose Convert Txt, QBA, Venn, or Manhattan. Upload an eligible file, set options, run analysis, and download the results ZIP.
4. COSINOR is visible as disabled, and can be added later only after validation.

## Important

- This starter has **not** been deployed or validated end-to-end on real R workbooks. Run known-good, de-identified test files and compare with standalone outputs before relying on it for research.
- It uses a **shared password** and does not provide institutional SSO, lab account management, or an audit trail. Do not use it with identifiable human-participant or sensitive regulated data before institutional IT/security approval.
- Render hosts the service, so neither your Mac nor Seagate drive must remain online.
- Code changes are intentionally not auto-deployed (`autoDeployTrigger: off`) to protect active lab users. Use **Manual Deploy** in Render after a validated code update.
- Please note that web app source code in your private GitHub repo and files on the host remain subject to the chosen vendor's terms and data handling.

For further details consult `README_WEB.md`.
