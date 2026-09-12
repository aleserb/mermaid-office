# Microsoft Marketplace publishing

This file contains listing copy and submission information for Mermaid Office.
Hosting the app on GitHub Pages does not publish a Marketplace listing. Submit
the add-in through Partner Center after completing the owner checklist below.
Nothing here asserts Microsoft certification or completed desktop coverage.

## Listing fields

| Field | Value |
| --- | --- |
| Product name | Mermaid Office |
| Current manifest publisher | aleserb |
| Copyright holder | Aleksandr Serbin |
| Host | Microsoft Word |
| Suggested category | Productivity, if available in the submission form |
| Language | English |
| Price | Free; no Mermaid Office account, subscription, or in-app purchases |
| Software license | MIT; third-party dependencies retain their own licenses |
| Manifest | [`manifest.xml`](manifest.xml) |
| Add-in ID | `9e2ab43d-705f-4f59-b861-a2b35eca6d41` |
| Support URL | https://aleserb.github.io/mermaid-office/support.html |
| Privacy policy URL | https://aleserb.github.io/mermaid-office/privacy.html |
| EULA URL | https://aleserb.github.io/mermaid-office/eula.html |
| License URL | https://github.com/aleserb/mermaid-office/blob/main/LICENSE |
| Third-party notices | https://aleserb.github.io/mermaid-office/third-party-licenses.txt |
| Source repository | https://github.com/aleserb/mermaid-office |

The publisher name in Partner Center must match or closely correspond to the
manifest's `ProviderName`. Confirm the registered publisher identity before
submission; do not substitute an invented company or contact address.

The add-in-only XML manifest has a SupportUrl field. Supply the separate privacy
and EULA URLs in Partner Center; do not add unsupported XML elements for them.
The public pages and stylesheet are copied into `dist` by Vite and deployed with
the app. They are readable anonymously without loading Office.js.

### Short description

Create editable Mermaid diagrams in Word, rendered locally as PNG pictures.

### Long description

The following HTML is ready to adapt for the store's description field:

```html
<p>Create and edit Mermaid diagrams directly in Microsoft Word.</p>
<p>Mermaid Office turns text-based diagrams into PNG pictures that everyone can
view, print, and share without installing the add-in. Select a Mermaid picture
with the pane open to edit its saved source.</p>
<ul>
  <li>Write Mermaid code with syntax highlighting, autocomplete, and error feedback.</li>
  <li>Use live updates for small diagrams and an explicit Update button for large diagrams.</li>
  <li>Customize themes, typography, supported layouts, and PNG quality.</li>
  <li>Keep editable source and settings with the document and original diagram image.</li>
  <li>Render locally in the Office webview without a separate account or rendering server.</li>
</ul>
<p>Free and open source under the MIT License. No Mermaid Office subscription or
additional purchases are required. Microsoft Word and any associated Microsoft
services require their own applicable licenses or accounts.</p>
<p>Designed for modern Word on the web, Windows, and macOS with the required
Office APIs. An internet connection is normally needed to load the add-in.
Source recovery after copying depends on preservation of image metadata.
Large image updates may display Word's own loading dialog. Original diagram
pictures and documents contain editable source, which can be recovered by
recipients. Review confidential source before sharing.</p>
```

Suggested search terms: Mermaid, diagrams, flowchart, sequence diagram, Word.
Confirm the current field limits in Partner Center when entering listing copy.

## Certification notes

No add-in-specific login, test credentials, license key, paid feature, SSO flow,
or external rendering account is required. Reviewers need access to Microsoft
Word and the hosted app. Do not supply personal account credentials.

Required APIs are WordApi 1.4 and ImageCoercion 1.1. The Office-hosted settings
window additionally requires DialogApi 1.2. The manifest requests
ReadWriteDocument so the add-in can read diagram metadata and update pictures,
content controls, and document settings.

Use a new, non-sensitive test document:

1. Open **Insert > Mermaid**. The pane should open without inserting anything.
2. Put the cursor on a blank line and choose **Insert** for the default flowchart.
   A PNG picture should appear in Word.
3. Select the picture and change a label in the editor. A valid small diagram
   should update after a short typing pause while keeping its displayed size.
4. Introduce invalid Mermaid syntax. An error should appear below the editor,
   and the last valid picture should remain unchanged.
5. Restore valid source. Open the gear, change a theme, and choose Apply.
   Reopen settings and verify the saved value. Cancel should discard form edits.
6. Create or edit a diagram with at least 50 nonblank source lines. Update should
   replace live writes. Edit it, wait, and confirm Word changes only after Update.
   The editor should retain edits typed during a write for the next Update.
7. Change document selection while settings are open. Apply should update the
   original diagram; Cancel should resume selection without applying changes.
8. Save, close, and reopen the document. Select the diagram and confirm its source
   and settings can be edited. Resize a picture, then update and verify its width.
9. Copy the original diagram within and between Word documents using supported
   picture paste paths. Confirm source recovery where metadata is preserved.
   Also check Undo/Redo, deletion, and recipients viewing without the add-in.
10. Check keyboard navigation, accessible control names, error announcements,
    narrow task panes, and settings window dismissal.

Repeat the relevant cases on Word web, Windows, and macOS; record exact versions
and browsers. This checklist is a test plan, not evidence that these checks
have already passed on every platform.

## Owner checklist before submission

- [ ] Enroll in the relevant Partner Center program and complete publisher verification.
- [ ] Confirm `aleserb` matches the registered publisher, or update the manifest
  and public pages consistently.
- [ ] Review and approve the privacy policy, EULA, and MIT licensing decision;
  obtain legal advice appropriate to your jurisdiction and distribution plans.
- [ ] Provide genuine Partner Center contact details and confirm the support
  channel is monitored. Consider a private contact channel for sensitive requests;
  the current GitHub tracker is public. Do not put private account details in Git.
- [ ] Supply at least one current screenshot using a synthetic document with
  no personal account details. Prepare store logo assets at the dimensions
  required by the current submission form; the ribbon's 16/32/80px icons alone
  are not a complete store asset package.
- [ ] Review dependency licenses and ensure required third-party notices are
  included in the distributed build. The project MIT License does not replace them.
- [ ] Complete the cross-platform and accessibility coverage above and retain results.
- [ ] Run the existing test, lint, and build commands and Microsoft's official
  manifest/store validation. The repository's manifest test is not a substitute
  for Microsoft's validation service.
- [ ] Confirm all policy/support URLs and manifest assets are live over HTTPS,
  accessible without sign-in, and do not redirect to the task pane.
- [ ] Enter listing details, legal URLs, screenshots, and certification notes
  in Partner Center. Disclose Word requirements and select no additional
  Mermaid Office purchases; verify all claims against the released build.
- [ ] Submit for review, address certification feedback, and add the approved
  Marketplace installation link to the support page after approval.

Do not claim offline installation, native image right-click editing, vector PNG
output, guaranteed lossless source recovery after metadata stripping, or Microsoft
endorsement.

## Release and policy maintenance

Keep the add-in ID unchanged for updates. Increment the manifest version when
appropriate, deploy the app and public pages together, and verify live URLs.
Changes to data handling must be reflected in the privacy policy with a new
effective date. Update the EULA if relevant terms change while respecting MIT
rights already granted.

The public policy pages are the canonical documents; link to them rather than
maintaining duplicate Markdown versions. Root `LICENSE` is the canonical MIT text.
Vite generates `dist/third-party-licenses.txt` from bundled dependencies during
each production build. Review that output for missing notices or additional
license obligations; a generated notice file is not a legal compliance audit.

## Official references

Requirements can change; consult these before each submission:

- [Publish your Office Add-in to Microsoft Marketplace](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource)
- [Submission checklist, including support, privacy, and EULA URLs](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist)
- [Marketplace certification policies](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies)
- [Office manifest validation](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest)
- [Store screenshots and imagery](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/craft-effective-appsource-store-images)
