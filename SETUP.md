# How to publish these docs to your repo

## 1. Copy the files into your repo

Drop these into the root of `inforcer_API_toolkit`:

```
inforcer_API_toolkit/
├── docs/                          ← all the .md files
│   ├── index.md
│   ├── getting-started.md
│   ├── cli.md
│   ├── troubleshooting.md
│   └── workflow/
│       ├── index.md
│       ├── phase-1-export.md
│       ├── phase-2-review.md
│       └── phase-3-build.md
├── mkdocs.yml                     ← config
└── .github/
    └── workflows/
        └── deploy-docs.yml        ← auto-deploy action
```

Your existing `README.md` stays where it is — GitHub still shows it on the repo home page.

## 2. Commit and push to `main`

```bash
git add docs mkdocs.yml .github/workflows/deploy-docs.yml
git commit -m "Add MkDocs documentation site"
git push
```

## 3. Enable GitHub Pages

In your repo on github.com:

1. **Settings → Pages**
2. Under **Source**, choose **GitHub Actions**
3. Save

## 4. Watch it deploy

- Go to the **Actions** tab — the "Deploy docs" workflow should be running.
- When it finishes (~1 minute), your site is live at:
  **https://jamesinforcer.github.io/inforcer_API_toolkit/**

## 5. Updating the docs

Edit any `.md` file in `docs/`, push to `main`, and the action redeploys automatically. No manual rebuild needed.

## Optional — preview locally before pushing

```bash
pip install mkdocs-material
mkdocs serve
```

Then open `http://127.0.0.1:8000` in your browser. It live-reloads as you edit.
