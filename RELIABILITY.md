# O83 Work Alert — Reliability & Recovery

## Canonical URLs

- Primary: https://natt4witsfz.github.io/O83_Work-Alert/
- Stable display entry point: https://natt4witsfz.github.io/O83_Work-Alert/display/
- GitHub repository: https://github.com/natt4witsfz/O83_Work-Alert

Use the primary URL in bookmarks and QR codes. Do not publish an unplanned path such as
`/display` without keeping its `display/index.html` entry point.

## GitHub Pages settings

Verify these settings in **Settings → Pages**:

1. Publishing source is the `main` branch and repository root (`/`), or a deliberately
   configured GitHub Actions deployment. Keep only one publishing model.
2. Enforce HTTPS is enabled.
3. The Pages URL shown by GitHub matches the canonical URL above.

GitHub Pages uses `index.html` as the site entry point. `404.html` is a fallback for
nonexistent paths; it does not turn an unknown path into an HTTP 200 route.

## Main branch protection

Enable protection for `main`:

- Require a pull request before merging.
- Require the **Validate Pages source** status check.
- Require branches to be up to date before merging.
- Block force-push and branch deletion.
- Allow direct emergency changes only to the repository owner.

The GitHub Pages health workflow runs source checks on pushes and live checks on a schedule.

## Recovery procedure

1. Check the Pages health workflow and the live primary URL.
2. If the source check fails, fix the indicated file before publishing.
3. If a recent commit broke the site, use **Revert** on that commit in GitHub.
4. Wait for Pages to republish, then verify:
   - `/` returns HTTP 200.
   - `/index.html` returns HTTP 200.
   - `/display/` returns HTTP 200.
   - The page title contains `ตารางเข้างาน`.
5. If the page opens but staff cannot load, use the **ลองโหลดใหม่** button. If saving
   times out, check Google Sheet before submitting again to avoid duplicate records.

Do not use force-push or `reset --hard` as a rollback method.

## What the health check covers

- Required source files exist.
- `index.html` and `404.html` stay synchronized.
- `display/index.html` remains present.
- Live root, `index.html`, `404.html`, and `display/` respond with the expected app.
- GitHub's default “Page not found” body is not being served.
