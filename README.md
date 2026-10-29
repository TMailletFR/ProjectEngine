# ProjectEngine GitHub Pages

Static landing page for:

https://tmailletfr.github.io/ProjectEngine/

## Deployment

Recommended deployment method: a dedicated `gh-pages` branch.

1. Create a new orphan branch named `gh-pages`.
2. Copy the contents of this folder to the root of that branch.
3. Commit and push the branch.
4. In the repository, open `Settings → Pages`.
5. Under `Build and deployment`, select `Deploy from a branch`.
6. Select branch `gh-pages` and folder `/ (root)`.
7. Save.

The site is a plain static site. No build process, package manager or dependency is required.

## Included SEO files

- canonical URL
- Open Graph and Twitter metadata
- SoftwareApplication structured data
- robots.txt
- sitemap.xml
- custom 404 page
- favicon
- `.nojekyll`


## v2 corrections

- Hero slogan locked to two lines on desktop.
- Workflow fit panel vertically centered.
- Analytics and S-Curve cards retain their natural image ratios.
- Download CTAs use `download.html`, which resolves the `.xlsm` asset from the latest GitHub release dynamically.


## v3 correction

The stylesheet URL now carries a version query (`styles.css?v=3`) so GitHub Pages and browsers do not keep serving the older cached CSS. Explicit hero colours and stronger desktop/landscape layout rules were also added.

## Privacy-friendly audience measurement

The site uses Cloudflare Web Analytics for aggregate audience and performance measurement.

- no analytics cookies;
- no local storage or persistent visitor identifier;
- no fingerprinting or cross-site tracking;
- no advertising or profiling;
- no consent banner;
- a dedicated `privacy.html` page explains the measurement.

The Cloudflare beacon is installed on `index.html`, `download.html`, `404.html` and `privacy.html`.



## Release archive (added October 2026)

- `releases/index.html` is the indexable release archive.
- `releases/v*-*.html` are static, permanently linked summaries of significant releases.
- GitHub remains the authoritative source for releases, full technical changelogs and downloads.
- Release dates in structured data must only be added when their exact dates are verified.
- When publishing a new version, add a curated HTML page, link it from `releases/index.html` and the homepage, and update `sitemap.xml`. Do not change old canonical URLs.
- Publish a fresh `lastmod` only for genuinely updated pages.
- Do not submit `download.html` to the sitemap: it is deliberately `noindex` and redirects to GitHub.
- Pages are pure static HTML/CSS; no site generator or additional runtime dependency is needed.


## v1.5 website update (October 29, 2026)

- Added `releases/v1-5.html`, a curated public summary of the Financial Project Controls & Multi-Project Management release.
- Updated `releases/index.html` with v1.5 as latest and added the page to `CollectionPage.hasPart`.
- Updated the homepage's SEO metadata, feature summary and latest release links to reflect Finance and multi-project cost reporting.
- Updated `sitemap.xml` with the v1.5 canonical URL and change dates for the pages updated on October 29.
- The existing `download.html` resolves the latest published `.xlsm` asset through GitHub's API and therefore requires no hardcoded version change. Publish the GitHub release with the `v1.5` tag and a prepared `.xlsm` asset before sending users to the updated site.
- The stylesheet, analytics configuration, privacy page, robots and legacy release pages were not changed. This patch introduces no new tracking.
