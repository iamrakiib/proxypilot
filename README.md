# AdsPower + IPQS Browser Automation

A Node.js browser automation project built with Playwright, with optional AdsPower profile management and IPQualityScore (IPQS) proxy checks.

## Features

- Run browser sessions in desktop or mobile mode.
- Rotate configured proxies and cycle through URL groups.
- Open group URLs in parallel, with an optional tab limit.
- Create, start, stop, and delete temporary AdsPower profiles.
- Optionally check proxy exit IPs with IPQS and rotate API keys when limits are reached.
- Load proxies and user agents from an Excel workbook or plain text files.
- Support direct site tests and a mock search page for controlled testing.

Use automation only on sites and accounts you are authorized to test. The default flow can click matching calls to action and follow links; review the target pages before running it.

## Requirements

- Node.js 22 or newer and npm.
- A working proxy; the application stops if no proxies are configured.
- AdsPower with its local API enabled when using AdsPower mode.
- An IPQS API key if you want proxy scoring.

## Setup

From the project folder:

~~~powershell
npm ci
npx playwright install chromium
Copy-Item ads-config.example.json ads-config.json
Copy-Item groups.example.json groups.json
Copy-Item proxies.example.txt proxies.txt
~~~

The Chromium download is needed for standalone Playwright mode. AdsPower mode connects to the browser managed by AdsPower.

Edit the new local files:

1. In groups.json, replace the example URL with a site you are authorized to test.
2. In proxies.txt, replace the placeholder with your real proxy details, one proxy per line.
3. In ads-config.json, set your local AdsPower API address and optional keys.

Local configuration files are deliberately excluded from Git. Copy commands overwrite existing files, so use them only on a fresh checkout or after backing up your local settings.

## Run

~~~powershell
# Standalone Playwright
npm start

# AdsPower profile lifecycle
node clickbot.js --adspower

# Direct site test
$env:SITE_TEST_URL = "https://example.com/"
node clickbot.js --site-test

# Mock search page that links to the configured test site
node clickbot.js --mock-search-test
~~~

At startup, enter the number of profiles, choose mobile or desktop, and select whether to enable IPQS checks when keys are available. Start with one profile to inspect the behavior.

AdsPower mode deletes the temporary profiles it creates during cleanup. Cleanup after a forced process termination may require manual attention.

## Configuration

| File | Purpose |
| --- | --- |
| ads-config.json | Local API address, credentials, headful setting, and IPQS thresholds. |
| groups.json | Named URL arrays, cycled one group per profile. |
| proxies.txt | Proxy list, one entry per line. |
| useragents.txt | Optional mobile user agents, one per line. |
| desktop_useragents.txt | Optional desktop user agents, one per line. |
| profiles.xlsx | Optional workbook with column A in Proxies, UserAgents, and Desktop UA sheets. Nonempty workbook pools take priority over their text-file equivalents. |

Supported proxy formats include host:port, user:pass@host:port, scheme://user:pass@host:port, and user:pass:host:port. IPQS exit-IP discovery currently does not support SOCKS proxies.

| Environment variable | Purpose |
| --- | --- |
| ADSP_USE | Set to 1 to enable AdsPower. |
| ADSP_BASEURL | Override the local AdsPower API address. |
| ADSP_API_KEY | Override the AdsPower API key. |
| ADSP_HEADFUL | Set to 1 to show the AdsPower browser. |
| ADSP_OS | Desktop fingerprint OS: windows or linux. |
| ADSP_GROUP_ID | Use a specific AdsPower group. |
| ADSP_GROUP_NAME | Group name to find or create; default AutoGroup. |
| ADSP_GROUPS_PER_PROFILE | Comma-separated URL group names to include. |
| ADSP_GROUP_MAX_TABS | Limit parallel group tabs; 0 opens the full group. |
| IPQS_API_KEY | One key or comma-separated IPQS keys. |
| SITE_TEST_URL | Target for direct site or mock search testing. |
| GOOGLE_SEARCH_QUERY | Query for the optional --google-search mode. |
| MOCK_SEARCH_QUERY / MOCK_SEARCH_EXACT_NAME | Configure the mock search query and exact result name. |
| TWOCAPTCHA_API_KEY | Optional external CAPTCHA service integration already present in the code. |

Environment variables take precedence over matching JSON settings. ipqsStrictness and ipqsMaxScore are configured in JSON; an IPQS score must be lower than ipqsMaxScore to pass. The project does not automatically load .env files.

## Privacy

The .gitignore uses an explicit list of publishable files. API settings, proxy lists, workbooks, browser storage dumps, screenshots, logs, dependencies, and local assistant settings stay on your computer. New source files must be explicitly allowed before Git will include them.

Runtime data can contain sensitive information: .ipqs_daily.json stores exhausted API keys, and console output can include visited URLs and IP addresses. Review logs and screenshots before sharing them. Standalone operation still contacts configured websites and proxy/IP lookup services; optional integrations contact their respective APIs.

## Project files

~~~text
clickbot.js                Main application
ads-config.example.json    Credential-free API configuration
groups.example.json        Example URL group
proxies.example.txt        Placeholder proxy format
package.json               Dependencies and run commands
package-lock.json          Locked dependency versions
PUBLISHING.md              GitHub publishing and privacy checklist
~~~

## Validation

~~~powershell
node --check clickbot.js
~~~

There is currently no automated test suite. Browser behavior depends on the target site, proxy provider, and AdsPower installation.

See [PUBLISHING.md](PUBLISHING.md) for publishing instructions. Choose a license before inviting others to reuse the code; no license grant is included in this repository.
