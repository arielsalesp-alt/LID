# GitHub Support Request

Copy this text into the GitHub Support portal:

https://support.github.com/contact

## Subject

Please purge cached views and unreachable sensitive data from repository history

## Message

Hello GitHub Support,

I accidentally pushed a CSV file containing real lead/customer contact information to a public repository and GitHub Pages site.

Repository:

```text
arielsalesp-alt/LID
```

GitHub Pages URL:

```text
https://arielsalesp-alt.github.io/LID/
```

Sensitive data type:

```text
Customer/lead personal and business contact information, including phone numbers, business names, email-like fields, and free-text customer notes.
```

What I already did:

```text
1. Replaced the real CSV data with fully fictional demo data.
2. Replaced the embedded CSV data inside index.html with fully fictional demo data.
3. Created a new orphan branch history.
4. Force-pushed the sanitized branch to main.
5. Verified that the live GitHub Pages site and data/sample-leads.csv no longer contain the real data.
```

Current clean commit:

```text
580d79aaeb41b8a0f07ae75b2b84cdfaf6bd9af6
```

Potentially affected old commits that previously existed on main:

```text
7bec15d Fix GitHub Pages standalone app
62654fb Prefer embedded sample CSV on GitHub Pages
8f313d1 Use real lead sample file
```

Affected pull requests:

```text
0 known affected pull requests. I did not create pull requests for this repository.
```

Request:

```text
Please remove cached views and dereference/garbage-collect unreachable Git objects related to the old commits that contained the real CSV/customer data. Please also purge any GitHub Pages cache that may still contain the old CSV or embedded data.
```

Additional note:

```text
No Git LFS was used.
```

Thank you.

