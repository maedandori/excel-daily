# Excel Daily contributor checks

- Preserve compatibility with Windows PowerShell 5.1 and Excel 2019 (merged cells, COM and managed Shapes).
- Use literal replacement callbacks or exact slices when editing PowerShell/YAML via JavaScript; a replacement string containing dollar followed by an apostrophe, ampersand or digit can interpolate unintended source text.
- Re-read changed files, check diff size, function uniqueness, and all modified blocks before pushing.
- For future changes use a feature branch and PR into main. Require successful Windows CI (PowerShell syntax, regression tests, Rust, Ribbon and packaging) before merging to main and creating a release.
- Never skip failed checks to publish a release. Keep the version stable while repairing an unpublished build. Confirm release assets exist before reporting success.
- Test the reported case, both shifts (including midnight), duplicate order IDs, protected/completed rows, and managed Shape movement when relevant. CI without installed Excel is not a substitute for real Excel COM verification.
