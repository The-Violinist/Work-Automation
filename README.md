# Work-Automation

## Programs and Scripts to automate daily tasks

This repo contains various mini projects which were inspired by the boredom of repetitive work. These programs have begun to alleviate the boredom by automating certain mundane tasks.

The following programs are the most deveoloped and have been the most useful in my day to day work in compiling security reports.

- [merge_pdfs](https://github.com/The-Violinist/Work-Automation/tree/main/merge_pdfs)
    - By far my favorite! This takes my weekly analysis (docx file) and converts it to PDF then merges it with all of the relevent weekly automatically generated PDF files in the specified order for each client.
- [png_to_pdf](https://github.com/The-Violinist/Work-Automation/tree/main/png_to_pdf)
    - As part of the analysis, there are screenshots from the Microsoft Partner Center which get saved as PNG files. This program searches for them in the weekly folders for each client and converts them to PDF.
- [popular_domains](https://github.com/The-Violinist/Work-Automation/tree/main/popular_domains)
    - Extracts URLs from PDF files and runs them thru nslookup. Any benign URL is added to a safelist and is not presented to the user on subsequent reports.
- [ADreports](https://github.com/The-Violinist/Work-Automation/tree/main/ADreports)
    - Reads xlsx files exported from servers (Inactive users, computers, sslvpn users, admin users). The files are then formatted and converted to PDF.
- [wg_reports](https://github.com/The-Violinist/Work-Automation/tree/main/wg_reports)
    - Parses data extracted from automated WatchGuard reports and returns it in a format which can be cut/pasted into the weekly analysis. Currently, it is not able to handle a number of exceptions very well. It is also a bit sprawling...