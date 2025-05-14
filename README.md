# Alma_Link_Checker

This PowerShell script checks URLs exported in XLSX format from the Alma library system. It was made by Jai Parker, Information Access Librarian at the Queensland University of Technology with help from Microsoft Copilot. As per the license, caveat emptor.

To run this script:

Export a list of Portfolios for checking from Alma. Ensure you do the Extended Export, so that Row 58 matches the Web Address field the script checks.  The XLSX file name format must end in _portfolios.xlsx for the script to pick it up.

Input
URLs are fed into the script via XLSX files. Files must be named to match the pattern *_portfolios.xlsx (for example Freely_Available_Website_portfolios.xlsx). When it runs, the script will start checking URLs in Row 58 of the first xlsx file it finds with a matching filename pattern.

Ror every URL in Row 58, the script will flag it as broken if:

the hostname in the URL cannot be resolved (i.e. DNS error)
a connection timeout occurs (by default this is 30 seconds)
the connection terminates incorrectly, or
the webserver returns a "400 Bad Request" response
the webserver returns a "404 Not Found" response
the webserver returns a 500 - 599 range response
the webserver returns a 301 OR 308 redirect response AND the URL redirected to is a domain.
Other error values (e.g. "401 Not authorised", "403 Forbidden") are not flagged by this script.

Output
It will produce a CSV file of the broken links named broken-links.csv
This report file contains one column, the MMS ID of the Portfolio containing the broken link.
