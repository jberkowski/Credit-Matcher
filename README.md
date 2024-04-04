# Credit Matcher
v.1.0i
By Jakub Berkowski

## What is Credit Matcher?

Credit Matcher is a VBA macro that enables quick and error-free matching of overpayments and unpaid/shortpaid invoices.

![Credit Matcher main interface screnshot.](https://github.com/jberkowski/Credit-Matcher/assets/93047489/f3b0bba9-d5cf-41ea-a618-f00054efb729)

To work properly, it requires .xlsx export file from GetPaid. Output is a table that lists matched amounts for credits and debits.

## Instructions

Instructions are included in main interface.

1. **Below provide path to save archive copy of used data:** - this option allows for central storage of used GetPaid export files. The archive copy is saved in chosen location at the moment when GetPaid export file is selected and imported.
This ensures that saved copy is unchanged by the collector.

2. **Browse the file using "Browse" button to the right. Duplicate invoices will be removed automatically.** - opens browsing window and allows for choosing export file. If no file is selected, "File path is empty" warning will be shown. Also, all duplicate invoices will be automatically removed after import. Copy of choosen import file will appear in new worksheet "Table".

3. **Remove any credits and/or invoices that you DON'T want to match.** - manually delete all invoices/credits that you don't want to match. At this stage you can freely sort or filter the table. Please note that a lot of manual labor can be simplified by using options in the next step.

4. **Use the options to the right and choose the desired matching settings.** - there are several options that allow for filtering of invoices that should be taken into account when matching:
   - "Use only aged invoices (45+ days)" - only take into account invoices and credits that are 45 or more days past due (Cash Application requirement).
   - "Use only credits below $1,000" - only take into account invoices and credits that amount to less than $1,000 (Cash Application requirement).
   - "Remove Receipts from matching" - removes unapplied payment from matching. You can tell them apart by their "Invoice Number", which in this case starts with "2". This should be unchecked only in special cases.
   - "Remove disputes" - remove all positions that are under dispute. Please note that there is no difference between open and closed dispute in GetPaid export file. This option will remove all of them.
   - "Oldest first" or "Newest first" - prioritize oldest or youngest credits/invoices. These option are mutually exclusive.
