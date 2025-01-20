# Monitor OPOST "Create Invoice" Page For Indebted Businesses



## OverView

__A tool that monitors /invoices endpoint on OPOST system at user side__

*once an accountant in opost delivery and logistics company wanted to monitor the indebted businesses so he couldn't skip their invoice payment until the debt was reset*


*that option wasn't supported in opost system itself as the debts are stored manually in a separate Excel spreadsheet*



## How does it work 

__The System invokes selenium web drivers to control the invoices page and check for any entry in the business Name Field__

*the project is deployed into 2 main copies*

__Chrome__ and __Edge__ controls

*the code starts a separate browser session in debugging mode on port 9222 by calling* `def start_chrome_session():` *or* `def start_edge_session():`

