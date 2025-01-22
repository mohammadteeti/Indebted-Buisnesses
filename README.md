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

### API and Connection 

*the code utilizes service account authentication to access Google Spreadsheet on the Google API Project, SpreadSheet API is enabled*
*a __credential.json__ file is used for authentication*

*The spreadsheet on which the operation would be applied must be shared using the Project API Allowed Users (The Granted Email on service account API)*

__credential file is not included here as uploading secrete keys violates Github policy__

## Data Processing 

*Shop names from excel sheet are loaded in `shop_names` list in `def load_indebted_shops_from_sheet` function*<br/>
*`def monitor_shop_input` function is invoked and a new web driver instance is created with the appropriate options of debugging mode*
the driver instance reads the input field specified in <br/> 
```python
        driver.execute_script("""
            const inputField = arguments[0];
            inputField.addEventListener('input', function(event) {
                inputField.setAttribute('data-last-input', inputField.value.trim().toLowerCase());
            });
        """, input_field)
```
For each input inserted by the user a checking flow is invoked if the name is in the shop_name list a beep sound notifies the user 


### Quitting 

*the code ensures the ending of  any process on port 9222 and closes driver sessions*

## Tolerance aspect 
*the code relies on a __config file__ which contains the __debugging mode string__, __URL__ of the page under test, and the __status of the shop__, this configuration could be changed accordingly covering a variety of flows*

