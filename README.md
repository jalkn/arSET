# a r S E T

Django framework to automate audit process.

## 1. Preparation

Execute the main script `set.ps1`. This script installs the dependencies, creates the analysis scripts, and opens the analysis environment in your browser.

```powershell
.\set.ps1
```

## 2. Analysis Execution and Data Visualization

1. Run the script in the terminal:

```
python manage.py runserver
```
2. Click on import data and import your excel file data.

3. To filter the data:

- Use the buttons to add, view, reset, and apply filters.

- Save the filtered results to the downloads folder with the "Save Excel" button.

- By clicking on "details", you can view all data per row and save it to Excel.

## 3. Results

After "Analyze File", the `core/src/` folder will contain the analysis results in Excel files. 