# Activity 70 Confirmation (ZP030 and CN47N)

This Python script automates the Activity 70 confirmation process by checking for pending issues in the ZP030 transaction and confirming Activity 70 in the CN47N transaction for work orders without pending issues.

## Description

The script checks if the work orders listed in an Excel file have pending issues in the ZP030 transaction. Work orders without pending issues will have Activity 70 confirmed in the CN47N transaction. A result Excel file is generated with the status of each work order.

## Features

*   **Pending Issue Check (ZP030):** Checks if work orders have pending issues in the ZP030 transaction.
*   **Activity 70 Confirmation (CN47N):** Confirms Activity 70 in the CN47N transaction for work orders without pending issues.
*   **Report Generation:** Generates an Excel file containing the status of each work order in the ZP030 and CN47N transactions.
*   **Error Handling:** Handles errors during the process, such as connection problems with SAP or errors in transactions.

## How to use

1.  **Prerequisites:**
    *   Python 3 installed.
    *   `win32com`, `pandas`, and `traceback` libraries installed (`pip install pywin32 pandas`).
    *   SAP GUI installed and configured.
    *   Access to the ZP030 and CN47N transactions in SAP.
    *   Excel file "Obras CN33 (ZP030 e CN47N).xlsx" filled with the work orders and their respective PEP elements in the first and second columns, respectively, located at: `\\srvfile01\DADOS_PBI\Compartilhado_BI\DPCP\9. DEPM\Bases de Dados\CN33\Macro\Obras (Atividade 70).xlsx`.

2.  **Configuration:**
    *   Define the file paths in the `PATH_ZP030`, `PATH_OBRAS`, and `PATH_RESULTADO` variables at the beginning of the script, if necessary.

3.  **Execution:**
    *   Save and close the Excel file "Obras CN33 (ZP030 e CN47N).xlsx".
    *   Make sure you are logged into SAP.
    *   Run the Python script.
    *   Monitor the progress and errors through the console.
    *   The result file "Obras (Atividade 70)_Resultado.xlsx" will be generated at: `\\srvfile01\DADOS_PBI\Compartilhado_BI\DPCP\9. DEPM\Bases de Dados\CN33\Macro\Obras (Atividade 70)_Resultado.xlsx`.

## Dependencies

*   python >= 3.0
*   pywin32
*   pandas
*   traceback

## Modification History

*   06/20/2024: Work orders with "Stat.mat.espec.cent." status equal to 01 will be passed for confirmation in CN47N.
*   08/06/2024: Work orders with center equal to 5001 will be passed for confirmation in CN47N.

## Author

Gustavo Nunes Ferraz (DPCP)

## Date

04/24/2024

## Additional Information

The script is designed to automate the confirmation of Activity 70, streamlining the process and avoiding manual errors. Be sure to follow the usage and configuration instructions to ensure the proper functioning of the script. If you have any questions or problems, please contact the author.