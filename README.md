# pugh_matrix
The pugh matrix is a decision making application that should help teams to decide

### Pugh Matrix Enhanced Application User Manual

#### 1. **Introduction:**
   Pugh Matrix Enhanced is a Python-based desktop application utilizing the Tkinter library for its graphical user interface. It aids users in decision-making by comparing multiple options based on several criteria through a matrix.

#### 2. **Dependencies:**
   Ensure the following Python libraries are installed. You can install these using pip:

   - **tkinter** (normally comes pre-installed with Python)
   - **openpyxl**
   - **ttkthemes**

   Installation commands:
   ```bash
   pip install openpyxl
   pip install ttkthemes
   ```

#### 3. **User Interface:**

   - **Menu Bar:** Contains 'File' and 'Tools' menus for various operations like creating a new file, importing/exporting Excel files, adding criteria/options, etc.
   - **Criteria and Options Entries:** Users can add criteria and options that are to be compared in the matrix.
   - **Weight Labels:** Shows the weightage of each criterion after the pairwise comparison.
   - **Scales:** Users can adjust the scales for each option under different criteria.

#### 4. **Operations:**

   - **New File:** Clears the current matrix and allows users to start afresh.
   - **Import from Excel:** Users can import data from an existing Excel file to populate the matrix.
   - **Export to Excel:** Allows users to save the current matrix data into an Excel file.
   - **Add Criteria:** Adds new criteria to the matrix.
   - **Add Option:** Adds new options to the matrix.
   - **Pairwise Comparison:** Helps in comparing the pair of criteria to assign weights.
   - **Calculate Score:** Calculates the score for each option based on the set criteria and their weights.

#### 5. **Usage Example:**

   **Step 1:** Open the application. The initial window has menu options for File and Tools.

   **Step 2:** Add criteria using 'Tools' -> 'Add Criteria'. A popup will ask for the number of criteria to add.

   **Step 3:** Similarly, add options using 'Tools' -> 'Add Option'.

   **Step 4:** Perform a pairwise comparison to set weights for the criteria.


   **Step 5:** Adjust the scales for each option under different criteria as per requirements.


   **Step 6:** Calculate the score using 'Tools' -> 'Calculate Score'.


   **Step 7:** (Optional) Export the matrix to an Excel file using 'File' -> 'Export to Excel' for future reference.


#### 6. **Notes:**

   - Ensure that Python and the required libraries are installed before running the application.
   - Follow the on-screen instructions/prompts for seamless operation.
   - The application is best suited for decision-making where multiple options need to be evaluated based on several criteria.

#### 7. **Troubleshooting:**

   If you encounter any issues or errors while using the application:
   - Check if all the required libraries are installed and up-to-date.
   - Ensure the Python script is not corrupted or missing parts of the code.
   - For specific errors, refer to the error message to diagnose and fix the issue, or seek assistance from online Python communities or forums.

