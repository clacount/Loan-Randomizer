# Loan Randomizer

This is a simple local web app for assigning loans to loan officers using fair random distribution.

## How to use
1. Download and unzip the project.
2. Open `index.html` in a current version of Microsoft Edge or Google Chrome.
3. Click **Choose Output Folder** and pick the folder where audit PDFs should be saved.
4. Enter your loan officer names.
5. Enter your loan names or IDs.
6. Click **Randomize Assignments**.

Each run creates a PDF named like `Loan-Randomized-Results-2026-04-14-091530.pdf` in the selected folder. The PDF includes the timestamp, the total loan officers and loans entered, plus both **Assignments by Loan** and **Assignments by Officer** so managers can review every run.

## Rule behavior
- You can use **any number of loan officers**.
- You must have **at least 1 loan officer**.
- Each loan is assigned to **one** officer.
- The app first **shuffles the loans**.
- Then it assigns them in **random fair rounds**:
  - everyone gets one before anyone gets a second
  - everyone gets two before anyone gets a third
  - everyone gets three before anyone gets a fourth
  - and so on

## Examples
- **1 loan, 3 officers**: the single loan goes to one random officer.
- **2 loans, 3 officers**: two different officers get one each.
- **5 loans, 4 officers**: each officer gets one, then one random officer gets the fifth.
- **5 loans, 2 officers**: one officer gets 3 and the other gets 2, at random.
- **7 loans, 3 officers**: each officer gets two before one random officer gets the seventh.

## Notes
- Loan officer names must be unique so totals display correctly.
- Folder selection and direct PDF saving require a browser with the File System Access API.
- No install is needed.
- This runs fully in the browser and can be shared as files.
