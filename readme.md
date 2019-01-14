Excel as a Scripting Language

Turn excel files into calculators.

# Basics

- Set INPUT_var names
- Set OUTPUT_var names

- Load the formulas that map to the OUTPUT values
- all outputs must be dependent on the input variables and static values
- Show the input variables (and maybe non-input statics) in a UI Form where they can be edited.
- Show the result output values in an output UI form.

- Both Input & Output Forms can be organized in excel if an INPUTS_sheet or OUTPUTS_sheet exists.

# Verification

- For verification, the program can generate spreadsheets that contain the input variables and read the output values calculated in an excel process.
- These generated sheets may be a form of final result also that can be shared like any excel sheets.

# Debugging Formulas

- For debugging, the app can show all calculations that lead up to the output value in a formula breakdown.
- Highlighting a line of the formula can show that value in situ (original sheet context).

# In App Setup

- Alternatively, Inputs and outputs can be selected in app.
  - Select a sheet.
  - Select cells.
  - It could also be possible to select cells by pattern search (all yellow highlighted cells).
  - Those selections become the UI form and inputs.

# Batch Style

- Each set of inputs can be its own row.
- Which generates an output table where each output set is a row.

# As Code

- The Inputs => outputs mapping is a function that can be executed in any code. 
- The formulas can be auto-optimized for native execution.
- They can even generate native code to be compiled if needed.

# As Machine Learning

- In the case where the inputs=>outputs mapping is an approximation function, they could be used as a source for Machine Learning.
- Another mechanism would need to be determined to provide a score on how accurate the output values are.
- If the output values are accurate, but the calculation is simply slow, then it may be possible to learn an approximation calculator with the goal of generating fast-performing native code on a specific target device.
